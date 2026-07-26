#!/usr/bin/env python3
"""DOCX -> PDF conversion backends.

Order of preference:
1. LibreOffice (soffice) headless conversion.
2. Microsoft Word COM automation via PowerShell (Windows only, optional).
3. WPS Office COM automation via PowerShell (Windows only, optional).

All backends are optional; callers must tolerate a None result and keep
producing the DOCX deliverable without a PDF.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional


def detect_libreoffice() -> Optional[str]:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    if sys.platform == "win32":
        for candidate in (
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "LibreOffice" / "program" / "soffice.exe",
        ):
            if candidate.exists():
                return str(candidate)
    return None


def _run_powershell_script(script_body: str, timeout_sec: int) -> subprocess.CompletedProcess:
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as handle:
        handle.write(script_body)
        script_path = handle.name
    try:
        return subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path],
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout_sec,
        )
    finally:
        try:
            os.unlink(script_path)
        except OSError:
            pass


_COM_SCRIPT_TEMPLATE = """
$ErrorActionPreference = 'Stop'
try {{
    $app = New-Object -ComObject {com_progid}
    $app.Visible = $false
}} catch {{
    Write-Output 'BACKEND_UNAVAILABLE'
    exit 3
}}
try {{
    $doc = $app.Documents.Open('{docx}', $false, $true)
    $doc.ExportAsFixedFormat('{pdf}', 17)
    $doc.Close(0)
    Write-Output 'CONVERT_OK'
}} catch {{
    Write-Output ('CONVERT_FAIL: ' + $_.Exception.Message)
    exit 4
}} finally {{
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}}
"""


def _convert_via_com(progid: str, docx_path: Path, pdf_path: Path, timeout_sec: int) -> Dict[str, Any]:
    def ps_quote(value: str) -> str:
        return value.replace("'", "''")

    script = _COM_SCRIPT_TEMPLATE.format(
        com_progid=progid,
        docx=ps_quote(str(docx_path)),
        pdf=ps_quote(str(pdf_path)),
    )
    try:
        completed = _run_powershell_script(script, timeout_sec)
    except subprocess.TimeoutExpired:
        return {"ok": False, "reason": f"{progid} conversion timed out"}
    stdout = (completed.stdout or "").strip()
    if "CONVERT_OK" in stdout and pdf_path.exists():
        return {"ok": True, "reason": None}
    if "BACKEND_UNAVAILABLE" in stdout:
        return {"ok": False, "reason": f"{progid} COM object not available", "unavailable": True}
    return {"ok": False, "reason": stdout or (completed.stderr or "").strip() or f"{progid} conversion failed"}


def convert_docx_to_pdf(docx_path: Path, output_dir: Path, timeout_sec: int = 180) -> Dict[str, Any]:
    """Try all available backends; return a report describing what happened."""
    docx_path = Path(docx_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = output_dir / f"{docx_path.stem}.pdf"
    attempts = []

    soffice = detect_libreoffice()
    if soffice:
        command = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(docx_path)]
        try:
            completed = subprocess.run(command, check=False, capture_output=True, text=True, timeout=timeout_sec)
            if completed.returncode == 0 and pdf_path.exists():
                return {
                    "available": True,
                    "backend": "libreoffice",
                    "backendPath": soffice,
                    "pdfPath": str(pdf_path),
                    "status": "ok",
                    "reason": None,
                    "attempts": attempts,
                }
            attempts.append({"backend": "libreoffice", "reason": f"exit code {completed.returncode}: {(completed.stderr or '').strip()[:400]}"})
        except subprocess.TimeoutExpired:
            attempts.append({"backend": "libreoffice", "reason": "conversion timed out"})
    else:
        attempts.append({"backend": "libreoffice", "reason": "soffice not found in PATH"})

    if sys.platform == "win32":
        for backend_name, progid in (("word_com", "Word.Application"), ("wps_com", "KWPS.Application")):
            result = _convert_via_com(progid, docx_path, pdf_path, timeout_sec)
            if result.get("ok"):
                return {
                    "available": True,
                    "backend": backend_name,
                    "backendPath": progid,
                    "pdfPath": str(pdf_path),
                    "status": "ok",
                    "reason": None,
                    "attempts": attempts,
                }
            attempts.append({"backend": backend_name, "reason": result.get("reason")})
            if not result.get("unavailable"):
                # Backend exists but conversion failed; still try the next one.
                continue

    return {
        "available": False,
        "backend": None,
        "backendPath": None,
        "pdfPath": None,
        "status": "unavailable",
        "reason": "No DOCX->PDF backend succeeded (tried LibreOffice, Word COM, WPS COM)",
        "attempts": attempts,
    }
