#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional


def detect_libreoffice() -> Optional[str]:
    return shutil.which("soffice") or shutil.which("libreoffice")


def detect_pdftoppm() -> Optional[str]:
    return shutil.which("pdftoppm")


def count_pdf_pages(pdf_path: Path) -> Optional[int]:
    if not pdf_path.exists():
        return None
    data = pdf_path.read_bytes()
    matches = re.findall(rb"/Type\s*/Page\b", data)
    return len(matches) or None


def suggested_review_pages(page_count: Optional[int]) -> Dict[str, Optional[int]]:
    if not page_count or page_count <= 0:
        return {
            "cover": None,
            "toc": None,
            "abstract": None,
            "firstBodyPage": None,
            "references": None,
        }
    return {
        "cover": 1,
        "toc": min(2, page_count),
        "abstract": min(3, page_count),
        "firstBodyPage": min(4, page_count),
        "references": page_count,
    }


def write_preview_manifest(preview_dir: Path, entries: List[Dict[str, Any]]) -> Path:
    preview_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = preview_dir / "preview_manifest.json"
    manifest_path.write_text(json.dumps({"pages": entries}, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path


def planned_preview_entries(suggested_pages: Dict[str, Optional[int]]) -> List[Dict[str, Any]]:
    entries: List[Dict[str, Any]] = []
    seen_pages = set()
    for label, page in suggested_pages.items():
        if page is None:
            continue
        key = (label, page)
        if key in seen_pages:
            continue
        seen_pages.add(key)
        entries.append(
            {
                "label": label,
                "page": page,
                "status": "planned",
                "imagePath": None,
            }
        )
    return entries


def export_pdf_page_preview(pdftoppm_path: str, pdf_path: Path, page_number: int, output_base: Path, timeout_sec: int) -> Optional[Path]:
    command = [
        pdftoppm_path,
        "-png",
        "-f",
        str(page_number),
        "-l",
        str(page_number),
        str(pdf_path),
        str(output_base),
    ]
    completed = subprocess.run(command, check=False, capture_output=True, text=True, timeout=timeout_sec)
    if completed.returncode != 0:
        return None
    candidates = sorted(output_base.parent.glob(f"{output_base.name}*.png"))
    return candidates[0] if candidates else None


def export_preview_images(
    pdf_path: Path,
    preview_dir: Path,
    suggested_pages: Dict[str, Optional[int]],
    timeout_sec: int,
) -> Dict[str, Any]:
    pdftoppm_path = detect_pdftoppm()
    entries = planned_preview_entries(suggested_pages)
    preview_manifest_path = write_preview_manifest(preview_dir, entries)
    if not pdftoppm_path:
        return {
            "available": False,
            "backend": None,
            "backendPath": None,
            "manifestPath": str(preview_manifest_path),
            "images": entries,
            "reason": "pdftoppm not found in PATH",
        }

    rendered_entries: List[Dict[str, Any]] = []
    for entry in entries:
        page_number = entry.get("page")
        label = str(entry.get("label") or "page")
        if not isinstance(page_number, int) or page_number <= 0:
            rendered_entries.append(entry)
            continue
        output_base = preview_dir / f"{label}_page_{page_number:03d}"
        image_path = export_pdf_page_preview(pdftoppm_path, pdf_path, page_number, output_base, timeout_sec)
        rendered_entry = dict(entry)
        if image_path is not None and image_path.exists():
            rendered_entry["status"] = "exported"
            rendered_entry["imagePath"] = str(image_path)
        else:
            rendered_entry["status"] = "failed"
            rendered_entry["imagePath"] = None
        rendered_entries.append(rendered_entry)

    preview_manifest_path = write_preview_manifest(preview_dir, rendered_entries)
    return {
        "available": True,
        "backend": "pdftoppm",
        "backendPath": pdftoppm_path,
        "manifestPath": str(preview_manifest_path),
        "images": rendered_entries,
        "reason": None,
    }


def run_render_validation(docx_path: str, output_dir: str, timeout_sec: int = 120) -> Dict[str, Any]:
    source = Path(docx_path).resolve()
    render_dir = Path(output_dir).resolve()
    render_dir.mkdir(parents=True, exist_ok=True)
    preview_dir = render_dir / "before_after_page_preview"
    preview_dir.mkdir(parents=True, exist_ok=True)

    backend = detect_libreoffice()
    report: Dict[str, Any] = {
        "sourceDocx": str(source),
        "outputDir": str(render_dir),
        "available": bool(backend),
        "backend": "libreoffice" if backend else None,
        "backendPath": backend,
        "status": "unavailable",
        "reason": None,
        "pdfPath": None,
        "pageCount": None,
        "previewDir": str(preview_dir),
        "suggestedReviewPages": suggested_review_pages(None),
        "previewExport": {
            "available": False,
            "backend": None,
            "backendPath": None,
            "manifestPath": None,
            "images": [],
            "reason": "PDF not available yet",
        },
    }
    if not backend:
        report["reason"] = "LibreOffice/soffice not found in PATH"
        preview_manifest = write_preview_manifest(preview_dir, planned_preview_entries(report["suggestedReviewPages"]))
        report["previewExport"] = {
            "available": False,
            "backend": None,
            "backendPath": None,
            "manifestPath": str(preview_manifest),
            "images": planned_preview_entries(report["suggestedReviewPages"]),
            "reason": "PDF preview export skipped because LibreOffice/soffice not found in PATH",
        }
        return report

    command = [
        backend,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(render_dir),
        str(source),
    ]
    completed = subprocess.run(command, check=False, capture_output=True, text=True, timeout=timeout_sec)
    pdf_path = render_dir / f"{source.stem}.pdf"
    report["command"] = command
    report["stdout"] = completed.stdout.strip()
    report["stderr"] = completed.stderr.strip()
    report["pdfPath"] = str(pdf_path) if pdf_path.exists() else None

    if completed.returncode != 0:
        report["status"] = "failed"
        report["reason"] = f"LibreOffice exited with code {completed.returncode}"
        return report
    if not pdf_path.exists():
        report["status"] = "failed"
        report["reason"] = "PDF conversion command completed but expected output PDF was not found"
        return report

    page_count = count_pdf_pages(pdf_path)
    report["pageCount"] = page_count
    report["suggestedReviewPages"] = suggested_review_pages(page_count)
    report["previewExport"] = export_preview_images(pdf_path, preview_dir, report["suggestedReviewPages"], timeout_sec)
    report["status"] = "ok"
    report["reason"] = None
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Run minimal DOCX render validation through LibreOffice if available.")
    parser.add_argument("docx", help="DOCX file to render")
    parser.add_argument("--output-dir", required=True, help="Directory for PDF and render artifacts")
    parser.add_argument("--output", "-o", help="Write JSON report to this path")
    parser.add_argument("--timeout-sec", type=int, default=120, help="Renderer timeout in seconds")
    args = parser.parse_args()

    report = run_render_validation(args.docx, args.output_dir, timeout_sec=args.timeout_sec)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
