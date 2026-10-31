"""Opt-in benchmark instrumentation, inherited by Python subprocesses."""
import atexit
import json
import os
from pathlib import Path
import sys
import time

directory = os.environ.get("THESIS_PERF_DIR")
if directory:
    names = {
        "extract_text_evidence": "semanticExtraction", "extract_docx_evidence": "semanticExtraction",
        "parse_document": "docxParsing", "extract_template_baseline": "templateParsing",
        "build_repair_plan": "repairPlanning", "build_source_docx_from_markdown": "docxBuild",
        "convert_docx_to_pdf": "pdfConversion", "export_pages_png": "pageRendering",
        "normalize_docx": "ooxmlWriting",
    }
    started = {}
    metrics = {}
    processes = [0]

    def profile(frame, event, arg):
        name = names.get(frame.f_code.co_name)
        if frame.f_code.co_filename.endswith("apply_ooxml_fixes.py") and frame.f_code.co_name == "main":
            name = "ooxmlWriting"
        if frame.f_code.co_filename.endswith("validate_docx.py") and frame.f_code.co_name == "main":
            name = "finalValidation"
        if name and event == "call":
            started[id(frame)] = time.perf_counter()
        elif name and event == "return":
            start = started.pop(id(frame), None)
            if start is not None:
                item = metrics.setdefault(name, {"seconds": 0, "calls": 0})
                item["seconds"] += time.perf_counter() - start
                item["calls"] += 1

    def audit(event, args):
        if event == "subprocess.Popen":
            processes[0] += 1

    sys.setprofile(profile)
    sys.addaudithook(audit)

    @atexit.register
    def save():
        sys.setprofile(None)
        Path(directory).mkdir(parents=True, exist_ok=True)
        Path(directory, f"{os.getpid()}.json").write_text(json.dumps({"stages": metrics, "externalProcesses": processes[0]}), encoding="utf-8")
