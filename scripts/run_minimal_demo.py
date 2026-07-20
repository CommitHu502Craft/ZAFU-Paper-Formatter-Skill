#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from docx_ooxml import build_repair_plan, extract_template_baseline, parse_document, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the minimal preflight/baseline/audit/plan/apply/validate thesis demo.")
    parser.add_argument("docx", help="Input DOCX")
    parser.add_argument("--template-docx", required=True, help="Template DOCX")
    parser.add_argument("--rules-yaml", required=True, help="Rules YAML")
    parser.add_argument("--output-dir", required=True, help="Directory to write demo outputs")
    parser.add_argument("--skip-apply", action="store_true", help="Only write baseline/audit/plan outputs")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    baseline = extract_template_baseline(args.template_docx, rules_path=args.rules_yaml)
    write_json(baseline, str(output_dir / "template_baseline.json"))

    subprocess.run(
        [
            sys.executable,
            str(Path(__file__).with_name("preflight_semantic_normalization.py")),
            args.docx,
            "--template-docx",
            args.template_docx,
            "--rules-yaml",
            args.rules_yaml,
            "--output",
            str(output_dir / "preflight_report.json"),
        ],
        check=True,
    )

    audit = parse_document(args.docx, template_docx=args.template_docx, rules_path=args.rules_yaml)
    write_json(audit, str(output_dir / "audit_report.json"))

    plan = build_repair_plan(audit)
    write_json(plan, str(output_dir / "repair_plan.json"))

    if args.skip_apply:
        return

    repaired_docx = output_dir / "repaired.docx"
    diff_report = output_dir / "repair_execution.json"
    validate_report = output_dir / "validation_report.json"

    subprocess.run(
        [
            sys.executable,
            str(Path(__file__).with_name("apply_ooxml_fixes.py")),
            args.docx,
            str(repaired_docx),
            "--plan-json",
            str(output_dir / "repair_plan.json"),
            "--template-docx",
            args.template_docx,
            "--rules-yaml",
            args.rules_yaml,
            "--report-json",
            str(diff_report),
        ],
        check=True,
    )

    subprocess.run(
        [
            sys.executable,
            str(Path(__file__).with_name("validate_docx.py")),
            str(repaired_docx),
            "--before-docx",
            args.docx,
            "--template-docx",
            args.template_docx,
            "--rules-yaml",
            args.rules_yaml,
            "--plan-json",
            str(output_dir / "repair_plan.json"),
            "--output",
            str(validate_report),
        ],
        check=True,
    )


if __name__ == "__main__":
    main()
