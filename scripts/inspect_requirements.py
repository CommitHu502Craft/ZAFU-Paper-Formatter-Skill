#!/usr/bin/env python3
"""Read-only inventory before choosing the paper formatting route."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx_ooxml import load_rules, parse_document
from extract_source_evidence import extract_docx_evidence
from thesis_ir import build_thesis_ir
from formatter_core.sections import inspect_requirements


def main():
    parser = argparse.ArgumentParser(description="列出论文缺项，不修改原稿或模板。")
    parser.add_argument("input", type=Path)
    parser.add_argument("--rules-yaml", type=Path, default=Path(__file__).resolve().parents[1] / "references" / "zafu_2022_rules.yaml")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.output and args.output.resolve() in {args.input.resolve(), args.rules_yaml.resolve()}:
        raise ValueError("Inventory output must not overwrite the source or rules")
    audit = parse_document(str(args.input), rules_path=str(args.rules_yaml), detail="semantic")
    ir = build_thesis_ir(extract_docx_evidence(args.input, None, str(args.rules_yaml), audit=audit))
    result = inspect_requirements(audit, load_rules(str(args.rules_yaml)), ir)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
