#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

from docx_ooxml import build_repair_plan, parse_document, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a repair plan from deep DOCX audit results.")
    parser.add_argument("docx", help="DOCX file to inspect")
    parser.add_argument("--template-docx", help="Template DOCX used as physical baseline")
    parser.add_argument("--rules-yaml", help="Rules YAML used for defaults and cross-checking")
    parser.add_argument("--style-map-yaml", help="Optional profile style-role mapping YAML")
    parser.add_argument("--front-matter-policy-yaml", help="Optional profile front-matter policy YAML")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    parser.add_argument("--thesis-ir-json", help="IR for this original DOCX; applies explicit role corrections only")
    parser.add_argument("--allow-structural-rebuild", action="store_true", help="Allow template front-matter replacement")
    args = parser.parse_args()

    audit = parse_document(args.docx, template_docx=args.template_docx, rules_path=args.rules_yaml)
    plan = build_repair_plan(
        audit,
        style_map_path=args.style_map_yaml,
        front_matter_policy_path=args.front_matter_policy_yaml,
        allow_structural_rebuild=args.allow_structural_rebuild,
    )
    if args.thesis_ir_json:
        from formatter_core.word_semantics import apply_ir_corrections

        apply_ir_corrections(plan, audit, json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8")))
    write_json(plan, args.output)


if __name__ == "__main__":
    main()
