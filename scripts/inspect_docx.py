#!/usr/bin/env python3
import argparse

from docx_ooxml import parse_document, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit a DOCX thesis with deep OOXML inspection.")
    parser.add_argument("docx", help="DOCX file to inspect")
    parser.add_argument("--template-docx", help="Optional template DOCX used as the physical baseline")
    parser.add_argument("--rules-yaml", help="Optional rules YAML used for recommended defaults and comparison")
    parser.add_argument("--style-map-yaml", help="Optional profile style-role mapping YAML")
    parser.add_argument("--front-matter-policy-yaml", help="Optional profile front-matter policy YAML")
    parser.add_argument("--validators-yaml", help="Optional profile validator selection YAML")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    args = parser.parse_args()

    audit = parse_document(args.docx, template_docx=args.template_docx, rules_path=args.rules_yaml)
    write_json(audit, args.output)


if __name__ == "__main__":
    main()
