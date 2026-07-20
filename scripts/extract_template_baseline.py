#!/usr/bin/env python3
import argparse

from docx_ooxml import extract_template_baseline, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract a deep OOXML baseline from a thesis template DOCX.")
    parser.add_argument("docx", help="Template DOCX path")
    parser.add_argument("--rules-yaml", help="Optional rules YAML for baseline/rules comparison")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    args = parser.parse_args()

    baseline = extract_template_baseline(args.docx, rules_path=args.rules_yaml)
    write_json(baseline, args.output)


if __name__ == "__main__":
    main()
