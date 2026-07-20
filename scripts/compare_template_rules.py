#!/usr/bin/env python3
import argparse

from docx_ooxml import compare_rules_to_baseline, extract_template_baseline, load_rules, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare template baseline against thesis rules YAML.")
    parser.add_argument("template_docx", help="Template DOCX")
    parser.add_argument("rules_yaml", help="Rules YAML")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    args = parser.parse_args()

    baseline = extract_template_baseline(args.template_docx, rules_path=args.rules_yaml)
    result = {
        "generatedAt": baseline["generatedAt"],
        "templateDocx": baseline["source"],
        "rulesYaml": args.rules_yaml,
        "comparison": compare_rules_to_baseline(load_rules(args.rules_yaml), baseline),
    }
    write_json(result, args.output)


if __name__ == "__main__":
    main()
