#!/usr/bin/env python3
import argparse

from docx_ooxml import parse_document, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Emit paragraph/run effective formatting for a DOCX.")
    parser.add_argument("docx", help="DOCX file to inspect")
    parser.add_argument("--rules-yaml", help="Optional rules YAML")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    args = parser.parse_args()

    audit = parse_document(args.docx, rules_path=args.rules_yaml)
    result = {
        "generatedAt": audit["generatedAt"],
        "source": audit["source"],
        "docDefaults": audit["docDefaults"],
        "theme": audit["theme"],
        "paragraphs": [
            {
                "index": paragraph["index"],
                "text": paragraph["text"],
                "styleId": paragraph["styleId"],
                "styleName": paragraph["styleName"],
                "manualNumbering": paragraph["manualNumbering"],
                "effectiveParagraph": paragraph["effectiveParagraph"],
                "effectiveRunSummary": paragraph["effectiveRunSummary"],
                "runs": [
                    {
                        "index": run["index"],
                        "text": run["text"],
                        "styleId": run["styleId"],
                        "styleName": run["styleName"],
                        "effectiveFormat": run["effectiveFormat"],
                    }
                    for run in paragraph["runs"]
                ],
            }
            for paragraph in audit["paragraphs"]
        ],
    }
    write_json(result, args.output)


if __name__ == "__main__":
    main()
