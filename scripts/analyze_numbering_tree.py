#!/usr/bin/env python3
import argparse

from docx_ooxml import parse_document, write_json


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze heading tree, numbering families, and repair candidates.")
    parser.add_argument("docx", help="DOCX file to inspect")
    parser.add_argument("--rules-yaml", help="Optional rules YAML")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    args = parser.parse_args()

    audit = parse_document(args.docx, rules_path=args.rules_yaml)
    result = {
        "generatedAt": audit["generatedAt"],
        "source": audit["source"],
        "numberingAnalysis": audit["numberingAnalysis"],
        "headingParagraphs": [
            {
                "index": paragraph["index"],
                "text": paragraph["text"],
                "styleId": paragraph["styleId"],
                "styleName": paragraph["styleName"],
                "numId": paragraph["numId"],
                "ilvl": paragraph["ilvl"],
                "manualNumbering": paragraph["manualNumbering"],
                "role": paragraph["role"],
            }
            for paragraph in audit["paragraphs"]
            if (paragraph.get("role") or {}).get("role", "").startswith("heading")
        ],
    }
    write_json(result, args.output)


if __name__ == "__main__":
    main()
