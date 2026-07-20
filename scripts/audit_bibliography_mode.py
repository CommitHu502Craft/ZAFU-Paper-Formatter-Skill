#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx_ooxml import load_rules, parse_document, write_json


NUMERIC_CITATION_RE = re.compile(r"\[\s*\d+(?:\s*[-,，]\s*\d+)*\s*\]")
AUTHOR_YEAR_CITATION_RE = re.compile(
    r"[（(]\s*(?:[A-Z][A-Za-z .'\-]+(?:\s+et al\.)?|[\u4e00-\u9fff]{1,6}(?:等)?)\s*[，,]\s*\d{4}[a-z]?(?:\s*[；;]\s*(?:[A-Z][A-Za-z .'\-]+(?:\s+et al\.)?|[\u4e00-\u9fff]{1,6}(?:等)?)\s*[，,]\s*\d{4}[a-z]?)*\s*[）)]"
)
NUMERIC_BIB_ENTRY_RE = re.compile(r"^\[\s*\d+\s*\]")
FOREIGN_ENTRY_RE = re.compile(r"^[A-Z][A-Za-z' \-]+[,，]")
STOP_HEADINGS = {"致谢", "致  谢", "附录", "附 录"}


def split_references_from_docx(audit: Dict[str, Any]) -> Dict[str, Any]:
    paragraphs = audit.get("paragraphs") or []
    body_paragraphs = []
    bibliography_entries = []
    in_references = False
    references_heading_index = None

    for paragraph in paragraphs:
        text = (paragraph.get("text") or "").strip()
        role = (paragraph.get("role") or {}).get("role")
        if role == "references_heading" or text == "参考文献":
            in_references = True
            references_heading_index = paragraph.get("index")
            continue
        if in_references and (role in {"ack_heading", "appendix_heading"} or text in STOP_HEADINGS):
            in_references = False
        if in_references:
            if text:
                bibliography_entries.append({"index": paragraph.get("index"), "text": text})
        else:
            if text:
                body_paragraphs.append({"index": paragraph.get("index"), "text": text})

    return {
        "referencesHeadingIndex": references_heading_index,
        "bodyParagraphs": body_paragraphs,
        "bibliographyEntries": bibliography_entries,
    }


def split_references_from_text(path: Path) -> Dict[str, Any]:
    lines = [line.strip() for line in path.read_text(encoding="utf-8").splitlines()]
    body_paragraphs = []
    bibliography_entries = []
    in_references = False
    references_heading_index = None

    for index, line in enumerate(lines):
        if not line:
            continue
        if line == "参考文献":
            in_references = True
            references_heading_index = index
            continue
        if in_references and line in STOP_HEADINGS:
            in_references = False
        if in_references:
            bibliography_entries.append({"index": index, "text": line})
        else:
            body_paragraphs.append({"index": index, "text": line})

    return {
        "referencesHeadingIndex": references_heading_index,
        "bodyParagraphs": body_paragraphs,
        "bibliographyEntries": bibliography_entries,
    }


def detect_language_bucket(text: str) -> str:
    return "zh" if re.search(r"[\u4e00-\u9fff]", text or "") else "en"


def audit_blocks(
    body_paragraphs: List[Dict[str, Any]],
    bibliography_entries: List[Dict[str, Any]],
    rules: Dict[str, Any],
) -> Dict[str, Any]:
    numeric_citations = []
    author_year_citations = []
    numeric_bibliography = []
    zh_entries = []
    en_entries = []

    for paragraph in body_paragraphs:
        text = paragraph["text"]
        if NUMERIC_CITATION_RE.search(text):
            numeric_citations.append({"index": paragraph["index"], "text": text[:200]})
        if AUTHOR_YEAR_CITATION_RE.search(text):
            author_year_citations.append({"index": paragraph["index"], "text": text[:200]})

    for entry in bibliography_entries:
        text = entry["text"]
        if NUMERIC_BIB_ENTRY_RE.match(text):
            numeric_bibliography.append({"index": entry["index"], "text": text[:200]})
        if detect_language_bucket(text) == "zh":
            zh_entries.append(entry)
        else:
            en_entries.append(entry)

    configured_mode = ((rules.get("references") or {}).get("citation_mode")) or None
    numbering_mode = ((rules.get("references") or {}).get("bibliography_numbering")) or None
    inferred_mode = "author_year" if len(author_year_citations) >= len(numeric_citations) else "numeric"
    issues = []

    if configured_mode == "author_year" and numeric_citations:
        issues.append(
            {
                "kind": "numeric_citation_found_under_author_year_mode",
                "count": len(numeric_citations),
                "examples": numeric_citations[:20],
            }
        )
    if numbering_mode == "none" and numeric_bibliography:
        issues.append(
            {
                "kind": "numeric_bibliography_entry_found_under_unnumbered_mode",
                "count": len(numeric_bibliography),
                "examples": numeric_bibliography[:20],
            }
        )
    if zh_entries and en_entries:
        first_en = min(item["index"] for item in en_entries)
        last_zh = max(item["index"] for item in zh_entries)
        if first_en < last_zh:
            issues.append(
                {
                    "kind": "bibliography_language_group_order_mismatch",
                    "expectedOrder": ["zh", "en"],
                    "details": {"firstEnglishIndex": first_en, "lastChineseIndex": last_zh},
                }
            )

    return {
        "configuredMode": configured_mode,
        "inferredMode": inferred_mode,
        "bibliographyNumbering": numbering_mode,
        "bodyCitationStats": {
            "authorYearCount": len(author_year_citations),
            "numericCount": len(numeric_citations),
        },
        "bibliographyStats": {
            "entryCount": len(bibliography_entries),
            "numericLabelCount": len(numeric_bibliography),
            "zhCount": len(zh_entries),
            "enCount": len(en_entries),
        },
        "issues": issues,
        "examples": {
            "authorYearCitations": author_year_citations[:20],
            "numericCitations": numeric_citations[:20],
            "numericBibliography": numeric_bibliography[:20],
            "bibliographyHead": bibliography_entries[:20],
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit bibliography mode against Zhejiang A&F author-year expectations.")
    parser.add_argument("source", help="Input .docx, .md, or .txt")
    parser.add_argument("--template-docx", help="Optional template for DOCX audit")
    parser.add_argument("--rules-yaml", default="references/zafu_2022_rules.yaml", help="Rules YAML")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    source = Path(args.source)
    rules = load_rules(args.rules_yaml)

    if source.suffix.lower() == ".docx":
        audit = parse_document(str(source), template_docx=args.template_docx, rules_path=args.rules_yaml)
        split = split_references_from_docx(audit)
        result = {
            "source": str(source.resolve()),
            "sourceKind": "docx",
            "referencesHeadingIndex": split["referencesHeadingIndex"],
            **audit_blocks(split["bodyParagraphs"], split["bibliographyEntries"], rules),
        }
    elif source.suffix.lower() in {".md", ".txt", ".markdown"}:
        split = split_references_from_text(source)
        result = {
            "source": str(source.resolve()),
            "sourceKind": source.suffix.lower().lstrip("."),
            "referencesHeadingIndex": split["referencesHeadingIndex"],
            **audit_blocks(split["bodyParagraphs"], split["bibliographyEntries"], rules),
        }
    else:
        raise SystemExit(f"Unsupported source type: {source.suffix}")

    write_json(result, args.output)


if __name__ == "__main__":
    main()
