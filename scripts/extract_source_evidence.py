#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx_ooxml import parse_document


ABSTRACT_LABEL_RE = re.compile(r"^(摘\s*要|Abstract|ABSTRACT)[:：]?", re.IGNORECASE)
KEYWORDS_LABEL_RE = re.compile(r"^(关键词|Key\s*words?|Keywords)[:：]?", re.IGNORECASE)
REFERENCES_RE = re.compile(r"^(参考文献|References)$", re.IGNORECASE)
ACKNOWLEDGEMENTS_RE = re.compile(r"^(致谢|致\s*谢|Acknowledg(?:e)?ments?)$", re.IGNORECASE)
APPENDIX_RE = re.compile(r"^(附录(?:[A-ZＡ-Ｚ0-9一二三四五六七八九十]*)?.*|Appendix(?:\s+[A-Z0-9]+)?(?:\s*[:：.-]\s*.*)?)$", re.IGNORECASE)
CAPTION_RE = re.compile(
    r"^(?:(图|表)\s*[0-9一二三四五六七八九十]+(?:[.．\-—][0-9一二三四五六七八九十]+)?|"
    r"(Figure|Fig\.?|Table)\s*[0-9A-Za-z]+(?:[.．\-—][0-9A-Za-z]+)?)\s*[:：.．]?\s*.+$",
    re.IGNORECASE,
)
HEADING_HINT_RE = re.compile(
    r"^(?:\d+(?:\.\d+)*\s+.+|第[一二三四五六七八九十百千]+[章节].*|[一二三四五六七八九十百千]+[、.．]\s*.+|[①②③④⑤⑥⑦⑧⑨⑩].+)$"
)
MARKDOWN_HEADING_RE = re.compile(r"^(?P<marks>#{1,6})\s+(?P<text>.+?)\s*$")
MARKDOWN_IMAGE_RE = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)]+)\)\s*$")


def canonical_source_format(source: Path) -> str:
    suffix = source.suffix.lower()
    if suffix == ".docx":
        return "docx"
    if suffix in {".md", ".markdown"}:
        return "markdown"
    return "text"


def markdown_table_separator(line: str) -> bool:
    raw = line.strip()
    if not raw.startswith("|"):
        return False
    cells = [cell.strip() for cell in raw.strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell or "") for cell in cells)


def extract_markdown_source_blocks(lines: List[str]) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    paragraph_lines: List[str] = []
    paragraph_start: Optional[int] = None

    def flush_paragraph() -> None:
        nonlocal paragraph_start
        text = "\n".join(paragraph_lines).strip()
        if text:
            blocks.append(
                {
                    "kind": "paragraph",
                    "sourceIndex": paragraph_start,
                    "text": text,
                    "sourceAnchor": {"kind": "line", "index": paragraph_start},
                }
            )
        paragraph_lines.clear()
        paragraph_start = None

    index = 0
    while index < len(lines):
        raw = lines[index]
        stripped = raw.strip()
        heading = MARKDOWN_HEADING_RE.match(stripped)
        image = MARKDOWN_IMAGE_RE.match(stripped)
        if not stripped:
            flush_paragraph()
            index += 1
            continue
        if heading:
            flush_paragraph()
            blocks.append(
                {
                    "kind": "heading",
                    "sourceIndex": index,
                    "text": heading.group("text").strip(),
                    "level": len(heading.group("marks")),
                    "sourceAnchor": {"kind": "line", "index": index},
                }
            )
            index += 1
            continue
        if image:
            flush_paragraph()
            blocks.append(
                {
                    "kind": "image",
                    "sourceIndex": index,
                    "text": image.group("alt").strip(),
                    "alt": image.group("alt").strip(),
                    "path": image.group("path").strip(),
                    "sourceAnchor": {"kind": "line", "index": index},
                }
            )
            index += 1
            continue
        if stripped.startswith("|") and index + 1 < len(lines) and markdown_table_separator(lines[index + 1]):
            flush_paragraph()
            start = index
            table_lines = [raw, lines[index + 1]]
            index += 2
            while index < len(lines) and lines[index].strip().startswith("|"):
                table_lines.append(lines[index])
                index += 1
            blocks.append(
                {
                    "kind": "table",
                    "sourceIndex": start,
                    "lines": table_lines,
                    "sourceAnchor": {"kind": "line_range", "start": start, "endExclusive": index},
                }
            )
            continue
        if paragraph_start is None:
            paragraph_start = index
        paragraph_lines.append(raw)
        index += 1
    flush_paragraph()
    return blocks


def extract_plain_text_source_blocks(line_evidence: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    for item in line_evidence:
        if item.get("isBlank"):
            continue
        kind = "heading" if item.get("candidateHeading") or any(
            (item.get("candidateLabels") or {}).get(key)
            for key in ("references", "acknowledgements", "appendix")
        ) else "paragraph"
        block = {
            "kind": kind,
            "sourceIndex": item.get("index"),
            "text": item.get("text"),
            "sourceAnchor": {"kind": "line", "index": item.get("index")},
        }
        if kind == "heading":
            block["level"] = 1
        blocks.append(block)
    return blocks


def make_line_entry(index: int, text: str) -> Dict[str, Any]:
    stripped = (text or "").strip()
    markdown_heading = MARKDOWN_HEADING_RE.match(stripped)
    semantic_stripped = markdown_heading.group("text").strip() if markdown_heading else stripped
    semantic_stripped = re.sub(r"^\*\*(.+?)\*\*", r"\1", semantic_stripped).strip()
    return {
        "index": index,
        "text": text,
        "isBlank": not stripped,
        "blankLineBoundaryBefore": False,
        "candidateLabels": {
            "abstract": bool(ABSTRACT_LABEL_RE.match(semantic_stripped)),
            "keywords": bool(KEYWORDS_LABEL_RE.match(semantic_stripped)),
            "references": bool(REFERENCES_RE.match(semantic_stripped)),
            "acknowledgements": bool(ACKNOWLEDGEMENTS_RE.match(semantic_stripped)),
            "appendix": bool(APPENDIX_RE.match(semantic_stripped)),
            "caption": bool(CAPTION_RE.match(semantic_stripped)),
        },
        "candidateHeading": bool(HEADING_HINT_RE.match(stripped) or markdown_heading),
        "markdownHeadingLevel": len(markdown_heading.group("marks")) if markdown_heading else None,
        "normalizedText": semantic_stripped,
    }


def extract_text_evidence(source: Path) -> Dict[str, Any]:
    lines = source.read_text(encoding="utf-8").splitlines()
    line_evidence = [make_line_entry(index, line) for index, line in enumerate(lines)]
    previous_blank = True
    for item in line_evidence:
        item["blankLineBoundaryBefore"] = previous_blank
        previous_blank = bool(item["isBlank"])
    candidate_blocks = [item for item in line_evidence if item["candidateHeading"] or any(item["candidateLabels"].values())]
    source_format = canonical_source_format(source)
    source_blocks = (
        extract_markdown_source_blocks(lines)
        if source_format == "markdown"
        else extract_plain_text_source_blocks(line_evidence)
    )
    return {
        "source": str(source.resolve()),
        "sourceType": source.suffix.lower().lstrip(".") or "txt",
        "sourceFormat": source_format,
        "evidenceVersion": "2.0",
        "lineCount": len(lines),
        "lineEvidence": line_evidence,
        "sourceBlocks": source_blocks,
        "candidateBlocks": candidate_blocks[:120],
        "preserveableAssets": {"tables": 0, "images": 0, "equations": 0, "sections": 0},
        "ambiguities": [],
    }


def extract_docx_evidence(source: Path, template_docx: Optional[str], rules_yaml: Optional[str]) -> Dict[str, Any]:
    audit = parse_document(str(source), template_docx=template_docx, rules_path=rules_yaml)
    paragraph_evidence: List[Dict[str, Any]] = []
    for paragraph in audit.get("paragraphs") or []:
        text = str(paragraph.get("text") or "")
        role = paragraph.get("role") or {}
        role_name = str(role.get("role") or "")
        paragraph_evidence.append(
            {
                "index": paragraph.get("index"),
                "text": text,
                "styleId": paragraph.get("styleId"),
                "styleName": paragraph.get("styleName"),
                "role": role,
                "roleName": role_name,
                "roleConfidence": role.get("confidence"),
                "manualNumbering": paragraph.get("manualNumbering"),
                "numId": paragraph.get("numId"),
                "ilvl": paragraph.get("ilvl"),
                "hasDrawing": paragraph.get("hasDrawing"),
                "drawingCount": paragraph.get("drawingCount"),
                "hasSectPr": paragraph.get("hasSectPr"),
                "isBlank": not text.strip(),
                "candidateLabels": {
                    "abstract": bool(ABSTRACT_LABEL_RE.match(text.strip())),
                    "keywords": bool(KEYWORDS_LABEL_RE.match(text.strip())),
                    "references": bool(REFERENCES_RE.match(text.strip())),
                    "acknowledgements": role_name == "ack_heading" or bool(ACKNOWLEDGEMENTS_RE.match(text.strip())),
                    "appendix": role_name == "appendix_heading" or bool(APPENDIX_RE.match(text.strip())),
                    "caption": role_name in {"figure_caption", "table_caption"} or bool(CAPTION_RE.match(text.strip())),
                },
                "candidateHeading": bool(HEADING_HINT_RE.match(text.strip()))
                or role_name.startswith("heading")
                or role_name in {"ack_heading", "appendix_heading"},
            }
        )
    header_footer_parts = audit.get("headerFooterParts") or {}
    preservation_hints = audit.get("preservationHints") or {}
    caption_layout = audit.get("captionLayout") or {}
    cross_reference_analysis = audit.get("crossReferenceAnalysis") or {}
    body_sequence = audit.get("bodySequence") or []
    drawing_paragraphs = [
        {
            "sourceIndex": item.get("index"),
            "drawingCount": item.get("drawingCount"),
            "text": str(item.get("text") or "")[:120],
        }
        for item in paragraph_evidence
        if item.get("hasDrawing")
    ]
    image_count = sum(int(item.get("drawingCount") or 0) for item in drawing_paragraphs)
    paragraph_by_index = {item.get("index"): item for item in paragraph_evidence}
    source_blocks: List[Dict[str, Any]] = []
    for item in body_sequence:
        if item.get("kind") == "paragraph":
            paragraph = paragraph_by_index.get(item.get("paragraphIndex")) or {}
            source_blocks.append(
                {
                    "kind": "paragraph",
                    "sourceIndex": paragraph.get("index"),
                    "text": paragraph.get("text"),
                    "role": paragraph.get("roleName"),
                    "styleId": paragraph.get("styleId"),
                    "styleName": paragraph.get("styleName"),
                    "attributes": {
                        "hasDrawing": paragraph.get("hasDrawing"),
                        "drawingCount": paragraph.get("drawingCount"),
                        "hasSectPr": paragraph.get("hasSectPr"),
                    },
                    "sourceAnchor": {"kind": "docx_paragraph", "index": paragraph.get("index")},
                }
            )
        elif item.get("kind") == "table":
            source_blocks.append(
                {
                    "kind": "table",
                    "sourceIndex": item.get("tableIndex"),
                    "sourceAnchor": {"kind": "docx_table", "index": item.get("tableIndex")},
                }
            )
    return {
        "source": str(source.resolve()),
        "sourceType": "docx",
        "sourceFormat": "docx",
        "evidenceVersion": "2.0",
        "paragraphCount": len(paragraph_evidence),
        "paragraphEvidence": paragraph_evidence,
        "sourceBlocks": source_blocks,
        "sections": audit.get("sections") or [],
        "candidateBlocks": [item for item in paragraph_evidence if item["candidateHeading"] or any(item["candidateLabels"].values())][:160],
        "preserveableAssets": {
            "tables": len(audit.get("tables") or []),
            "images": image_count,
            "equations": len(audit.get("equationAnalysis") or []),
            "headers": len([name for name in header_footer_parts if "header" in name]),
            "footers": len([name for name in header_footer_parts if "footer" in name]),
            "sections": len(audit.get("sections") or []),
        },
        "preservationHints": preservation_hints,
        "assetAnchorHints": {
            "captions": (caption_layout.get("captions") or [])[:80],
            "captionIssues": (caption_layout.get("issues") or [])[:80],
            "crossReferences": (cross_reference_analysis.get("references") or [])[:80],
            "imageUsage": drawing_paragraphs[:80],
            "drawingParagraphs": drawing_paragraphs[:80],
            "bodySequencePreview": body_sequence[:200],
        },
        "ambiguities": [],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract source evidence before thesis structure recovery.")
    parser.add_argument("source", help="Input .docx, .md, or .txt")
    parser.add_argument("--template-docx", help="Optional template DOCX for DOCX inputs")
    parser.add_argument("--rules-yaml", help="Optional rules YAML for DOCX inputs")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    source = Path(args.source)
    suffix = source.suffix.lower()
    if suffix == ".docx":
        payload = extract_docx_evidence(source, args.template_docx, args.rules_yaml)
    elif suffix in {".md", ".markdown", ".txt"}:
        payload = extract_text_evidence(source)
    else:
        raise SystemExit(f"Unsupported source type: {suffix}")

    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
