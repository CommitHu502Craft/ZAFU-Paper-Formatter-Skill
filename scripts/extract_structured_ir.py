#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx_ooxml import parse_document, write_json
from extract_source_evidence import extract_docx_evidence, extract_text_evidence
from thesis_ir import build_thesis_ir


IMAGE_RE = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)]+)\)$")
HEADING_RE = re.compile(r"^(?P<marks>#{1,6})\s+(?P<text>.+?)\s*$")
PLAIN_TEXT_HEADING_RE = re.compile(
    r"^(?:"
    r"\d+(?:\.\d+)*\s+.+|"
    r"[一二三四五六七八九十百千]+[、.．]\s*.+|"
    r"第[一二三四五六七八九十百千]+章.+|"
    r"第[一二三四五六七八九十百千]+节.+|"
    r"参考文献|致谢|附录.+|摘要|English Abstract|ABSTRACT"
    r")$"
)


def shorten(text: str, limit: int = 200) -> str:
    value = (text or "").strip()
    if len(value) <= limit:
        return value
    return value[: limit - 1] + "…"


def block_limit(blocks: List[Dict[str, Any]], max_blocks: Optional[int]) -> List[Dict[str, Any]]:
    if max_blocks is None or max_blocks < 0:
        return blocks
    return blocks[:max_blocks]


def extract_docx_ir(
    source: Path,
    template_docx: Optional[str],
    rules_yaml: Optional[str],
    max_blocks: Optional[int],
) -> Dict[str, Any]:
    audit = parse_document(str(source), template_docx=template_docx, rules_path=rules_yaml)
    paragraph_map = {item["index"]: item for item in audit.get("paragraphs") or []}
    table_map = {item["index"]: item for item in audit.get("tables") or []}
    blocks: List[Dict[str, Any]] = []

    for item in audit.get("bodySequence") or []:
        if item.get("kind") == "paragraph":
            paragraph = paragraph_map.get(item.get("paragraphIndex"), {})
            effective_paragraph = paragraph.get("effectiveParagraph") or {}
            effective_run = paragraph.get("effectiveRunSummary") or {}
            blocks.append(
                {
                    "kind": "paragraph",
                    "index": paragraph.get("index"),
                    "role": (paragraph.get("role") or {}).get("role"),
                    "roleConfidence": (paragraph.get("role") or {}).get("confidence"),
                    "text": paragraph.get("text"),
                    "styleId": paragraph.get("styleId"),
                    "styleName": paragraph.get("styleName"),
                    "manualNumbering": paragraph.get("manualNumbering"),
                    "numbering": {
                        "numId": paragraph.get("numId"),
                        "ilvl": paragraph.get("ilvl"),
                        "outlineLvl": paragraph.get("outlineLvl"),
                    },
                    "effectiveParagraph": {
                        "jc": effective_paragraph.get("jc"),
                        "spacing": effective_paragraph.get("spacing"),
                        "ind": effective_paragraph.get("ind"),
                        "pageBreakBefore": effective_paragraph.get("pageBreakBefore"),
                    },
                    "effectiveRun": {
                        "fonts": (effective_run.get("fonts") or {}).get("resolved"),
                        "sizePt": effective_run.get("sizePt"),
                        "bold": effective_run.get("bold"),
                        "italic": effective_run.get("italic"),
                    },
                    "flags": {
                        "hasMath": paragraph.get("hasMath"),
                        "hasFootnoteRef": paragraph.get("hasFootnoteRef"),
                        "hasEndnoteRef": paragraph.get("hasEndnoteRef"),
                        "hasTocField": paragraph.get("hasTocField"),
                        "hasSectPr": paragraph.get("hasSectPr"),
                        "breakTypes": paragraph.get("breakTypes") or [],
                    },
                }
            )
        elif item.get("kind") == "table":
            table = table_map.get(item.get("tableIndex"), {})
            blocks.append(
                {
                    "kind": "table",
                    "index": table.get("index"),
                    "textPreview": table.get("textPreview"),
                }
            )
        else:
            blocks.append(item)

    title_candidates = [
        paragraph.get("text")
        for paragraph in audit.get("paragraphs") or []
        if (paragraph.get("role") or {}).get("role") in {"title_cn", "title_en"}
    ]

    return {
        "generatedAt": audit.get("generatedAt"),
        "source": str(source.resolve()),
        "sourceKind": "docx",
        "extractionMode": "ooxml_structured_ir",
        "templateDocx": template_docx,
        "rulesPath": rules_yaml,
        "documentSummary": audit.get("summary"),
        "titleCandidates": title_candidates,
        "preservationHints": audit.get("preservationHints"),
        "sections": audit.get("sections"),
        "frontMatterAnalysis": audit.get("frontMatterAnalysis"),
        "captionLayout": audit.get("captionLayout"),
        "crossReferenceAnalysis": audit.get("crossReferenceAnalysis"),
        "numberingAnalysis": {
            "dominantFamily": (audit.get("numberingAnalysis") or {}).get("dominantFamily"),
            "familiesPresent": (audit.get("numberingAnalysis") or {}).get("familiesPresent"),
            "issues": (audit.get("numberingAnalysis") or {}).get("issues"),
        },
        "blockCount": len(blocks),
        "blocks": block_limit(blocks, max_blocks),
        "notes": [
            "This IR preserves Word-aware structure for planning and validation.",
            "Use Markdown only as a text-first ingest path, not as the universal lossless format.",
        ],
    }


def table_separator(line: str) -> bool:
    raw = line.strip()
    if not raw.startswith("|"):
        return False
    cells = [cell.strip() for cell in raw.strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell or "") for cell in cells)


def flush_paragraph(buffer: List[str], blocks: List[Dict[str, Any]]) -> None:
    if not buffer:
        return
    text = "\n".join(line.rstrip() for line in buffer).strip()
    if text:
        blocks.append({"kind": "paragraph", "text": text})
    buffer.clear()


def extract_text_ir(source: Path, max_blocks: Optional[int]) -> Dict[str, Any]:
    lines = source.read_text(encoding="utf-8").splitlines()
    blocks: List[Dict[str, Any]] = []
    if source.suffix.lower() == ".txt":
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            if PLAIN_TEXT_HEADING_RE.match(stripped):
                level = 1
                if stripped in {"摘要", "English Abstract", "ABSTRACT"}:
                    level = 1
                elif re.match(r"^(?:[一二三四五六七八九十百千]+[、.．]|第[一二三四五六七八九十百千]+节)", stripped):
                    level = 2
                blocks.append({"kind": "heading", "level": level, "text": stripped})
            else:
                blocks.append({"kind": "paragraph", "text": line.rstrip()})
        title_candidates = []
        for block in blocks:
            if block.get("kind") == "heading" and block.get("level") == 1:
                title_candidates.append(block.get("text"))
                break
        if not title_candidates:
            for block in blocks:
                if block.get("kind") == "paragraph":
                    title_candidates.append(shorten(block.get("text", ""), 80))
                    break
        return {
            "generatedAt": None,
            "source": str(source.resolve()),
            "sourceKind": source.suffix.lower().lstrip(".") or "text",
            "extractionMode": "plain_text_structured_ir",
            "documentSummary": {
                "lineCount": len(lines),
                "blockCount": len(blocks),
            },
            "titleCandidates": title_candidates,
            "blockCount": len(blocks),
            "blocks": block_limit(blocks, max_blocks),
            "notes": [
                "Plain TXT keeps one non-empty line as one structure candidate block by default.",
                "This avoids flattening degraded thesis text before structure recovery.",
            ],
        }
    paragraph_buffer: List[str] = []
    index = 0

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        heading_match = HEADING_RE.match(stripped)
        image_match = IMAGE_RE.match(stripped)

        if not stripped:
            flush_paragraph(paragraph_buffer, blocks)
            index += 1
            continue
        if heading_match:
            flush_paragraph(paragraph_buffer, blocks)
            blocks.append(
                {
                    "kind": "heading",
                    "level": len(heading_match.group("marks")),
                    "text": heading_match.group("text").strip(),
                }
            )
            index += 1
            continue
        if image_match:
            flush_paragraph(paragraph_buffer, blocks)
            blocks.append(
                {
                    "kind": "image",
                    "alt": image_match.group("alt"),
                    "path": image_match.group("path"),
                }
            )
            index += 1
            continue
        if stripped.startswith("|") and index + 1 < len(lines) and table_separator(lines[index + 1]):
            flush_paragraph(paragraph_buffer, blocks)
            table_lines = [line, lines[index + 1]]
            index += 2
            while index < len(lines) and lines[index].strip().startswith("|"):
                table_lines.append(lines[index])
                index += 1
            blocks.append({"kind": "table_markdown", "lines": table_lines})
            continue

        paragraph_buffer.append(line)
        index += 1

    flush_paragraph(paragraph_buffer, blocks)

    title_candidates = []
    for block in blocks:
        if block.get("kind") == "heading" and block.get("level") == 1:
            title_candidates.append(block.get("text"))
    if not title_candidates:
        for block in blocks:
            if block.get("kind") == "paragraph":
                title_candidates.append(shorten(block.get("text", ""), 80))
                break

    return {
        "generatedAt": None,
        "source": str(source.resolve()),
        "sourceKind": source.suffix.lower().lstrip(".") or "text",
        "extractionMode": "markdown_text_structured_ir",
        "documentSummary": {
            "lineCount": len(lines),
            "blockCount": len(blocks),
        },
        "titleCandidates": title_candidates,
        "blockCount": len(blocks),
        "blocks": block_limit(blocks, max_blocks),
        "notes": [
            "This path is intentionally lighter than DOCX extraction.",
            "Markdown and text inputs should still be normalized before building the source DOCX.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the unified ThesisIR from DOCX, Markdown, or text.")
    parser.add_argument("source", help="Input .docx, .md, or .txt")
    parser.add_argument("--template-docx", help="Optional template DOCX for DOCX extraction")
    parser.add_argument("--rules-yaml", help="Optional rules YAML for DOCX extraction")
    parser.add_argument("--output", "-o", help="Output JSON path")
    parser.add_argument("--max-blocks", type=int, help="Optional block limit for sample outputs")
    parser.add_argument(
        "--legacy-source-ir",
        action="store_true",
        help="Emit the deprecated source-specific IR instead of unified ThesisIR v2",
    )
    args = parser.parse_args()

    source = Path(args.source)
    suffix = source.suffix.lower()
    if args.legacy_source_ir and suffix == ".docx":
        data = extract_docx_ir(source, args.template_docx, args.rules_yaml, args.max_blocks)
    elif args.legacy_source_ir and suffix in {".md", ".markdown", ".txt"}:
        data = extract_text_ir(source, args.max_blocks)
    elif suffix == ".docx":
        evidence = extract_docx_evidence(source, args.template_docx, args.rules_yaml)
        data = build_thesis_ir(evidence)
    elif suffix in {".md", ".markdown", ".txt"}:
        evidence = extract_text_evidence(source)
        data = build_thesis_ir(evidence)
    else:
        raise SystemExit(f"Unsupported source type: {suffix}")

    write_json(data, args.output)


if __name__ == "__main__":
    main()
