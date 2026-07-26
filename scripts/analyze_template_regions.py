#!/usr/bin/env python3
"""Analyze a thesis template (or any thesis DOCX) into named structural regions.

Regions: cover, integrity_statement, toc, cn_abstract, en_abstract,
body_start, references, acknowledgements, appendix, final_sectPr.

Each region records its anchor paragraph index, matched marker, style id,
visible text, confidence, and replaceability policy. Detection combines
marker text, TOC fields, section breaks, and page-break boundaries — never
paragraph counts alone. The output feeds template_overlay / hybrid_rebuild
decisions and is written to debug output as template_regions.json.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def paragraph_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS))


def _squash(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


REGION_SPECS: List[Dict[str, Any]] = [
    {
        "key": "integrity_statement",
        "markers": [r"诚信承诺", r"承诺书", r"原创性声明"],
        "mustPreserve": True,
        "allowReplace": False,
    },
    {
        "key": "toc",
        "markers": [r"^目录$", r"^目次$"],
        "mustPreserve": True,
        "allowReplace": False,
    },
    {
        "key": "cn_abstract",
        "markers": [r"^摘要[:：]", r"^中文摘要$", r"^摘要$"],
        "allowLongText": True,
        "mustPreserve": False,
        "allowReplace": True,
    },
    {
        "key": "en_abstract",
        "markers": [r"(?i)^abstract[::]?", r"(?i)^abstract$"],
        "allowLongText": True,
        "mustPreserve": False,
        "allowReplace": True,
    },
    {
        "key": "references",
        "markers": [r"^参考文献", r"(?i)^references$"],
        "mustPreserve": False,
        "allowReplace": True,
    },
    {
        "key": "acknowledgements",
        "markers": [r"^致谢$", r"(?i)^acknowledg"],
        "mustPreserve": False,
        "allowReplace": True,
    },
    {
        "key": "appendix",
        "markers": [r"^附录", r"(?i)^appendix"],
        "mustPreserve": False,
        "allowReplace": True,
    },
]

BODY_START_RE = re.compile(r"^(第[一二三四五六七八九十1-9]\s*[章部分]|\d{1,2}(?:\.\d+)?\s*[一-鿿A-Za-z]|引言$|绪论$|前言$)")
TOC_ENTRY_TAIL_RE = re.compile(r"[.…·]{3,}|[0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+$")


def looks_like_toc_entry(text: str) -> bool:
    """TOC entries carry dot leaders or a trailing page number / roman numeral."""
    squashed = _squash(text)
    if re.search(r"[.…·]{3,}", squashed):
        return True
    return bool(re.search(r"[^0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ][0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]{1,3}$", squashed)) and len(squashed) <= 24


def paragraph_style_id(paragraph: ET.Element) -> Optional[str]:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    return pstyle.attrib.get(qn("val")) if pstyle is not None else None


def paragraph_has_toc_field(paragraph: ET.Element) -> bool:
    for field in paragraph.findall(".//w:fldSimple", NS):
        if "TOC" in (field.attrib.get(qn("instr")) or "").upper():
            return True
    for instr in paragraph.findall(".//w:instrText", NS):
        if (instr.text or "").strip().upper().startswith("TOC"):
            return True
    return False


def paragraph_has_page_break(paragraph: ET.Element) -> bool:
    for br in paragraph.findall(".//w:br", NS):
        if br.attrib.get(qn("type")) == "page":
            return True
    ppr = paragraph.find("w:pPr", NS)
    return ppr is not None and ppr.find("w:pageBreakBefore", NS) is not None


def paragraph_has_section_break(paragraph: ET.Element) -> bool:
    ppr = paragraph.find("w:pPr", NS)
    return ppr is not None and ppr.find("w:sectPr", NS) is not None


def analyze_regions(docx_path: Path) -> Dict[str, Any]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        return {"source": str(docx_path), "regions": {}, "error": "no_body"}

    paragraphs = [child for child in body if child.tag == qn("p")]
    regions: Dict[str, Dict[str, Any]] = {}

    section_breaks = [i for i, p in enumerate(paragraphs) if paragraph_has_section_break(p)]
    page_breaks = [i for i, p in enumerate(paragraphs) if paragraph_has_page_break(p)]

    # Cover: everything before the first structural boundary.
    first_boundary = min(section_breaks + page_breaks) if (section_breaks or page_breaks) else None
    first_text_index = next((i for i, p in enumerate(paragraphs) if _squash(paragraph_text(p))), None)
    if first_text_index is not None:
        regions["cover"] = {
            "anchorParagraphIndex": first_text_index,
            "endParagraphExclusive": (first_boundary + 1) if first_boundary is not None else None,
            "matchedMarker": None,
            "styleId": paragraph_style_id(paragraphs[first_text_index]),
            "visibleText": paragraph_text(paragraphs[first_text_index]).strip()[:60],
            "confidence": 0.85 if first_boundary is not None else 0.5,
            "mustPreserve": True,
            "allowReplace": False,
            "evidence": ["document_start"] + (["followed_by_break_boundary"] if first_boundary is not None else []),
        }

    toc_field_index = next((i for i, p in enumerate(paragraphs) if paragraph_has_toc_field(p)), None)

    for spec in REGION_SPECS:
        found = None
        for index, paragraph in enumerate(paragraphs):
            raw_text = paragraph_text(paragraph)
            text = _squash(raw_text)
            max_len = 200 if spec.get("allowLongText") else 40
            if not text or len(text) > max_len:
                continue
            if spec["key"] != "toc" and looks_like_toc_entry(raw_text):
                continue
            for marker in spec["markers"]:
                if re.search(marker, text):
                    evidence = ["marker_text"]
                    confidence = 0.8
                    style_id = paragraph_style_id(paragraph)
                    if style_id and ("heading" in style_id.lower() or "title" in style_id.lower() or "toc" in style_id.lower()):
                        confidence += 0.1
                        evidence.append("heading_like_style")
                    if paragraph_has_page_break(paragraph):
                        confidence += 0.05
                        evidence.append("page_break_boundary")
                    if spec["key"] == "toc" and toc_field_index is not None and abs(toc_field_index - index) <= 3:
                        confidence = 0.98
                        evidence.append("adjacent_toc_field")
                    found = {
                        "anchorParagraphIndex": index,
                        "matchedMarker": marker,
                        "styleId": style_id,
                        "visibleText": paragraph_text(paragraph).strip()[:60],
                        "confidence": round(min(confidence, 0.98), 2),
                        "mustPreserve": spec["mustPreserve"],
                        "allowReplace": spec["allowReplace"],
                        "evidence": evidence,
                    }
                    break
            if found:
                break
        if found:
            regions[spec["key"]] = found

    # body_start: first chapter-like heading after the TOC/abstract region.
    search_from = max(
        [r["anchorParagraphIndex"] for k, r in regions.items() if k in {"toc", "cn_abstract", "en_abstract"}] or [0]
    )
    for index in range(search_from + 1, len(paragraphs)):
        text = paragraph_text(paragraphs[index]).strip()
        if not text:
            continue
        style_id = paragraph_style_id(paragraphs[index])
        heading_style = bool(style_id and ("heading" in style_id.lower() or style_id in {"1", "2", "3"}))
        if BODY_START_RE.match(text) or heading_style:
            stop_keys = {"references", "acknowledgements", "appendix"}
            if any(regions.get(k, {}).get("anchorParagraphIndex") == index for k in stop_keys):
                break
            regions["body_start"] = {
                "anchorParagraphIndex": index,
                "matchedMarker": "chapter_pattern" if BODY_START_RE.match(text) else "heading_style",
                "styleId": style_id,
                "visibleText": text[:60],
                "confidence": 0.85 if (BODY_START_RE.match(text) and heading_style) else 0.65,
                "mustPreserve": False,
                "allowReplace": True,
                "evidence": [m for m, ok in (("chapter_pattern", bool(BODY_START_RE.match(text))), ("heading_style", heading_style)) if ok],
            }
            break

    final_sectpr = body.find("w:sectPr", NS)
    regions["final_sectPr"] = {
        "present": final_sectpr is not None,
        "mustPreserve": True,
        "allowReplace": False,
        "confidence": 1.0 if final_sectpr is not None else 0.0,
        "evidence": ["body_level_sectPr"] if final_sectpr is not None else [],
    }

    ordered = sorted(
        (k for k, r in regions.items() if isinstance(r.get("anchorParagraphIndex"), int)),
        key=lambda k: regions[k]["anchorParagraphIndex"],
    )
    return {
        "source": str(docx_path),
        "paragraphCount": len(paragraphs),
        "sectionBreakParagraphs": section_breaks,
        "pageBreakParagraphs": page_breaks[:40],
        "tocFieldParagraphIndex": toc_field_index,
        "regions": regions,
        "regionOrder": ordered,
        "unlocatedRegions": [s["key"] for s in REGION_SPECS if s["key"] not in regions],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze template/thesis DOCX structural regions.")
    parser.add_argument("docx", help="Template or thesis DOCX")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    payload = analyze_regions(Path(args.docx).resolve())
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
