#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from audit_bibliography_mode import (
    audit_blocks,
    split_references_from_docx,
    split_references_from_text,
)
from docx_ooxml import load_rules, now_iso, parse_document, write_json
from extract_source_evidence import extract_text_evidence
from extract_structured_ir import extract_text_ir
from thesis_ir import build_thesis_ir


ROLE_ORDER = [
    "toc_heading",
    "toc_field",
    "title_cn",
    "abstract_paragraph_cn",
    "keywords_cn",
    "title_en",
    "abstract_paragraph_en",
    "keywords_en",
]

ROLE_LABELS = {
    "toc_heading": "目录标题",
    "toc_field": "目录域",
    "title_cn": "中文题目",
    "abstract_paragraph_cn": "中文摘要",
    "keywords_cn": "中文关键词",
    "title_en": "英文题目",
    "abstract_paragraph_en": "英文摘要",
    "keywords_en": "英文关键词",
}

SCIENCE_HEADING_RE = re.compile(r"^(?P<prefix>\d+(?:\.\d+){0,3})\s+")
HUMANITIES_CN_RE = re.compile(r"^[一二三四五六七八九十百千]+、")
HUMANITIES_SECTION_RE = re.compile(r"^（[一二三四五六七八九十百千]+）")
HUMANITIES_NUMERIC_RE = re.compile(r"^\d+、")
HUMANITIES_CHAPTER_RE = re.compile(r"^第[一二三四五六七八九十百千]+章")
HUMANITIES_JIE_RE = re.compile(r"^第[一二三四五六七八九十百千]+节")
CAPTION_LABEL_TOKEN_RE = r"(?:\d+|[一二三四五六七八九十百千]+)(?:[a-zA-Z])?(?:[-—.．](?:\d+|[一二三四五六七八九十百千]+))?"
TABLE_CAPTION_RE = re.compile(rf"^表\s*{CAPTION_LABEL_TOKEN_RE}")
FIGURE_CAPTION_RE = re.compile(rf"^图\s*{CAPTION_LABEL_TOKEN_RE}")
KEYWORDS_CN_RE = re.compile(r"^关键词[:：]")
KEYWORDS_EN_RE = re.compile(r"^key\s*words?[:：]", re.IGNORECASE)
ABSTRACT_EN_RE = re.compile(r"^abstract[:：]", re.IGNORECASE)
CROSS_REF_RE = re.compile(rf"(?P<kind>图|表)\s*(?P<label>{CAPTION_LABEL_TOKEN_RE})")
ABSTRACT_CN_HEADINGS = {"中文摘要", "摘  要", "摘要"}
ABSTRACT_EN_HEADINGS = {"English Abstract", "ABSTRACT"}


def strip_markdown_markers(text: str) -> str:
    value = (text or "").strip()
    value = re.sub(r"[*_`#]+", "", value)
    return value.strip()


def shorten(text: str, limit: int = 120) -> str:
    value = (text or "").strip()
    if len(value) <= limit:
        return value
    return value[: limit - 1] + "…"


def make_issue(kind: str, severity: str, message: str, **details: Any) -> Dict[str, Any]:
    return {
        "kind": kind,
        "severity": severity,
        "message": message,
        **details,
    }


def identify_text_role(text: str) -> Optional[str]:
    value = strip_markdown_markers(text)
    collapsed = value.replace(" ", "")
    if collapsed == "目录":
        return "toc_heading"
    if value in ABSTRACT_CN_HEADINGS:
        return "abstract_heading_cn"
    if value in ABSTRACT_EN_HEADINGS:
        return "abstract_heading_en"
    if value.startswith("摘要：") or value.startswith("摘要:"):
        return "abstract_paragraph_cn"
    if KEYWORDS_CN_RE.match(value):
        return "keywords_cn"
    if ABSTRACT_EN_RE.match(value):
        return "abstract_paragraph_en"
    if KEYWORDS_EN_RE.match(value):
        return "keywords_en"
    if re.fullmatch(r"[A-Za-z0-9,.:;()'\"“”‘’\-\s]{6,}", value) and "abstract" not in value.lower() and "key words" not in value.lower():
        return "title_en"
    if re.search(r"[\u4e00-\u9fff]", value) and len(value) <= 80 and not any(
        value.startswith(prefix) for prefix in ("图", "表", "第", "关键词", "摘要")
    ):
        return "title_cn"
    return None


def extract_front_roles_from_docx(audit: Dict[str, Any]) -> List[Dict[str, Any]]:
    roles: List[Dict[str, Any]] = []
    for paragraph in audit.get("paragraphs") or []:
        text = (paragraph.get("text") or "").strip()
        role = (paragraph.get("role") or {}).get("role")
        if role in ROLE_ORDER:
            roles.append(
                {
                    "role": role,
                    "index": paragraph.get("index"),
                    "text": shorten(text),
                }
            )
    return roles


def extract_front_roles_from_text(source: Path) -> List[Dict[str, Any]]:
    ir = extract_text_ir(source, max_blocks=80)
    roles: List[Dict[str, Any]] = []
    for index, block in enumerate(ir.get("blocks") or []):
        kind = block.get("kind")
        if kind not in {"heading", "paragraph"}:
            continue
        text = (block.get("text") or "").strip()
        if kind == "heading":
            detected_heading = detect_heading_family(strip_markdown_markers(text))
            if detected_heading and detected_heading[1] == 1 and roles:
                break
            if text in ABSTRACT_CN_HEADINGS or text in ABSTRACT_EN_HEADINGS:
                role = identify_text_role(text)
                if role:
                    roles.append({"role": role, "index": index, "text": shorten(strip_markdown_markers(text))})
                continue
        role = identify_text_role(text)
        if role:
            roles.append({"role": role, "index": index, "text": shorten(strip_markdown_markers(text))})
    return roles


def extract_front_roles_from_thesis_ir(evidence: Dict[str, Any], thesis_ir: Dict[str, Any]) -> List[Dict[str, Any]]:
    front_matter = thesis_ir.get("frontMatter") or {}
    source_blocks = {int(index) for index in (front_matter.get("sourceBlocks") or []) if index is not None}
    roles: List[Dict[str, Any]] = []
    if not source_blocks:
        return roles

    for item in evidence.get("lineEvidence") or []:
        index = item.get("index")
        if index not in source_blocks:
            continue
        text = strip_markdown_markers(str(item.get("text") or ""))
        role = identify_text_role(text)
        if role:
            roles.append({"role": role, "index": index, "text": shorten(text)})

    if front_matter.get("title") and not any(item["role"] == "title_cn" for item in roles):
        title_index = min(source_blocks)
        roles.append({"role": "title_cn", "index": title_index, "text": shorten(str(front_matter.get("title") or ""))})
    return sorted(roles, key=lambda item: (int(item.get("index") or 0), item.get("role") or ""))


def analyze_front_matter_docx(audit: Dict[str, Any], rules: Dict[str, Any]) -> Dict[str, Any]:
    detected_roles = extract_front_roles_from_docx(audit)
    issues = list((audit.get("frontMatterAnalysis") or {}).get("issues") or [])
    paragraph_map = {item.get("index"): item for item in (audit.get("paragraphs") or [])}
    detected_lookup = {item["role"]: item for item in detected_roles}

    for role, prefixes in {
        "abstract_paragraph_cn": ("摘要：", "摘要:"),
        "keywords_cn": ("关键词：", "关键词:"),
        "abstract_paragraph_en": ("Abstract:", "ABSTRACT:"),
        "keywords_en": ("Key words:", "Key Words:", "KEY WORDS:", "Keywords:"),
    }.items():
        item = detected_lookup.get(role)
        if not item:
            continue
        paragraph = paragraph_map.get(item["index"], {})
        text = (paragraph.get("text") or "").strip()
        if not text.startswith(prefixes):
            issues.append(
                make_issue(
                    "front_matter_inline_label_mismatch",
                    "high",
                    f"{ROLE_LABELS[role]}没有使用学校要求的标签前缀。",
                    role=role,
                    paragraphIndex=item["index"],
                    text=shorten(text),
                    expectedPrefixes=list(prefixes),
                )
            )
        if role == "keywords_cn":
            payload = re.sub(r"^关键词[:：]\s*", "", text)
            if payload and "；" not in payload and ";" not in payload:
                issues.append(
                    make_issue(
                        "keywords_cn_separator_mismatch",
                        "medium",
                        "中文关键词建议使用分号分隔。",
                        role=role,
                        paragraphIndex=item["index"],
                        text=shorten(text),
                    )
                )
        if role == "keywords_en":
            payload = re.sub(r"^(Key words|Key Words|KEY WORDS|Keywords)[:：]\s*", "", text, flags=re.IGNORECASE)
            if payload and ("；" in payload or ";" in payload):
                issues.append(
                    make_issue(
                        "keywords_en_separator_mismatch",
                        "medium",
                        "英文关键词建议使用逗号分隔。",
                        role=role,
                        paragraphIndex=item["index"],
                        text=shorten(text),
                    )
                )
        if role == "abstract_paragraph_en" and re.match(r"^(Abstract:|ABSTRACT:)\S", text):
            issues.append(
                make_issue(
                    "abstract_en_spacing_after_label",
                    "medium",
                    "英文摘要标签后建议保留一个英文空格。",
                    role=role,
                    paragraphIndex=item["index"],
                    text=shorten(text),
                )
            )

    expected_sequence = ["toc_heading", *(((rules.get("front_matter") or {}).get("abstract_sequence")) or [])]
    normalized_expected = []
    for item in expected_sequence:
        if item == "title_cn":
            normalized_expected.append("title_cn")
        elif item == "title_en":
            normalized_expected.append("title_en")
        elif item == "abstract_cn":
            normalized_expected.append("abstract_paragraph_cn")
        elif item == "keywords_cn":
            normalized_expected.append("keywords_cn")
        elif item == "abstract_en":
            normalized_expected.append("abstract_paragraph_en")
        elif item == "keywords_en":
            normalized_expected.append("keywords_en")
        elif item == "toc_heading":
            normalized_expected.append("toc_heading")

    return {
        "expectedSequence": normalized_expected,
        "detectedSequence": detected_roles,
        "issues": issues,
    }


def analyze_front_matter_text(source: Path, rules: Dict[str, Any]) -> Dict[str, Any]:
    detected_roles = extract_front_roles_from_text(source)
    role_positions = {item["role"]: item["index"] for item in detected_roles}
    issues: List[Dict[str, Any]] = []

    normalized_expected = [
        "abstract_heading_cn",
        "keywords_cn",
        "abstract_heading_en",
        "keywords_en",
    ]

    missing = [item for item in normalized_expected if item not in role_positions]
    if missing:
        issues.append(
            make_issue(
                "front_matter_missing_roles",
                "high",
                "前置部分缺少学校要求的关键区块。",
                roles=missing,
            )
        )

    present = [item for item in normalized_expected if item in role_positions]
    for earlier, later in zip(present, present[1:]):
        if role_positions[earlier] > role_positions[later]:
            issues.append(
                make_issue(
                    "front_matter_order_mismatch",
                    "high",
                    f"{ROLE_LABELS.get(earlier, earlier)}出现在{ROLE_LABELS.get(later, later)}之后。",
                    earlierRole=earlier,
                    laterRole=later,
                    earlierIndex=role_positions[earlier],
                    laterIndex=role_positions[later],
                )
            )

    for role, prefixes in {
        "keywords_cn": ("关键词：", "关键词:"),
        "keywords_en": ("Key words:", "Key Words:", "KEY WORDS:", "Keywords:"),
    }.items():
        item = next((entry for entry in detected_roles if entry["role"] == role), None)
        if item and not item["text"].startswith(prefixes):
            issues.append(
                make_issue(
                    "front_matter_inline_label_mismatch",
                    "high",
                    f"{ROLE_LABELS[role]}没有使用学校要求的标签前缀。",
                    role=role,
                    blockIndex=item["index"],
                    text=item["text"],
                    expectedPrefixes=list(prefixes),
                )
            )
        if item and role == "keywords_cn":
            payload = re.sub(r"^关键词[:：]\s*", "", item["text"])
            if payload and "；" not in payload and ";" not in payload:
                issues.append(
                    make_issue(
                        "keywords_cn_separator_mismatch",
                        "medium",
                        "中文关键词建议使用分号分隔。",
                        role=role,
                        blockIndex=item["index"],
                        text=item["text"],
                    )
                )
        if item and role == "keywords_en":
            payload = re.sub(r"^(Key words|Key Words|KEY WORDS|Keywords)[:：]\s*", "", item["text"], flags=re.IGNORECASE)
            if payload and ("；" in payload or ";" in payload):
                issues.append(
                    make_issue(
                        "keywords_en_separator_mismatch",
                        "medium",
                        "英文关键词建议使用逗号分隔。",
                        role=role,
                        blockIndex=item["index"],
                        text=item["text"],
                    )
                )

    return {
        "expectedSequence": normalized_expected,
        "detectedSequence": detected_roles,
        "issues": issues,
    }


def analyze_front_matter_text_from_ir(
    source: Path,
    rules: Dict[str, Any],
    evidence: Dict[str, Any],
    thesis_ir: Dict[str, Any],
) -> Dict[str, Any]:
    detected_roles = extract_front_roles_from_thesis_ir(evidence, thesis_ir)
    if not detected_roles:
        return analyze_front_matter_text(source, rules)

    role_positions = {item["role"]: item["index"] for item in detected_roles}
    issues: List[Dict[str, Any]] = []
    normalized_expected = [
        "abstract_heading_cn",
        "keywords_cn",
        "abstract_heading_en",
        "keywords_en",
    ]

    missing = [item for item in normalized_expected if item not in role_positions]
    if missing:
        issues.append(
            make_issue(
                "front_matter_missing_roles",
                "high",
                "前置部分缺少学校要求的关键区块。",
                roles=missing,
            )
        )

    present = [item for item in normalized_expected if item in role_positions]
    for earlier, later in zip(present, present[1:]):
        if role_positions[earlier] > role_positions[later]:
            issues.append(
                make_issue(
                    "front_matter_order_mismatch",
                    "high",
                    f"{ROLE_LABELS.get(earlier, earlier)}出现在{ROLE_LABELS.get(later, later)}之后。",
                    earlierRole=earlier,
                    laterRole=later,
                    earlierIndex=role_positions[earlier],
                    laterIndex=role_positions[later],
                )
            )

    for role, prefixes in {
        "keywords_cn": ("关键词：", "关键词:"),
        "keywords_en": ("Key words:", "Key Words:", "KEY WORDS:", "Keywords:"),
    }.items():
        item = next((entry for entry in detected_roles if entry["role"] == role), None)
        if item and not item["text"].startswith(prefixes):
            issues.append(
                make_issue(
                    "front_matter_inline_label_mismatch",
                    "high",
                    f"{ROLE_LABELS[role]}没有使用学校要求的标签前缀。",
                    role=role,
                    blockIndex=item["index"],
                    text=item["text"],
                    expectedPrefixes=list(prefixes),
                )
            )

    return {
        "expectedSequence": normalized_expected,
        "detectedSequence": detected_roles,
        "issues": issues,
    }


def detect_heading_family(text: str) -> Optional[Tuple[str, int, Optional[List[int]]]]:
    value = (text or "").strip()
    match = SCIENCE_HEADING_RE.match(value)
    if match:
        numbers = [int(part) for part in match.group("prefix").split(".")]
        return ("science_decimal", len(numbers), numbers)
    if HUMANITIES_CHAPTER_RE.match(value):
        return ("humanities_chapter", 1, None)
    if HUMANITIES_JIE_RE.match(value):
        return ("humanities_chapter", 2, None)
    if HUMANITIES_CN_RE.match(value):
        return ("humanities_cn", 1, None)
    if HUMANITIES_SECTION_RE.match(value):
        return ("humanities_cn", 2, None)
    if HUMANITIES_NUMERIC_RE.match(value):
        return ("humanities_cn", 3, None)
    return None


def analyze_heading_blocks_text(source: Path) -> Dict[str, Any]:
    ir = extract_text_ir(source, max_blocks=None)
    heading_entries: List[Dict[str, Any]] = []
    families_present: List[str] = []
    issues: List[Dict[str, Any]] = []
    last_level: Optional[int] = None
    last_numbers: Optional[List[int]] = None

    for index, block in enumerate(ir.get("blocks") or []):
        if block.get("kind") != "heading":
            continue
        text = (block.get("text") or "").strip()
        detected = detect_heading_family(text)
        family = detected[0] if detected else "unknown"
        level = detected[1] if detected else block.get("level")
        numbers = detected[2] if detected else None
        heading_entries.append(
            {
                "index": index,
                "text": shorten(text),
                "family": family,
                "level": level,
            }
        )
        if family != "unknown" and family not in families_present:
            families_present.append(family)
        if last_level is not None and level is not None and level > last_level + 1:
            issues.append(
                make_issue(
                    "heading_level_jump",
                    "high",
                    "标题层级发生跳级。",
                    blockIndex=index,
                    previousLevel=last_level,
                    currentLevel=level,
                    text=shorten(text),
                )
            )
        if numbers and last_numbers and len(numbers) == len(last_numbers):
            if numbers[:-1] == last_numbers[:-1] and numbers[-1] > last_numbers[-1] + 1:
                issues.append(
                    make_issue(
                        "science_heading_gap",
                        "medium",
                        "理工科编号存在明显断档。",
                        blockIndex=index,
                        previous=last_numbers,
                        current=numbers,
                        text=shorten(text),
                    )
                )
        last_level = level
        last_numbers = numbers or last_numbers

    if len(families_present) > 1:
        allowed_hybrid = set(families_present) == {"humanities_chapter", "science_decimal"}
        if not allowed_hybrid:
            issues.append(
                make_issue(
                    "mixed_heading_numbering_families",
                    "high",
                    "同一文档中混用了多套标题编号体系。",
                    families=families_present,
                )
            )

    if set(families_present) == {"humanities_chapter", "science_decimal"}:
        dominant_family = "hybrid_chapter_decimal"
    else:
        dominant_family = families_present[0] if len(families_present) == 1 else None
    return {
        "dominantFamily": dominant_family,
        "familiesPresent": families_present,
        "headingCount": len(heading_entries),
        "headings": heading_entries[:120],
        "issues": issues,
    }


def analyze_heading_blocks_from_thesis_ir(thesis_ir: Dict[str, Any]) -> Dict[str, Any]:
    heading_entries: List[Dict[str, Any]] = []
    numbering = thesis_ir.get("numberingAnalysis") or {}
    families_present = list(numbering.get("familiesPresent") or [])
    issues: List[Dict[str, Any]] = []

    for item in thesis_ir.get("headingTree") or []:
        heading_entries.append(
            {
                "index": item.get("sourceIndex"),
                "text": shorten(str(item.get("text") or "")),
                "family": item.get("numberingFamily") or "unknown",
                "level": item.get("level"),
            }
        )

    for conflict in numbering.get("conflicts") or []:
        kind = conflict.get("kind")
        if kind == "level_jump":
            issues.append(
                make_issue(
                    "heading_level_jump",
                    "high",
                    "标题层级发生跳级。",
                    blockIndex=conflict.get("sourceIndex"),
                    previousLevel=conflict.get("previousLevel"),
                    currentLevel=conflict.get("currentLevel"),
                    text=shorten(str(conflict.get("text") or "")),
                )
            )
        elif kind == "science_decimal_gap":
            issues.append(
                make_issue(
                    "science_heading_gap",
                    "medium",
                    "理工科编号存在明显断档。",
                    blockIndex=conflict.get("sourceIndex"),
                    previous=conflict.get("previous"),
                    current=conflict.get("current"),
                    text=shorten(str(conflict.get("text") or "")),
                )
            )

    if len(families_present) > 1 and set(families_present) != {"cn_chapter", "science_decimal"}:
        issues.append(
            make_issue(
                "mixed_heading_numbering_families",
                "high",
                "同一文档中混用了多套标题编号体系。",
                families=families_present,
            )
        )

    dominant_family = numbering.get("dominantFamily")
    if dominant_family == "cn_chapter":
        dominant_family = "humanities_chapter"
    return {
        "dominantFamily": dominant_family,
        "familiesPresent": families_present,
        "headingCount": len(heading_entries),
        "headings": heading_entries[:120],
        "issues": issues,
    }


def analyze_heading_docx(audit: Dict[str, Any]) -> Dict[str, Any]:
    numbering = audit.get("numberingAnalysis") or {}
    issues: List[Dict[str, Any]] = []
    for item in numbering.get("issues") or []:
        issues.append(make_issue("numbering_issue", "high", str(item), raw=item))
    for item in (numbering.get("anomalies") or [])[:80]:
        issues.append(make_issue("numbering_anomaly", "high", "检测到标题编号异常。", raw=item))
    if numbering.get("mixedManualAndAuto"):
        issues.append(
            make_issue(
                "mixed_manual_and_auto_heading_numbering",
                "high",
                "文档同时存在手打标题编号和自动编号污染。",
            )
        )
    heading_count = 0
    for paragraph in audit.get("paragraphs") or []:
        role = (paragraph.get("role") or {}).get("role")
        if role in {"heading1", "heading2", "heading3", "heading4"}:
            heading_count += 1
    return {
        "dominantFamily": numbering.get("dominantFamily"),
        "familiesPresent": numbering.get("familiesPresent") or [],
        "headingCount": heading_count,
        "issues": issues,
        "repairHints": numbering.get("repairActions") or [],
    }


def analyze_caption_text(source: Path) -> Dict[str, Any]:
    ir = extract_text_ir(source, max_blocks=None)
    blocks = ir.get("blocks") or []
    issues: List[Dict[str, Any]] = []
    captions: List[Dict[str, Any]] = []

    for index, block in enumerate(blocks):
        if block.get("kind") != "paragraph":
            continue
        text = (block.get("text") or "").strip()
        if TABLE_CAPTION_RE.match(text):
            captions.append({"index": index, "kind": "table", "text": shorten(text)})
            next_kind = blocks[index + 1].get("kind") if index + 1 < len(blocks) else None
            if next_kind == "table_markdown":
                issues.append(
                    make_issue(
                        "table_caption_before_table",
                        "medium",
                        "表题出现在表格之前，学校当前要求是表题放在表格下一行。",
                        blockIndex=index,
                        text=shorten(text),
                    )
                )
        elif FIGURE_CAPTION_RE.match(text):
            captions.append({"index": index, "kind": "figure", "text": shorten(text)})

    caption_labels = {
        re.sub(r"\s+", "", item["text"].split(" ", 1)[0].replace("—", "-").replace("．", ".").replace(".", "-")): item
        for item in captions
    }
    unresolved = []
    for index, block in enumerate(blocks):
        if block.get("kind") not in {"heading", "paragraph"}:
            continue
        text = (block.get("text") or "").strip()
        for match in CROSS_REF_RE.finditer(text):
            key = f"{match.group('kind')}{match.group('label').replace('—', '-').replace('．', '.').replace('.', '-')}"
            if key not in caption_labels:
                unresolved.append(
                    {
                        "blockIndex": index,
                        "reference": f"{match.group('kind')}{match.group('label')}",
                        "text": shorten(text),
                    }
                )

    return {
        "captions": captions[:80],
        "issues": issues,
        "crossReferenceUnresolved": unresolved[:80],
    }


def analyze_caption_docx(audit: Dict[str, Any]) -> Dict[str, Any]:
    caption_layout = audit.get("captionLayout") or {}
    cross_reference = audit.get("crossReferenceAnalysis") or {}
    issues: List[Dict[str, Any]] = []
    for item in caption_layout.get("issues") or []:
        issues.append(
            make_issue(
                item.get("kind") or "caption_issue",
                "medium",
                "检测到图表题注位置问题。",
                paragraphIndex=item.get("paragraphIndex"),
                text=shorten(item.get("text") or ""),
            )
        )
    return {
        "captions": (caption_layout.get("captions") or [])[:80],
        "issues": issues,
        "crossReferenceUnresolved": (cross_reference.get("unresolved") or [])[:80],
    }


def analyze_reference_indentation_docx(audit: Dict[str, Any], rules: Dict[str, Any]) -> Dict[str, Any]:
    split = split_references_from_docx(audit)
    paragraph_map = {item.get("index"): item for item in (audit.get("paragraphs") or [])}
    expected_indent_chars = (
        ((rules.get("references") or {}).get("first_line_indent_chars"))
        if (rules.get("references") or {}).get("first_line_indent_chars") is not None
        else (((rules.get("styles") or {}).get("references_cn") or {}).get("indent_chars", 0))
    )
    issues: List[Dict[str, Any]] = []

    for entry in split.get("bibliographyEntries") or []:
        paragraph = paragraph_map.get(entry.get("index")) or {}
        effective = (paragraph.get("effectiveParagraph") or {}).get("ind") or {}
        first_line_chars = effective.get("firstLineChars")
        first_line_twips = effective.get("firstLine")
        hanging_chars = effective.get("hangingChars")
        hanging_twips = effective.get("hanging")
        has_indent = any(
            value not in (None, 0, "0")
            for value in (first_line_chars, first_line_twips, hanging_chars, hanging_twips)
        )
        if has_indent:
            issues.append(
                make_issue(
                    "reference_entry_indent_mismatch",
                    "medium",
                    "参考文献条目带有首行或悬挂缩进，不应继承正文两字符首行缩进。",
                    paragraphIndex=entry.get("index"),
                    text=shorten(entry.get("text") or ""),
                    effectiveIndent=effective,
                )
            )

    return {
        "expectedFirstLineIndentChars": expected_indent_chars,
        "issues": issues[:80],
        "entryCount": len(split.get("bibliographyEntries") or []),
    }


def analyze_reference_indentation_text(_: Path, rules: Dict[str, Any]) -> Dict[str, Any]:
    expected_indent_chars = (
        ((rules.get("references") or {}).get("first_line_indent_chars"))
        if (rules.get("references") or {}).get("first_line_indent_chars") is not None
        else (((rules.get("styles") or {}).get("references_cn") or {}).get("indent_chars", 0))
    )
    return {
        "expectedFirstLineIndentChars": expected_indent_chars,
        "issues": [],
        "entryCount": None,
        "note": "纯文本输入不保留 Word 段落缩进，预检阶段只声明目标规则。",
    }


def build_actions(
    source_kind: str,
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    bibliography: Dict[str, Any],
    caption_audit: Dict[str, Any],
    reference_indent: Dict[str, Any],
) -> List[Dict[str, Any]]:
    actions: List[Dict[str, Any]] = []

    if front_matter.get("issues"):
        actions.append(
            {
                "priority": 1,
                "action": "normalize_front_matter_structure",
                "stage": "preflight",
                "confidence": "high",
                "reason": "目录、摘要、关键词或中英文题名顺序与学校要求不一致。",
                "preferredExecution": "reuse_template_front_matter_skeleton_for_docx" if source_kind == "docx" else "normalize_text_ir_before_build",
            }
        )

    if heading_audit.get("issues"):
        actions.append(
            {
                "priority": 2,
                "action": "normalize_heading_numbering_tree",
                "stage": "preflight",
                "confidence": "high",
                "reason": "标题层级或编号体系存在异常，应先纠偏再做版式修复。",
            }
        )

    if bibliography.get("issues"):
        actions.append(
            {
                "priority": 3,
                "action": "normalize_bibliography_mode",
                "stage": "preflight",
                "confidence": "high",
                "reason": "正文引用或参考文献列表与学校 author-year 模式不一致。",
            }
        )

    if reference_indent.get("issues"):
        actions.append(
            {
                "priority": 4,
                "action": "clear_reference_body_indent_inheritance",
                "stage": "preflight",
                "confidence": "high",
                "reason": "参考文献不应继承正文首行缩进规则。",
            }
        )

    if caption_audit.get("issues"):
        actions.append(
            {
                "priority": 5,
                "action": "normalize_caption_positions",
                "stage": "preflight",
                "confidence": "medium",
                "reason": "图表题注位置不符合学校要求，尤其是表题应放在表格下一行。",
            }
        )

    if caption_audit.get("crossReferenceUnresolved"):
        actions.append(
            {
                "priority": 6,
                "action": "resolve_or_defer_cross_references",
                "stage": "preflight",
                "confidence": "medium",
                "reason": "存在无法定位到题注锚点的图表交叉引用。",
            }
        )

    return actions


def build_confirmation_requests(
    source_kind: str,
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    bibliography: Dict[str, Any],
    caption_audit: Dict[str, Any],
    reference_indent: Dict[str, Any],
) -> List[Dict[str, Any]]:
    requests: List[Dict[str, Any]] = []

    heading_issue_kinds = {item.get("kind") for item in (heading_audit.get("issues") or [])}
    if heading_issue_kinds:
        requests.append(
            {
                "decisionId": "confirm_heading_numbering_rewrite",
                "priority": 2,
                "requiresConfirmation": True,
                "scope": "heading_numbering",
                "question": "检测到标题树或编号体系异常，是否允许自动重建或批量改写标题编号？",
                "reason": "标题编号修复会直接改变章节前缀，影响目录、交叉引用和全篇层级显示。",
                "impact": [
                    "可能批量修改一级、二级、三级标题前缀",
                    "可能统一手打编号与自动编号的混合污染",
                    "可能影响目录和图表编号联动结果",
                ],
                "recommendedDefault": "ask_user_first",
                "relatedIssueKinds": sorted(heading_issue_kinds),
            }
        )

    front_issue_kinds = {item.get("kind") for item in (front_matter.get("issues") or [])}
    front_matter_confirmation_kinds = {"front_matter_order_mismatch", "toc_after_abstract"}
    if front_issue_kinds & front_matter_confirmation_kinds:
        requests.append(
            {
                "decisionId": "confirm_front_matter_rebuild",
                "priority": 3,
                "requiresConfirmation": True,
                "scope": "front_matter",
                "question": "前置部分顺序或摘要格式与学校要求不一致，是否允许按模板骨架重建目录、摘要和题名区域？",
                "reason": "前置部分修复可能移动段落位置、替换局部结构，并优先套用模板骨架。",
                "impact": [
                    "可能调整目录、中文摘要、英文摘要的先后顺序",
                    "可能重建题名、摘要、关键词段落的标签与布局",
                    "DOCX 输入时可能优先复用模板前置骨架",
                ],
                "recommendedDefault": "ask_user_first",
                "relatedIssueKinds": sorted(front_issue_kinds & front_matter_confirmation_kinds),
                "preferredExecution": "reuse_template_front_matter_skeleton_for_docx" if source_kind == "docx" else "normalize_text_ir_before_build",
            }
        )

    if caption_audit.get("crossReferenceUnresolved"):
        requests.append(
            {
                "decisionId": "confirm_cross_reference_rewrite",
                "priority": 4,
                "requiresConfirmation": True,
                "scope": "cross_references",
                "question": "检测到图表交叉引用需要补锚点或改写为 Word 域，是否继续自动处理交叉引用？",
                "reason": "交叉引用修复会改写局部文本为 Word 字段或书签引用，影响后续编辑体验。",
                "impact": [
                    "可能插入书签和 REF 字段",
                    "可能改写纯文本图表引用",
                    "未解析成功的引用仍会保留人工复核项",
                ],
                "recommendedDefault": "ask_user_first",
                "relatedIssueKinds": ["unresolved_cross_reference"],
            }
        )

    if reference_indent.get("issues") and not bibliography.get("issues"):
        requests.append(
            {
                "decisionId": "confirm_reference_indent_cleanup",
                "priority": 5,
                "requiresConfirmation": True,
                "scope": "reference_paragraph_indent",
                "question": "检测到参考文献继承了正文缩进，是否只清理参考文献段落缩进而不改正文内容？",
                "reason": "这是低风险排版修复，但会批量修改参考文献段落属性。",
                "impact": [
                    "只改参考文献段落缩进属性",
                    "不改参考文献条目文本内容",
                ],
                "recommendedDefault": "ask_user_first",
                "relatedIssueKinds": [item.get("kind") for item in (reference_indent.get("issues") or [])[:20]],
            }
        )

    return requests

def collect_manual_review(
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    bibliography: Dict[str, Any],
    caption_audit: Dict[str, Any],
    reference_indent: Dict[str, Any],
) -> List[Dict[str, Any]]:
    review: List[Dict[str, Any]] = []
    for bucket, payload in {
        "front_matter": front_matter.get("issues") or [],
        "headings": heading_audit.get("issues") or [],
        "bibliography": bibliography.get("issues") or [],
        "captions": caption_audit.get("issues") or [],
        "reference_indent": reference_indent.get("issues") or [],
    }.items():
        for item in payload[:40]:
            review.append(
                {
                    "bucket": bucket,
                    "kind": item.get("kind"),
                    "message": item.get("message"),
                    "details": item,
                }
            )
    for item in (caption_audit.get("crossReferenceUnresolved") or [])[:40]:
        review.append(
            {
                "bucket": "cross_references",
                "kind": "unresolved_cross_reference",
                "message": "交叉引用未能匹配到题注锚点。",
                "details": item,
            }
        )
    return review


def summarize_issue_counts(issues: List[Dict[str, Any]]) -> Dict[str, int]:
    counts = {"high": 0, "medium": 0, "low": 0}
    for item in issues:
        severity = str(item.get("severity") or "").lower()
        if severity in counts:
            counts[severity] += 1
    return counts


def build_style_drift_docx(audit: Dict[str, Any]) -> List[Dict[str, Any]]:
    drift: List[Dict[str, Any]] = []
    for mismatch in ((audit.get("rulesComparison") or {}).get("mismatches") or []):
        if mismatch.get("scope") == "style":
            drift.append(
                {
                    "kind": "rule_style_mismatch",
                    "field": mismatch.get("field"),
                    "ruleValue": mismatch.get("ruleValue"),
                    "baselineValue": mismatch.get("baselineValue"),
                }
            )
    heading_risks = ((audit.get("styleAnalysis") or {}).get("builtinHeadingRisks") or [])[:20]
    if heading_risks:
        drift.append(
            {
                "kind": "builtin_heading_risks",
                "count": len((audit.get("styleAnalysis") or {}).get("builtinHeadingRisks") or []),
                "examples": heading_risks,
            }
        )
    suspicious_runs = ((audit.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or [])[:20]
    if suspicious_runs:
        drift.append(
            {
                "kind": "font_slot_risks",
                "count": len((audit.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or []),
                "examples": suspicious_runs,
            }
        )
    return drift


def build_numbering_drift_docx(audit: Dict[str, Any], heading_audit: Dict[str, Any]) -> List[Dict[str, Any]]:
    numbering = audit.get("numberingAnalysis") or {}
    drift: List[Dict[str, Any]] = []
    for issue in numbering.get("issues") or []:
        drift.append(
            {
                "kind": issue.get("kind") or "numbering_issue",
                "details": issue,
            }
        )
    anomalies = numbering.get("anomalies") or []
    if anomalies:
        drift.append(
            {
                "kind": "numbering_anomalies",
                "count": len(anomalies),
                "examples": anomalies[:20],
            }
        )
    if heading_audit.get("dominantFamily") and heading_audit.get("familiesPresent"):
        families_present = heading_audit.get("familiesPresent") or []
        if len([item for item in families_present if item]) > 1:
            drift.append(
                {
                    "kind": "mixed_heading_numbering_families",
                    "families": families_present,
                }
            )
    return drift


def build_section_drift_docx(audit: Dict[str, Any]) -> List[Dict[str, Any]]:
    drift: List[Dict[str, Any]] = []
    for mismatch in ((audit.get("rulesComparison") or {}).get("mismatches") or []):
        if mismatch.get("scope") == "page":
            drift.append(
                {
                    "kind": "page_geometry_mismatch",
                    "field": mismatch.get("field"),
                    "ruleValue": mismatch.get("ruleValue"),
                    "baselineValue": mismatch.get("baselineValue"),
                }
            )
    sections = audit.get("sections") or []
    template_sections = ((audit.get("templateBaseline") or {}).get("sections") or [])
    if template_sections and len(sections) != len(template_sections):
        drift.append(
            {
                "kind": "section_count_mismatch",
                "actual": len(sections),
                "baseline": len(template_sections),
            }
        )
    if len(sections) >= 6:
        drift.append(
            {
                "kind": "many_sections",
                "actual": len(sections),
            }
        )
    return drift


def compute_template_similarity(
    audit: Dict[str, Any],
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    caption_audit: Dict[str, Any],
    style_drift: List[Dict[str, Any]],
    numbering_drift: List[Dict[str, Any]],
    section_drift: List[Dict[str, Any]],
) -> Optional[float]:
    if not audit.get("templateDocx"):
        return None

    score = 1.0
    front_counts = summarize_issue_counts(front_matter.get("issues") or [])
    heading_counts = summarize_issue_counts(heading_audit.get("issues") or [])
    caption_issue_count = len(caption_audit.get("issues") or [])
    unresolved_cross_refs = len(caption_audit.get("crossReferenceUnresolved") or [])

    score -= min(0.28, len(style_drift) * 0.08)
    score -= min(0.32, len(numbering_drift) * 0.10)
    score -= min(0.22, len(section_drift) * 0.08)
    score -= min(0.18, front_counts["high"] * 0.06 + front_counts["medium"] * 0.03)
    score -= min(0.14, heading_counts["high"] * 0.04 + heading_counts["medium"] * 0.02)
    score -= min(0.08, caption_issue_count * 0.02)
    score -= min(0.08, unresolved_cross_refs * 0.01)
    if not ((audit.get("rulesComparison") or {}).get("match", True)):
        score -= 0.05

    return round(max(0.0, min(1.0, score)), 2)


def build_blocked_auto_repairs(
    source_kind: str,
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    bibliography: Dict[str, Any],
    caption_audit: Dict[str, Any],
    reference_indent: Dict[str, Any],
    *,
    document_risk_class: Optional[str] = None,
    section_drift: Optional[List[Dict[str, Any]]] = None,
) -> List[str]:
    blocked: List[str] = []
    heading_issue_kinds = {item.get("kind") for item in (heading_audit.get("issues") or [])}
    front_issue_kinds = {item.get("kind") for item in (front_matter.get("issues") or [])}
    front_matter_rebuild_kinds = {"front_matter_order_mismatch", "toc_after_abstract"}

    if "mixed_manual_and_auto_heading_numbering" in heading_issue_kinds or "numbering_anomaly" in heading_issue_kinds:
        blocked.append("rewrite_heading_numbering_tree")
    if "mixed_heading_numbering_families" in heading_issue_kinds:
        blocked.append("normalize_mixed_numbering_families")
    if bibliography.get("issues"):
        blocked.append("convert_bibliography_citation_mode")
        blocked.append("rewrite_body_citations_to_author_year")
    if caption_audit.get("crossReferenceUnresolved"):
        blocked.append("rewrite_cross_references_as_word_fields")
    if front_issue_kinds & front_matter_rebuild_kinds:
        blocked.append("rebuild_front_matter_from_template")
    if section_drift and any(item.get("kind") in {"many_sections", "section_count_mismatch"} for item in section_drift):
        blocked.append("rewrite_complex_section_header_footer_inheritance")
    if source_kind == "docx" and reference_indent.get("issues") and not bibliography.get("issues"):
        blocked.append("batch_reference_indent_cleanup_without_confirmation")
    if document_risk_class == "C":
        blocked.extend(
            [
                "automatic_high_impact_repair",
                "automatic_template_region_rebuild",
            ]
        )

    deduped: List[str] = []
    for item in blocked:
        if item not in deduped:
            deduped.append(item)
    return deduped

def classify_docx_risk(
    audit: Dict[str, Any],
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    bibliography: Dict[str, Any],
    caption_audit: Dict[str, Any],
    reference_indent: Dict[str, Any],
    style_drift: List[Dict[str, Any]],
    numbering_drift: List[Dict[str, Any]],
    section_drift: List[Dict[str, Any]],
    confirmation_requests: List[Dict[str, Any]],
    manual_review: List[Dict[str, Any]],
) -> Dict[str, Any]:
    similarity = compute_template_similarity(
        audit,
        front_matter,
        heading_audit,
        caption_audit,
        style_drift,
        numbering_drift,
        section_drift,
    )
    template_matched = bool(similarity is not None and similarity >= 0.8)
    high_manual_review_count = sum(1 for item in manual_review if (item.get("details") or {}).get("severity") == "high")
    suspicious_runs = len(((audit.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or []))
    mixed_manual_and_auto = len(((audit.get("numberingAnalysis") or {}).get("mixedManualAndAuto") or []))
    numbering_anomalies = len(((audit.get("numberingAnalysis") or {}).get("anomalies") or []))
    unresolved_cross_refs = len(caption_audit.get("crossReferenceUnresolved") or [])
    section_count = len(audit.get("sections") or [])
    risk_reasons: List[Dict[str, Any]] = []

    if similarity is not None and similarity < 0.55:
        risk_reasons.append(
            {
                "kind": "low_template_similarity",
                "severity": "high",
                "message": "文档与模板的整体指纹相似度偏低。",
                "templateSimilarity": similarity,
            }
        )
    if mixed_manual_and_auto:
        risk_reasons.append(
            {
                "kind": "mixed_manual_and_auto_numbering",
                "severity": "high",
                "message": "标题区域存在手打编号与自动编号混用。",
                "count": mixed_manual_and_auto,
            }
        )
    if numbering_anomalies >= 6:
        risk_reasons.append(
            {
                "kind": "many_numbering_anomalies",
                "severity": "high",
                "message": "标题编号异常较多，自动重写风险较高。",
                "count": numbering_anomalies,
            }
        )
    if section_count >= 6:
        risk_reasons.append(
            {
                "kind": "many_sections",
                "severity": "high",
                "message": "文档 section 数量较多，页眉页脚继承可能复杂。",
                "count": section_count,
            }
        )
    if suspicious_runs >= 25:
        risk_reasons.append(
            {
                "kind": "many_font_slot_risks",
                "severity": "medium",
                "message": "可疑字体槽位运行较多，说明样式体系可能漂移明显。",
                "count": suspicious_runs,
            }
        )
    if unresolved_cross_refs >= 8:
        risk_reasons.append(
            {
                "kind": "many_unresolved_cross_references",
                "severity": "medium",
                "message": "图表交叉引用未解析项较多。",
                "count": unresolved_cross_refs,
            }
        )
    if bibliography.get("issues"):
        risk_reasons.append(
            {
                "kind": "bibliography_mode_mismatch",
                "kind": "bibliography_mode_mismatch",
                "severity": "medium",
                "count": len(bibliography.get("issues") or []),
            }
        )
    if high_manual_review_count >= 8:
        risk_reasons.append(
            {
                "kind": "many_high_severity_manual_reviews",
                "severity": "high",
                "message": "高严重度人工复核项较多。",
                "count": high_manual_review_count,
            }
        )
    if len(confirmation_requests) >= 3:
        risk_reasons.append(
            {
                "kind": "many_confirmation_first_operations",
                "severity": "medium",
                "message": "需要用户确认的高影响操作较多。",
                "count": len(confirmation_requests),
            }
        )

    severe = any(item.get("severity") == "high" for item in risk_reasons)
    if severe:
        document_risk_class = "C"
    elif template_matched and (similarity or 0.0) >= 0.82 and len(numbering_drift) <= 1 and len(section_drift) <= 1:
        document_risk_class = "A"
    else:
        document_risk_class = "B"

    recommended_mode = "audit-only" if document_risk_class == "C" else "conservative-repair"
    blocked_auto_repairs = build_blocked_auto_repairs(
        "docx",
        front_matter,
        heading_audit,
        bibliography,
        caption_audit,
        reference_indent,
        document_risk_class=document_risk_class,
        section_drift=section_drift,
    )
    return {
        "documentRiskClass": document_risk_class,
        "riskReasons": risk_reasons,
        "recommendedMode": recommended_mode,
        "blockedAutoRepairs": blocked_auto_repairs,
        "templateSimilarity": similarity,
        "templateFingerprintMatched": template_matched,
        "styleDrift": style_drift,
        "numberingDrift": numbering_drift,
        "sectionDrift": section_drift,
    }


def classify_text_risk(
    source: Path,
    front_matter: Dict[str, Any],
    heading_audit: Dict[str, Any],
    bibliography: Dict[str, Any],
    caption_audit: Dict[str, Any],
    reference_indent: Dict[str, Any],
    confirmation_requests: List[Dict[str, Any]],
    manual_review: List[Dict[str, Any]],
) -> Dict[str, Any]:
    risk_reasons: List[Dict[str, Any]] = []
    style_drift: List[Dict[str, Any]] = []
    numbering_drift: List[Dict[str, Any]] = []
    section_drift: List[Dict[str, Any]] = []
    suffix = source.suffix.lower()

    if front_matter.get("issues"):
        risk_reasons.append(
            {
                "kind": "front_matter_structure_mismatch",
                "severity": "medium",
                "message": "文本前置部分结构还不满足模板生成要求。",
                "count": len(front_matter.get("issues") or []),
            }
        )
    if heading_audit.get("issues"):
        numbering_drift.append(
            {
                "kind": "text_heading_structure_issues",
                "count": len(heading_audit.get("issues") or []),
                "examples": (heading_audit.get("issues") or [])[:20],
            }
        )
    if bibliography.get("issues"):
        risk_reasons.append(
            {
                "kind": "bibliography_mode_mismatch",
                "kind": "bibliography_mode_mismatch",
                "severity": "medium",
                "count": len(bibliography.get("issues") or []),
            }
        )
    if caption_audit.get("crossReferenceUnresolved"):
        risk_reasons.append(
            {
                "kind": "unresolved_cross_references",
                "severity": "medium",
                "message": "文本中存在无法直接定位的图表交叉引用。",
                "count": len(caption_audit.get("crossReferenceUnresolved") or []),
            }
        )
    if len(confirmation_requests) >= 3 or len(manual_review) >= 12:
        risk_reasons.append(
            {
                "kind": "many_confirmation_or_review_items",
                "severity": "high" if len(manual_review) >= 20 else "medium",
                "message": "文本在生成前仍有较多需要确认或人工复核的结构问题。",
                "confirmationCount": len(confirmation_requests),
                "manualReviewCount": len(manual_review),
            }
        )

    severe = any(item.get("severity") == "high" for item in risk_reasons)
    if severe:
        document_risk_class = "C"
        recommended_mode = "audit-only"
    else:
        document_risk_class = "B"
        recommended_mode = "rebuild" if suffix in {".md", ".markdown", ".txt"} else "conservative-repair"

    blocked_auto_repairs = build_blocked_auto_repairs(
        suffix.lstrip("."),
        front_matter,
        heading_audit,
        bibliography,
        caption_audit,
        reference_indent,
        document_risk_class=document_risk_class,
        section_drift=section_drift,
    )
    return {
        "documentRiskClass": document_risk_class,
        "riskReasons": risk_reasons,
        "recommendedMode": recommended_mode,
        "blockedAutoRepairs": blocked_auto_repairs,
        "templateSimilarity": None,
        "templateFingerprintMatched": False,
        "styleDrift": style_drift,
        "numberingDrift": numbering_drift,
        "sectionDrift": section_drift,
    }


def analyze_docx(source: Path, rules: Dict[str, Any], template_docx: Optional[str], rules_yaml: str) -> Dict[str, Any]:
    audit = parse_document(str(source), template_docx=template_docx, rules_path=rules_yaml)
    split = split_references_from_docx(audit)
    bibliography = {
        "referencesHeadingIndex": split.get("referencesHeadingIndex"),
        **audit_blocks(split["bodyParagraphs"], split["bibliographyEntries"], rules),
    }
    front_matter = analyze_front_matter_docx(audit, rules)
    heading_audit = analyze_heading_docx(audit)
    caption_audit = analyze_caption_docx(audit)
    reference_indent = analyze_reference_indentation_docx(audit, rules)

    actions = build_actions("docx", front_matter, heading_audit, bibliography, caption_audit, reference_indent)
    confirmation_requests = build_confirmation_requests("docx", front_matter, heading_audit, bibliography, caption_audit, reference_indent)
    manual_review = collect_manual_review(front_matter, heading_audit, bibliography, caption_audit, reference_indent)
    style_drift = build_style_drift_docx(audit)
    numbering_drift = build_numbering_drift_docx(audit, heading_audit)
    section_drift = build_section_drift_docx(audit)
    risk = classify_docx_risk(
        audit,
        front_matter,
        heading_audit,
        bibliography,
        caption_audit,
        reference_indent,
        style_drift,
        numbering_drift,
        section_drift,
        confirmation_requests,
        manual_review,
    )
    blocking_count = sum(1 for item in manual_review if item.get("details", {}).get("severity") == "high")

    return {
        "generatedAt": now_iso(),
        "source": str(source.resolve()),
        "sourceKind": "docx",
        "pipelineStage": "preflight_semantic_normalization",
        "templateDocx": template_docx,
        "rulesPath": rules_yaml,
        "documentSummary": audit.get("summary"),
        "frontMatter": front_matter,
        "headings": heading_audit,
        "bibliography": bibliography,
        "referenceIndentation": reference_indent,
        "captions": caption_audit,
        "actions": actions,
        "confirmationRequests": confirmation_requests,
        "manualReview": manual_review,
        **risk,
        "readiness": {
            "canProceedToTypesetting": blocking_count == 0,
            "blockingIssueCount": blocking_count,
            "nextStep": "repair_planning_and_typesetting" if blocking_count == 0 else "normalize_structure_before_typesetting",
            "requiresUserConfirmation": bool(confirmation_requests),
        },
        "notes": [
            "DOCX preflight is OOXML-backed and should run before typography normalization.",
            "It is acceptable to preserve the first two pages from the validated template instead of rebuilding them.",
            "When confirmationRequests is non-empty, high-impact rewrites should be explained to the user before execution.",
        ],
    }


def analyze_text(source: Path, rules: Dict[str, Any], thesis_ir: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    evidence = extract_text_evidence(source)
    thesis_ir = thesis_ir or build_thesis_ir(evidence)
    reference_entries = [
        {"index": item.get("sourceIndex"), "text": item.get("text")}
        for item in (thesis_ir.get("references", {}).get("entries") or [])
        if item.get("text")
    ]
    body_blocks = [
        {"index": item.get("sourceIndex"), "text": item.get("text")}
        for item in (thesis_ir.get("bodyBlocks") or [])
        if item.get("text")
    ]
    split = split_references_from_text(source)
    bibliography = {
        "referencesHeadingIndex": thesis_ir.get("references", {}).get("headingSourceIndex", split.get("referencesHeadingIndex")),
        **audit_blocks(body_blocks, reference_entries, rules),
    }
    front_matter = analyze_front_matter_text_from_ir(source, rules, evidence, thesis_ir)
    heading_audit = analyze_heading_blocks_from_thesis_ir(thesis_ir)
    caption_audit = analyze_caption_text(source)
    reference_indent = analyze_reference_indentation_text(source, rules)

    actions = build_actions(source.suffix.lower().lstrip("."), front_matter, heading_audit, bibliography, caption_audit, reference_indent)
    confirmation_requests = build_confirmation_requests(source.suffix.lower().lstrip("."), front_matter, heading_audit, bibliography, caption_audit, reference_indent)
    manual_review = collect_manual_review(front_matter, heading_audit, bibliography, caption_audit, reference_indent)
    risk = classify_text_risk(
        source,
        front_matter,
        heading_audit,
        bibliography,
        caption_audit,
        reference_indent,
        confirmation_requests,
        manual_review,
    )
    blocking_count = sum(1 for item in manual_review if item.get("details", {}).get("severity") == "high")

    return {
        "generatedAt": now_iso(),
        "source": str(source.resolve()),
        "sourceKind": source.suffix.lower().lstrip("."),
        "pipelineStage": "preflight_semantic_normalization",
        "rulesPath": None,
        "documentSummary": extract_text_ir(source, max_blocks=None).get("documentSummary"),
        "sourceEvidenceSummary": thesis_ir.get("evidenceSummary"),
        "thesisIrConfidence": thesis_ir.get("confidence"),
        "frontMatter": front_matter,
        "headings": heading_audit,
        "bibliography": bibliography,
        "referenceIndentation": reference_indent,
        "captions": caption_audit,
        "actions": actions,
        "confirmationRequests": confirmation_requests,
        "manualReview": manual_review,
        **risk,
        "readiness": {
            "canProceedToBuildAndTypesetting": blocking_count == 0,
            "blockingIssueCount": blocking_count,
            "nextStep": "build_source_docx_then_typesetting" if blocking_count == 0 else "normalize_text_structure_before_build",
            "requiresUserConfirmation": bool(confirmation_requests),
        },
        "notes": [
            "Text-first inputs should be normalized before building the source DOCX.",
            "This stage focuses on semantics and ordering, not final Word metrics.",
            "When confirmationRequests is non-empty, high-impact rewrites should be explained to the user before execution.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Run semantic/structural preflight checks before thesis typesetting.")
    parser.add_argument("source", help="Input .docx, .md, or .txt")
    parser.add_argument("--template-docx", help="Optional template DOCX for DOCX preflight")
    parser.add_argument("--rules-yaml", default="references/zafu_2022_rules.yaml", help="Rules YAML")
    parser.add_argument("--style-map-yaml", help="Optional profile style-role mapping YAML")
    parser.add_argument("--front-matter-policy-yaml", help="Optional profile front-matter policy YAML")
    parser.add_argument("--validators-yaml", help="Optional profile validator selection YAML")
    parser.add_argument("--output", "-o", help="Output JSON path")
    parser.add_argument("--thesis-ir-json", help="Optional unified ThesisIR v2 produced by the ingest stage")
    args = parser.parse_args()

    source = Path(args.source)
    rules = load_rules(args.rules_yaml)
    suffix = source.suffix.lower()
    thesis_ir = json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8")) if args.thesis_ir_json else None

    if suffix == ".docx":
        data = analyze_docx(source, rules, args.template_docx, args.rules_yaml)
    elif suffix in {".md", ".markdown", ".txt"}:
        data = analyze_text(source, rules, thesis_ir=thesis_ir)
        data["rulesPath"] = args.rules_yaml
    else:
        raise SystemExit(f"Unsupported source type: {suffix}")

    if thesis_ir:
        data["semanticModel"] = {
            "schema": thesis_ir.get("schema"),
            "version": thesis_ir.get("version"),
            "sourceFormat": (thesis_ir.get("source") or {}).get("format"),
            "semanticBlockCount": thesis_ir.get("semanticBlockCount"),
            "consumedByPreflight": suffix != ".docx",
            "docxEvidenceSidecarRequired": suffix == ".docx",
        }
    write_json(data, args.output)


if __name__ == "__main__":
    main()




