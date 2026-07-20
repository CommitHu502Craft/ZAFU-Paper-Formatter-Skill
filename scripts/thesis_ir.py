#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from collections import Counter
from typing import Any, Dict, List, Optional

from recover_numbering import analyze_candidates, extract_items_from_evidence


KEYWORDS_SPLIT_RE = re.compile(r"[:：]\s*")
ABSTRACT_HEADING_SET = {"摘要", "摘  要", "中文摘要"}
REFERENCES_HEADING_RE = re.compile(r"^\s*(参考文献|References)(?:[（(].*?[)）])?\s*$")
ACKNOWLEDGEMENTS_HEADING_SET = {"致谢", "致 谢", "Acknowledgements", "Acknowledgment"}
NUMERIC_REFERENCE_LABEL_RE = re.compile(r"^\[\s*(\d+)\s*\]\s*")
NUMERIC_CITATION_RE = re.compile(r"\[\s*\d+(?:\s*[-,，]\s*\d+)*\s*\]")
SCIENCE_DECIMAL_LEVEL1_RE = re.compile(r"^\d+\s+")
CAPTION_LABEL_RE = re.compile(
    r"^(?:(?P<cnKind>图|表)\s*(?P<cnLabel>[0-9一二三四五六七八九十]+(?:[.．\-—][0-9一二三四五六七八九十]+)?)|"
    r"(?P<enKind>Figure|Fig\.?|Table)\s*(?P<enLabel>[0-9A-Za-z]+(?:[.．\-—][0-9A-Za-z]+)?))",
    re.IGNORECASE,
)


def normalize_text(value: Optional[str]) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def extract_front_matter_from_text(evidence: Dict[str, Any]) -> Dict[str, Any]:
    lines = evidence.get("lineEvidence") or []
    title = None
    abstract_cn = None
    keywords_cn = None
    source_blocks: List[int] = []
    in_abstract_cn = False
    abstract_parts: List[str] = []
    for item in lines[:80]:
        text = normalize_text(item.get("normalizedText") or item.get("text"))
        if not text:
            continue
        if title is None and not any(item.get("candidateLabels", {}).values()):
            title = text
            source_blocks.append(int(item.get("index")))
            continue
        if item.get("candidateLabels", {}).get("abstract") and abstract_cn is None:
            label, _, content = text.partition("：")
            if not content:
                label, _, content = text.partition(":")
            if text in ABSTRACT_HEADING_SET or (label == text and not content):
                in_abstract_cn = True
                source_blocks.append(int(item.get("index")))
                continue
            abstract_cn = normalize_text(content or label)
            source_blocks.append(int(item.get("index")))
            continue
        if item.get("candidateLabels", {}).get("keywords") and keywords_cn is None:
            parts = KEYWORDS_SPLIT_RE.split(text, maxsplit=1)
            keywords_cn = normalize_text(parts[1] if len(parts) > 1 else text)
            source_blocks.append(int(item.get("index")))
            in_abstract_cn = False
            continue
        if REFERENCES_HEADING_RE.match(text) or item.get("candidateHeading"):
            in_abstract_cn = False
        if in_abstract_cn and keywords_cn is None and not item.get("candidateHeading"):
            abstract_parts.append(text)
            source_blocks.append(int(item.get("index")))
            continue
        if abstract_cn is not None and keywords_cn is None and not item.get("candidateHeading"):
            abstract_cn = normalize_text(f"{abstract_cn} {text}")
            source_blocks.append(int(item.get("index")))
    if abstract_parts:
        abstract_cn = normalize_text(" ".join(abstract_parts))
    confidence = 0.45
    if abstract_cn:
        confidence += 0.25
    if keywords_cn:
        confidence += 0.15
    if title:
        confidence += 0.1
    return {
        "title": title,
        "abstractCn": abstract_cn,
        "abstractEn": None,
        "keywordsCn": keywords_cn,
        "keywordsEn": None,
        "sourceBlocks": source_blocks,
        "confidence": round(min(confidence, 1.0), 2),
    }


def extract_front_matter_from_docx(evidence: Dict[str, Any]) -> Dict[str, Any]:
    paragraphs = evidence.get("paragraphEvidence") or []
    title = None
    abstract_cn = None
    keywords_cn = None
    source_blocks: List[int] = []
    for item in paragraphs[:120]:
        text = normalize_text(item.get("text"))
        if not text:
            continue
        if title is None and len(text) <= 120 and not item.get("candidateHeading") and not any(item.get("candidateLabels", {}).values()):
            title = text
            source_blocks.append(int(item.get("index")))
            continue
        if item.get("candidateLabels", {}).get("abstract") and abstract_cn is None:
            label, _, content = text.partition("：")
            if not content:
                label, _, content = text.partition(":")
            abstract_cn = normalize_text(content or label)
            source_blocks.append(int(item.get("index")))
            continue
        if item.get("candidateLabels", {}).get("keywords") and keywords_cn is None:
            parts = KEYWORDS_SPLIT_RE.split(text, maxsplit=1)
            keywords_cn = normalize_text(parts[1] if len(parts) > 1 else text)
            source_blocks.append(int(item.get("index")))
            continue
    confidence = 0.55
    if abstract_cn:
        confidence += 0.2
    if keywords_cn:
        confidence += 0.1
    if title:
        confidence += 0.1
    return {
        "title": title,
        "abstractCn": abstract_cn,
        "abstractEn": None,
        "keywordsCn": keywords_cn,
        "keywordsEn": None,
        "sourceBlocks": source_blocks,
        "confidence": round(min(confidence, 1.0), 2),
    }


def build_heading_tree(numbering_analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
    heading_tree: List[Dict[str, Any]] = []
    for item in numbering_analysis.get("candidates") or []:
        heading_tree.append(
            {
                "text": item.get("text"),
                "level": item.get("level"),
                "numberingFamily": item.get("family"),
                "sourceIndex": item.get("sourceIndex"),
                "confidence": numbering_analysis.get("confidence"),
            }
        )
    return heading_tree


def _source_item_map(evidence: Dict[str, Any]) -> Dict[int, Dict[str, Any]]:
    items = iter_source_items(evidence)
    mapped: Dict[int, Dict[str, Any]] = {}
    for item in items:
        source_index = item.get("index")
        if source_index is None:
            continue
        mapped[int(source_index)] = item
    return mapped


def _looks_like_body_enumeration(candidate: Dict[str, Any], source_item: Optional[Dict[str, Any]]) -> bool:
    text = normalize_text(candidate.get("text"))
    if not text:
        return False
    family = str(candidate.get("family") or "")
    level = int(candidate.get("level") or 0)

    role_name = str((source_item or {}).get("roleName") or "")
    if family == "science_decimal" and level == 1:
        if not SCIENCE_DECIMAL_LEVEL1_RE.match(text):
            return False
        if role_name in {"body", "body_enumeration"}:
            return True
        if text.endswith(("。", "；", ";", "！", "？", "!", "?")):
            return True
        if len(text) >= 28:
            return True
        return False

    if family == "arabic_list" and level == 3:
        if role_name in {"body", "body_enumeration"}:
            return True
        if text.endswith(("。", "；", ";", "！", "？", "!", "?")):
            return True
        if len(text) >= 20:
            return True
        return False

    return False


def build_filtered_heading_tree(evidence: Dict[str, Any], numbering_analysis: Dict[str, Any]) -> List[Dict[str, Any]]:
    source_items = _source_item_map(evidence)
    heading_tree: List[Dict[str, Any]] = []
    for item in numbering_analysis.get("candidates") or []:
        source_index = item.get("sourceIndex")
        source_item = source_items.get(int(source_index)) if source_index is not None else None
        if _looks_like_body_enumeration(item, source_item):
            continue
        heading_tree.append(
            {
                "text": item.get("text"),
                "level": item.get("level"),
                "numberingFamily": item.get("family"),
                "sourceIndex": item.get("sourceIndex"),
                "confidence": numbering_analysis.get("confidence"),
            }
        )
    return heading_tree


def iter_source_items(evidence: Dict[str, Any]) -> List[Dict[str, Any]]:
    if evidence.get("sourceType") == "docx":
        return list(evidence.get("paragraphEvidence") or [])
    return list(evidence.get("lineEvidence") or [])


def is_heading_like(item: Dict[str, Any]) -> bool:
    labels = item.get("candidateLabels", {}) or {}
    role_name = str(item.get("roleName") or "")
    return bool(
        item.get("candidateHeading")
        or labels.get("references")
        or labels.get("acknowledgements")
        or labels.get("appendix")
        or role_name.startswith("heading")
        or role_name in {"ack_heading", "appendix_heading"}
    )


def classify_caption(text: str, role_name: Optional[str]) -> Dict[str, Any]:
    normalized = normalize_text(text)
    match = CAPTION_LABEL_RE.match(normalized)
    caption_type = "unknown"
    caption_label = None
    if role_name == "table_caption":
        caption_type = "table"
    elif role_name == "figure_caption":
        caption_type = "figure"
    elif match:
        if match.group("cnKind") == "表" or (match.group("enKind") or "").lower() == "table":
            caption_type = "table"
            caption_label = match.group("cnLabel") or match.group("enLabel")
        else:
            caption_type = "figure"
            caption_label = match.group("cnLabel") or match.group("enLabel")
    return {"captionType": caption_type, "label": caption_label}


def extract_caption_blocks(evidence: Dict[str, Any]) -> List[Dict[str, Any]]:
    items = iter_source_items(evidence)
    captions: List[Dict[str, Any]] = []
    for position, item in enumerate(items):
        labels = item.get("candidateLabels", {}) or {}
        role_name = str(item.get("roleName") or "")
        if not labels.get("caption") and role_name not in {"figure_caption", "table_caption"}:
            continue
        text = normalize_text(item.get("text"))
        if not text:
            continue
        caption_info = classify_caption(text, role_name)
        previous_non_blank = None
        next_non_blank = None
        for candidate in reversed(items[:position]):
            candidate_text = normalize_text(candidate.get("text"))
            if candidate_text:
                previous_non_blank = {"sourceIndex": candidate.get("index"), "textPreview": candidate_text[:120]}
                break
        for candidate in items[position + 1 :]:
            candidate_text = normalize_text(candidate.get("text"))
            if candidate_text:
                next_non_blank = {"sourceIndex": candidate.get("index"), "textPreview": candidate_text[:120]}
                break
        captions.append(
            {
                "sourceIndex": item.get("index"),
                "text": text,
                "captionType": caption_info["captionType"],
                "label": caption_info["label"],
                "anchorHint": {
                    "precedingSourceIndex": previous_non_blank.get("sourceIndex") if previous_non_blank else None,
                    "followingSourceIndex": next_non_blank.get("sourceIndex") if next_non_blank else None,
                    "precedingTextPreview": previous_non_blank.get("textPreview") if previous_non_blank else None,
                    "followingTextPreview": next_non_blank.get("textPreview") if next_non_blank else None,
                },
                "confidence": round(float(item.get("roleConfidence") or 0.78), 2) if role_name else 0.72,
            }
        )
    return captions[:160]


def _collect_section_blocks(
    items: List[Dict[str, Any]],
    start_position: int,
    stop_predicate,
) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    for item in items[start_position + 1 :]:
        if stop_predicate(item):
            break
        text = normalize_text(item.get("text"))
        if not text:
            continue
        blocks.append({"kind": "paragraph", "sourceIndex": item.get("index"), "text": text})
    return blocks


def extract_acknowledgements(evidence: Dict[str, Any]) -> Dict[str, Any]:
    items = iter_source_items(evidence)
    for position, item in enumerate(items):
        labels = item.get("candidateLabels", {}) or {}
        role_name = str(item.get("roleName") or "")
        text = normalize_text(item.get("text"))
        if not (
            labels.get("acknowledgements")
            or role_name == "ack_heading"
            or text in ACKNOWLEDGEMENTS_HEADING_SET
        ):
            continue
        blocks = _collect_section_blocks(
            items,
            position,
            lambda candidate: bool(
                bool(REFERENCES_HEADING_RE.match(normalize_text(candidate.get("text"))))
                or (candidate.get("candidateLabels", {}) or {}).get("appendix")
                or (candidate.get("candidateLabels", {}) or {}).get("references")
                or is_heading_like(candidate)
            ),
        )
        confidence = 0.95 if role_name == "ack_heading" else 0.78
        return {
            "heading": text or "致谢",
            "headingSourceIndex": item.get("index"),
            "blocks": blocks[:80],
            "confidence": round(confidence if blocks else max(confidence - 0.12, 0.55), 2),
        }
    return {
        "heading": None,
        "headingSourceIndex": None,
        "blocks": [],
        "confidence": 0.0,
    }


def extract_appendix(evidence: Dict[str, Any]) -> Dict[str, Any]:
    items = iter_source_items(evidence)
    sections: List[Dict[str, Any]] = []
    appendix_positions: List[int] = []
    for position, item in enumerate(items):
        labels = item.get("candidateLabels", {}) or {}
        role_name = str(item.get("roleName") or "")
        if labels.get("appendix") or role_name == "appendix_heading":
            appendix_positions.append(position)
    for offset, position in enumerate(appendix_positions):
        item = items[position]
        next_position = appendix_positions[offset + 1] if offset + 1 < len(appendix_positions) else None
        heading_text = normalize_text(item.get("text"))
        blocks: List[Dict[str, Any]] = []
        for candidate in items[position + 1 : next_position]:
            labels = candidate.get("candidateLabels", {}) or {}
            candidate_text = normalize_text(candidate.get("text"))
            if not candidate_text:
                continue
            if labels.get("references") or REFERENCES_HEADING_RE.match(candidate_text):
                break
            if labels.get("acknowledgements"):
                break
            blocks.append({"kind": "paragraph", "sourceIndex": candidate.get("index"), "text": candidate_text})
        sections.append(
            {
                "heading": heading_text,
                "headingSourceIndex": item.get("index"),
                "blocks": blocks[:120],
                "confidence": round(0.94 if str(item.get("roleName") or "") == "appendix_heading" else 0.76, 2),
            }
        )
    return {
        "sections": sections[:40],
        "confidence": round(max((section.get("confidence") or 0.0) for section in sections), 2) if sections else 0.0,
    }


def extract_references_section(evidence: Dict[str, Any]) -> Dict[str, Any]:
    heading_index: Optional[int] = None
    entries: List[Dict[str, Any]] = []
    if evidence.get("sourceType") == "docx":
        items = evidence.get("paragraphEvidence") or []
    else:
        items = evidence.get("lineEvidence") or []
    for item in items:
        index = item.get("index")
        text = normalize_text(item.get("text"))
        if not text:
            continue
        if heading_index is None and REFERENCES_HEADING_RE.match(text):
            heading_index = index
            continue
        if heading_index is not None and index is not None and int(index) > int(heading_index):
            entries.append({"kind": "reference_entry", "sourceIndex": index, "text": text})
    return {
        "heading": "参考文献" if heading_index is not None else None,
        "headingSourceIndex": heading_index,
        "entries": entries,
    }


def build_asset_anchor_ambiguities(
    evidence: Dict[str, Any],
    caption_blocks: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    preserveable_assets = evidence.get("preserveableAssets") or {}
    hints = evidence.get("assetAnchorHints") or {}
    ambiguities: List[Dict[str, Any]] = []

    image_count = int(preserveable_assets.get("images") or 0)
    table_count = int(preserveable_assets.get("tables") or 0)
    if image_count or table_count:
        ambiguities.append(
            {
                "kind": "preserved_assets_require_anchor_resolution",
                "message": "Preserveable DOCX assets exist and need explicit anchor resolution before hybrid rebuild.",
                "images": image_count,
                "tables": table_count,
                "captions": len(caption_blocks),
            }
        )

    for issue in (hints.get("captionIssues") or [])[:40]:
        issue_kind = issue.get("kind") or "caption_anchor_issue"
        ambiguities.append(
            {
                "kind": f"asset_anchor_{issue_kind}",
                "message": "Caption-to-asset placement is not stable enough for blind preservation.",
                "sourceIndex": issue.get("paragraphIndex"),
                "role": issue.get("role"),
                "textPreview": normalize_text(issue.get("text"))[:160] if issue.get("text") else None,
            }
        )

    unresolved_cross_refs = [item for item in (hints.get("crossReferences") or []) if not item.get("resolved")]
    if unresolved_cross_refs:
        ambiguities.append(
            {
                "kind": "unresolved_caption_cross_references",
                "message": "Cross references to figures/tables exist but are not fully resolved at evidence stage.",
                "count": len(unresolved_cross_refs),
                "examples": [
                    {
                        "sourceIndex": item.get("paragraphIndex"),
                        "label": item.get("label"),
                        "textPreview": normalize_text(item.get("text"))[:120] if item.get("text") else None,
                    }
                    for item in unresolved_cross_refs[:10]
                ],
            }
        )

    if caption_blocks and (image_count + table_count) > len(caption_blocks):
        ambiguities.append(
            {
                "kind": "asset_caption_count_mismatch",
                "message": "The number of preserveable assets exceeds the number of recovered captions.",
                "assetCount": image_count + table_count,
                "captionCount": len(caption_blocks),
            }
        )

    return ambiguities[:80]


def _section_index_for_source(source_index: Optional[int], sections: List[Dict[str, Any]]) -> Optional[int]:
    if source_index is None or not sections:
        return None
    selected = None
    for idx, section in enumerate(sections):
        paragraph_index = section.get("paragraphIndex")
        if paragraph_index is None:
            continue
        if int(source_index) >= int(paragraph_index):
            selected = idx
        else:
            break
    return selected


def _nearest_caption_candidate(
    source_index: Optional[int],
    caption_blocks: List[Dict[str, Any]],
    caption_type: str,
    *,
    prefer_following: bool,
) -> Optional[Dict[str, Any]]:
    if source_index is None:
        return None
    matches = [item for item in caption_blocks if item.get("captionType") == caption_type and item.get("sourceIndex") is not None]
    if not matches:
        return None
    if prefer_following:
        following = [item for item in matches if int(item["sourceIndex"]) >= int(source_index)]
        if following:
            return min(following, key=lambda item: abs(int(item["sourceIndex"]) - int(source_index)))
    preceding = [item for item in matches if int(item["sourceIndex"]) <= int(source_index)]
    if preceding:
        return min(preceding, key=lambda item: abs(int(item["sourceIndex"]) - int(source_index)))
    return min(matches, key=lambda item: abs(int(item["sourceIndex"]) - int(source_index)))


def _infer_body_start_index(
    items: List[Dict[str, Any]],
    front_matter: Dict[str, Any],
    heading_tree: List[Dict[str, Any]],
) -> Optional[int]:
    heading_indexes = [
        int(item["sourceIndex"])
        for item in heading_tree
        if item.get("sourceIndex") is not None
    ]
    if heading_indexes:
        return min(heading_indexes)
    source_blocks = [
        int(item)
        for item in (front_matter.get("sourceBlocks") or [])
        if item is not None
    ]
    if source_blocks:
        return max(source_blocks) + 1
    for item in items:
        if item.get("candidateHeading"):
            source_index = item.get("index")
            if source_index is not None:
                return int(source_index)
    return None


def _candidate_source_region(source_index: Optional[int], body_start_index: Optional[int]) -> str:
    if source_index is None or body_start_index is None:
        return "unknown"
    if int(source_index) < int(body_start_index):
        return "front_matter"
    return "body"


def _table_anchor_paragraph_index(table_index: int, body_sequence: List[Dict[str, Any]]) -> Optional[int]:
    table_position = None
    for position, item in enumerate(body_sequence):
        if item.get("kind") == "table" and int(item.get("tableIndex") or -1) == int(table_index):
            table_position = position
            break
    if table_position is None:
        return None
    for candidate in body_sequence[table_position + 1 :]:
        if candidate.get("kind") == "paragraph" and candidate.get("paragraphIndex") is not None:
            return int(candidate.get("paragraphIndex"))
    for candidate in reversed(body_sequence[:table_position]):
        if candidate.get("kind") == "paragraph" and candidate.get("paragraphIndex") is not None:
            return int(candidate.get("paragraphIndex"))
    return None


def build_attachable_asset_candidates(
    evidence: Dict[str, Any],
    caption_blocks: List[Dict[str, Any]],
    front_matter: Dict[str, Any],
    heading_tree: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    if evidence.get("sourceType") != "docx":
        return []

    sections = evidence.get("sections") or []
    items = evidence.get("paragraphEvidence") or []
    body_sequence = ((evidence.get("assetAnchorHints") or {}).get("bodySequencePreview") or [])
    body_start_index = _infer_body_start_index(items, front_matter, heading_tree)
    candidates: List[Dict[str, Any]] = []

    for paragraph in items:
        if not paragraph.get("hasDrawing"):
            continue
        source_index = paragraph.get("index")
        source_region = _candidate_source_region(source_index, body_start_index)
        caption = _nearest_caption_candidate(source_index, caption_blocks, "figure", prefer_following=True)
        distance = abs(int(caption.get("sourceIndex")) - int(source_index)) if caption and caption.get("sourceIndex") is not None else None
        confidence = 0.48
        decision_reason = "caption_distance_exceeds_threshold"
        recommended_action = "manual_review"
        if source_region == "front_matter":
            confidence = 0.95
            recommended_action = "skip"
            decision_reason = "pre_body_asset_outside_hybrid_scope"
        elif distance is None:
            decision_reason = "caption_not_recovered"
        else:
            if distance <= 1:
                confidence = 0.9
                decision_reason = "caption_neighbor_high_confidence"
            elif distance <= 3:
                confidence = 0.76
                decision_reason = "caption_neighbor_medium_confidence"
            elif distance <= 6:
                confidence = 0.62
                decision_reason = "caption_neighbor_low_confidence"
            recommended_action = "reattach_candidate" if confidence >= 0.7 else "manual_review"
        candidates.append(
            {
                "assetKind": "image",
                "sourceIndex": source_index,
                "drawingCount": paragraph.get("drawingCount"),
                "captionSourceIndex": caption.get("sourceIndex") if caption else None,
                "captionLabel": caption.get("label") if caption else None,
                "captionText": caption.get("text") if caption else None,
                "sectionIndex": _section_index_for_source(source_index, sections),
                "anchorMethod": "paragraph_neighbor",
                "sourceRegion": source_region,
                "distanceToCaption": distance,
                "confidence": round(confidence, 2),
                "recommendedAction": recommended_action,
                "decisionReason": decision_reason,
            }
        )

    table_positions: List[int] = []
    for item in body_sequence:
        if item.get("kind") == "table" and item.get("tableIndex") is not None:
            table_positions.append(int(item.get("tableIndex")))
    for table_index in table_positions:
        table_anchor_index = _table_anchor_paragraph_index(table_index, body_sequence)
        source_region = _candidate_source_region(table_anchor_index, body_start_index)
        caption = _nearest_caption_candidate(table_anchor_index, caption_blocks, "table", prefer_following=False)
        distance = (
            abs(int(caption.get("sourceIndex")) - int(table_anchor_index))
            if caption and caption.get("sourceIndex") is not None and table_anchor_index is not None
            else None
        )
        confidence = 0.52
        decision_reason = "caption_distance_exceeds_threshold"
        recommended_action = "manual_review"
        if source_region == "front_matter":
            confidence = 0.95
            recommended_action = "skip"
            decision_reason = "pre_body_asset_outside_hybrid_scope"
        elif distance is None:
            decision_reason = "caption_not_recovered"
        else:
            if distance <= 1:
                confidence = 0.92
                decision_reason = "caption_neighbor_high_confidence"
            elif distance <= 3:
                confidence = 0.78
                decision_reason = "caption_neighbor_medium_confidence"
            elif distance <= 6:
                confidence = 0.64
                decision_reason = "caption_neighbor_low_confidence"
            recommended_action = "reattach_candidate" if confidence >= 0.7 else "manual_review"
        candidates.append(
            {
                "assetKind": "table",
                "sourceIndex": table_index,
                "sourceAnchorIndex": table_anchor_index,
                "captionSourceIndex": caption.get("sourceIndex") if caption else None,
                "captionLabel": caption.get("label") if caption else None,
                "captionText": caption.get("text") if caption else None,
                "sectionIndex": _section_index_for_source(table_anchor_index, sections),
                "anchorMethod": "body_sequence_neighbor",
                "sourceRegion": source_region,
                "distanceToCaption": distance,
                "confidence": round(confidence, 2),
                "recommendedAction": recommended_action,
                "decisionReason": decision_reason,
            }
        )

    return candidates[:160]


def analyze_reference_normalization(references: Dict[str, Any], body_blocks: List[Dict[str, Any]]) -> Dict[str, Any]:
    entries = references.get("entries") or []
    numeric_label_count = 0
    normalized_entries_preview: List[Dict[str, Any]] = []
    for item in entries[:40]:
        text = normalize_text(item.get("text"))
        normalized_text = NUMERIC_REFERENCE_LABEL_RE.sub("", text, count=1).strip()
        has_numeric_label = bool(NUMERIC_REFERENCE_LABEL_RE.match(text))
        if has_numeric_label:
            numeric_label_count += 1
        normalized_entries_preview.append(
            {
                "sourceIndex": item.get("sourceIndex"),
                "hasNumericLabel": has_numeric_label,
                "textPreview": text[:160],
                "normalizedPreview": normalized_text[:160],
            }
        )

    if len(entries) > len(normalized_entries_preview):
        for item in entries[len(normalized_entries_preview):]:
            text = normalize_text(item.get("text"))
            if NUMERIC_REFERENCE_LABEL_RE.match(text):
                numeric_label_count += 1

    body_numeric_citation_count = 0
    body_numeric_citation_examples: List[Dict[str, Any]] = []
    for block in body_blocks:
        text = normalize_text(block.get("text"))
        if not text:
            continue
        if NUMERIC_CITATION_RE.search(text):
            body_numeric_citation_count += 1
            if len(body_numeric_citation_examples) < 20:
                body_numeric_citation_examples.append(
                    {
                        "sourceIndex": block.get("sourceIndex"),
                        "textPreview": text[:200],
                    }
                )

    return {
        "entryCount": len(entries),
        "numericLabelCount": numeric_label_count,
        "allEntriesHaveNumericLabels": bool(entries) and numeric_label_count == len(entries),
        "safeListLabelStripCandidate": bool(entries) and numeric_label_count == len(entries),
        "bodyNumericCitationCount": body_numeric_citation_count,
        "bodyNumericCitationExamples": body_numeric_citation_examples,
        "normalizedEntriesPreview": normalized_entries_preview,
    }


def build_body_blocks(
    evidence: Dict[str, Any],
    heading_tree: List[Dict[str, Any]],
    front_matter: Dict[str, Any],
    references: Dict[str, Any],
    acknowledgements: Dict[str, Any],
    appendix: Dict[str, Any],
    caption_blocks: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    heading_indexes = {item.get("sourceIndex") for item in heading_tree}
    front_indexes = set(front_matter.get("sourceBlocks") or [])
    reference_heading_index = references.get("headingSourceIndex")
    reference_entry_indexes = {item.get("sourceIndex") for item in references.get("entries") or []}
    acknowledgement_indexes = {
        acknowledgements.get("headingSourceIndex"),
        *[item.get("sourceIndex") for item in acknowledgements.get("blocks") or []],
    }
    appendix_indexes = {
        item.get("headingSourceIndex")
        for item in appendix.get("sections") or []
    }
    for section in appendix.get("sections") or []:
        appendix_indexes.update(block.get("sourceIndex") for block in section.get("blocks") or [])
    caption_indexes = {item.get("sourceIndex") for item in caption_blocks}
    blocks: List[Dict[str, Any]] = []
    if evidence.get("sourceType") == "docx":
        for paragraph in evidence.get("paragraphEvidence") or []:
            index = paragraph.get("index")
            if (
                index in heading_indexes
                or index in front_indexes
                or index == reference_heading_index
                or index in reference_entry_indexes
                or index in acknowledgement_indexes
                or index in appendix_indexes
                or index in caption_indexes
            ):
                continue
            text = normalize_text(paragraph.get("text"))
            if not text:
                continue
            blocks.append({"kind": "paragraph", "sourceIndex": index, "text": text})
    else:
        for line in evidence.get("lineEvidence") or []:
            index = line.get("index")
            if (
                index in heading_indexes
                or index in front_indexes
                or index == reference_heading_index
                or index in reference_entry_indexes
                or index in acknowledgement_indexes
                or index in appendix_indexes
                or index in caption_indexes
                or line.get("isBlank")
            ):
                continue
            text = normalize_text(line.get("text"))
            if not text:
                continue
            blocks.append({"kind": "paragraph", "sourceIndex": index, "text": text})
    return blocks[:300]


def choose_default_strategy(source_type: str, numbering_analysis: Dict[str, Any], front_matter: Dict[str, Any], preserveable_assets: Dict[str, Any]) -> str:
    if source_type in {"txt", "md", "markdown"}:
        return "text_rebuild"
    conflicts = numbering_analysis.get("conflicts") or []
    images = int(preserveable_assets.get("images") or 0)
    tables = int(preserveable_assets.get("tables") or 0)
    front_conf = float(front_matter.get("confidence") or 0.0)
    if conflicts and (images or tables):
        return "hybrid_rebuild"
    if front_conf < 0.45:
        return "audit_only"
    return "preserve_first"


def canonical_source_format(evidence: Dict[str, Any]) -> str:
    explicit = str(evidence.get("sourceFormat") or "").lower()
    if explicit in {"docx", "markdown", "text"}:
        return explicit
    source_type = str(evidence.get("sourceType") or "").lower()
    if source_type == "docx":
        return "docx"
    if source_type in {"md", "markdown"}:
        return "markdown"
    return "text"


def source_capabilities(source_format: str, evidence: Dict[str, Any]) -> Dict[str, bool]:
    return {
        "wordNativeStructure": source_format == "docx",
        "styles": source_format == "docx",
        "sections": source_format == "docx",
        "headersFooters": source_format == "docx",
        "embeddedAssets": source_format == "docx",
        "externalAssetPaths": source_format == "markdown",
        "markdownTables": source_format == "markdown",
        "plainTextOnly": source_format == "text",
    }


def build_semantic_blocks(
    evidence: Dict[str, Any],
    front_matter: Dict[str, Any],
    heading_tree: List[Dict[str, Any]],
    references: Dict[str, Any],
    acknowledgements: Dict[str, Any],
    appendix: Dict[str, Any],
    caption_blocks: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    heading_map = {item.get("sourceIndex"): item for item in heading_tree}
    caption_map = {item.get("sourceIndex"): item for item in caption_blocks}
    reference_heading = references.get("headingSourceIndex")
    reference_entries = {item.get("sourceIndex"): item for item in references.get("entries") or []}
    acknowledgement_heading = acknowledgements.get("headingSourceIndex")
    acknowledgement_blocks = {item.get("sourceIndex") for item in acknowledgements.get("blocks") or []}
    appendix_headings: Dict[Any, Dict[str, Any]] = {}
    appendix_blocks = set()
    for section in appendix.get("sections") or []:
        appendix_headings[section.get("headingSourceIndex")] = section
        appendix_blocks.update(item.get("sourceIndex") for item in section.get("blocks") or [])
    front_indexes = set(front_matter.get("sourceBlocks") or [])

    result: List[Dict[str, Any]] = []
    for ordinal, raw in enumerate(evidence.get("sourceBlocks") or []):
        source_index = raw.get("sourceIndex")
        kind = str(raw.get("kind") or "paragraph")
        text = normalize_text(raw.get("text"))
        role = str(raw.get("role") or "body")
        level = raw.get("level")
        confidence = 0.72

        if source_index in front_indexes:
            role, confidence = "front_matter", float(front_matter.get("confidence") or 0.7)
        elif source_index in heading_map:
            heading = heading_map[source_index]
            kind = "heading"
            text = normalize_text(heading.get("text") or text)
            level = heading.get("level") or level or 1
            role = f"heading_{level}"
            confidence = float(heading.get("confidence") or 0.8)
        elif kind == "heading":
            text = re.sub(r"^#{1,6}\s+", "", text).strip()
            level = int(level or 1)
            role = f"heading_{level}"
            confidence = 0.9
        elif source_index in caption_map:
            caption = caption_map[source_index]
            kind = "caption"
            text = normalize_text(caption.get("text") or text)
            role = str(caption.get("captionType") or "caption")
            confidence = float(caption.get("confidence") or 0.8)
        elif source_index == reference_heading:
            kind, role, confidence = "heading", "references_heading", 0.95
            level = 1
        elif source_index in reference_entries:
            kind, role, confidence = "reference", "reference_entry", 0.92
        elif source_index == acknowledgement_heading:
            kind, role, confidence = "heading", "acknowledgements_heading", 0.95
            level = 1
        elif source_index in acknowledgement_blocks:
            role, confidence = "acknowledgements_body", 0.85
        elif source_index in appendix_headings:
            kind, role, confidence = "heading", "appendix_heading", 0.95
            level = 1
        elif source_index in appendix_blocks:
            role, confidence = "appendix_body", 0.82
        if kind == "image":
            role, confidence = "figure_asset", 0.98
        elif kind == "table":
            role, confidence = "table_asset", 0.95

        result.append(
            {
                "id": f"block-{ordinal:05d}",
                "kind": kind,
                "role": role,
                "text": text or None,
                "level": level,
                "sourceAnchor": raw.get("sourceAnchor") or {"kind": "source_index", "index": source_index},
                "confidence": round(confidence, 2),
                "attributes": {
                    key: raw.get(key)
                    for key in ("path", "alt", "lines", "styleId", "styleName", "attributes")
                    if raw.get(key) is not None
                },
            }
        )
    return result


def build_thesis_ir(evidence: Dict[str, Any]) -> Dict[str, Any]:
    numbering_analysis = analyze_candidates(extract_items_from_evidence(evidence))
    if evidence.get("sourceType") == "docx":
        front_matter = extract_front_matter_from_docx(evidence)
    else:
        front_matter = extract_front_matter_from_text(evidence)
    heading_tree = build_filtered_heading_tree(evidence, numbering_analysis)
    acknowledgements = extract_acknowledgements(evidence)
    appendix = extract_appendix(evidence)
    caption_blocks = extract_caption_blocks(evidence)
    attachable_asset_candidates = build_attachable_asset_candidates(evidence, caption_blocks, front_matter, heading_tree)
    references = extract_references_section(evidence)
    body_blocks = build_body_blocks(
        evidence,
        heading_tree,
        front_matter,
        references,
        acknowledgements,
        appendix,
        caption_blocks,
    )
    reference_normalization = analyze_reference_normalization(references, body_blocks)
    ambiguities = list(numbering_analysis.get("conflicts") or [])
    if not front_matter.get("abstractCn"):
        ambiguities.append({"kind": "missing_abstract_candidate", "message": "No stable abstract block recovered."})
    if reference_normalization.get("bodyNumericCitationCount"):
        ambiguities.append(
            {
                "kind": "body_numeric_citation_mode_mismatch",
                "message": "Body text still contains numeric citations that are not safe to auto-convert.",
                "count": reference_normalization.get("bodyNumericCitationCount"),
            }
        )
    if caption_blocks and not (evidence.get("preserveableAssets") or {}).get("images") and evidence.get("sourceType") == "docx":
        ambiguities.append(
            {
                "kind": "caption_without_preserved_image_asset",
                "message": "Caption blocks were recovered but no preserveable image assets were detected in evidence.",
                "count": len(caption_blocks),
            }
        )
    low_confidence_asset_candidates = [item for item in attachable_asset_candidates if float(item.get("confidence") or 0.0) < 0.7]
    if low_confidence_asset_candidates:
        ambiguities.append(
            {
                "kind": "low_confidence_attachable_asset_candidates",
                "message": "Some preserved assets remain below the confidence threshold for automatic reattachment.",
                "count": len(low_confidence_asset_candidates),
            }
        )
    ambiguities.extend(build_asset_anchor_ambiguities(evidence, caption_blocks))
    strategy_candidate = choose_default_strategy(
        str(evidence.get("sourceType")),
        numbering_analysis,
        front_matter,
        evidence.get("preserveableAssets") or {},
    )
    confidence = {
        "frontMatter": front_matter.get("confidence"),
        "numbering": numbering_analysis.get("confidence"),
        "overall": round(
            (
                float(front_matter.get("confidence") or 0.0)
                + float(numbering_analysis.get("confidence") or 0.0)
            ) / 2,
            2,
        ),
    }
    candidate_action_counts = Counter(
        str(item.get("recommendedAction") or "unknown")
        for item in attachable_asset_candidates
    )
    source_format = canonical_source_format(evidence)
    semantic_blocks = build_semantic_blocks(
        evidence,
        front_matter,
        heading_tree,
        references,
        acknowledgements,
        appendix,
        caption_blocks,
    )
    return {
        "schema": "paper-formatter.thesis-ir",
        "version": "2.0",
        "sourceType": evidence.get("sourceType"),
        "source": {
            "path": evidence.get("source"),
            "format": source_format,
            "evidenceVersion": evidence.get("evidenceVersion") or "1.x",
            "capabilities": source_capabilities(source_format, evidence),
        },
        "strategyCandidate": strategy_candidate,
        "evidenceSummary": {
            "source": evidence.get("source"),
            "lineCount": evidence.get("lineCount"),
            "paragraphCount": evidence.get("paragraphCount"),
            "candidateBlockCount": len(evidence.get("candidateBlocks") or []),
        },
        "frontMatter": front_matter,
        "headingTree": heading_tree,
        "acknowledgements": acknowledgements,
        "appendix": appendix,
        "captionBlocks": caption_blocks,
        "attachableAssetCandidates": attachable_asset_candidates,
        "attachableAssetCandidateSummary": {
            "count": len(attachable_asset_candidates),
            "actionCounts": dict(candidate_action_counts),
        },
        "bodyBlocks": body_blocks,
        "semanticBlocks": semantic_blocks,
        "semanticBlockCount": len(semantic_blocks),
        "compatibility": {
            "legacyViewsPresent": True,
            "legacyFields": [
                "frontMatter", "headingTree", "bodyBlocks", "captionBlocks",
                "references", "acknowledgements", "appendix",
            ],
        },
        "references": {
            **references,
            "normalization": reference_normalization,
        },
        "numberingAnalysis": numbering_analysis,
        "preserveableAssets": evidence.get("preserveableAssets") or {},
        "assetAnchorAmbiguities": build_asset_anchor_ambiguities(evidence, caption_blocks),
        "ambiguities": ambiguities[:120],
        "confidence": confidence,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build ThesisIR from extracted source evidence.")
    parser.add_argument("--evidence-json", required=True, help="Path to source evidence JSON")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    evidence = json.loads(Path(args.evidence_json).read_text(encoding="utf-8"))
    payload = build_thesis_ir(evidence)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
