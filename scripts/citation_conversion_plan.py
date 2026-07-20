#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx_ooxml import load_rules


NUMERIC_CITATION_RE = re.compile(r"\[\s*(?P<labels>\d+(?:\s*[-,，]\s*\d+)*)\s*\]")
NUMERIC_REFERENCE_LABEL_RE = re.compile(r"^\[\s*(?P<label>\d+)\s*\]\s*")


def parse_numeric_labels(raw: str) -> List[str]:
    value = raw.replace("，", ",")
    parts = [part.strip() for part in re.split(r"\s*,\s*", value) if part.strip()]
    labels: List[str] = []
    for part in parts:
        if "-" in part:
            start, end = [item.strip() for item in part.split("-", 1)]
            if start.isdigit() and end.isdigit():
                start_num = int(start)
                end_num = int(end)
                if start_num <= end_num and end_num - start_num <= 20:
                    labels.extend([str(number) for number in range(start_num, end_num + 1)])
                    continue
        labels.append(part)
    return labels


def build_reference_label_map(thesis_ir: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    references = thesis_ir.get("references") or {}
    label_map: Dict[str, Dict[str, Any]] = {}
    for entry in references.get("entries") or []:
        text = str(entry.get("text") or "").strip()
        match = NUMERIC_REFERENCE_LABEL_RE.match(text)
        if not match:
            continue
        label = match.group("label")
        normalized_text = NUMERIC_REFERENCE_LABEL_RE.sub("", text, count=1).strip()
        label_map[label] = {
            "sourceIndex": entry.get("sourceIndex"),
            "originalText": text,
            "normalizedText": normalized_text,
        }
    return label_map


def find_duplicate_reference_labels(thesis_ir: Dict[str, Any]) -> List[str]:
    references = thesis_ir.get("references") or {}
    seen: Dict[str, int] = {}
    duplicates: List[str] = []
    for entry in references.get("entries") or []:
        text = str(entry.get("text") or "").strip()
        match = NUMERIC_REFERENCE_LABEL_RE.match(text)
        if not match:
            continue
        label = match.group("label")
        seen[label] = seen.get(label, 0) + 1
        if seen[label] == 2:
            duplicates.append(label)
    return duplicates


def classify_candidate_kind(labels: List[str], missing_labels: List[str], raw_labels: str) -> str:
    if missing_labels:
        return "unmatched_label"
    if "-" in raw_labels:
        return "range_label"
    if len(labels) > 1:
        return "multi_label"
    return "single_label"


def build_plan(thesis_ir: Dict[str, Any], rules: Dict[str, Any]) -> Dict[str, Any]:
    references = thesis_ir.get("references") or {}
    normalization = references.get("normalization") or {}
    reference_label_map = build_reference_label_map(thesis_ir)
    duplicate_reference_labels = find_duplicate_reference_labels(thesis_ir)
    candidates: List[Dict[str, Any]] = []
    unmatched_labels: List[str] = []
    candidates_by_kind = {
        "single_label": 0,
        "multi_label": 0,
        "range_label": 0,
        "unmatched_label": 0,
        "duplicate_label": len(duplicate_reference_labels),
    }
    manual_review_reasons: List[str] = []
    fully_matched_count = 0
    partially_matched_count = 0
    unmatched_span_count = 0
    range_expansion_warnings: List[Dict[str, Any]] = []

    for block in thesis_ir.get("bodyBlocks") or []:
        text = str(block.get("text") or "")
        if not text:
            continue
        for match in NUMERIC_CITATION_RE.finditer(text):
            raw_labels = match.group("labels")
            labels = parse_numeric_labels(raw_labels)
            matched_references: List[Dict[str, Any]] = []
            missing_for_span: List[str] = []
            for label in labels:
                reference = reference_label_map.get(label)
                if reference is None:
                    missing_for_span.append(label)
                    if label not in unmatched_labels:
                        unmatched_labels.append(label)
                    continue
                matched_references.append(
                    {
                        "label": label,
                        "sourceIndex": reference.get("sourceIndex"),
                        "normalizedTextPreview": str(reference.get("normalizedText") or "")[:200],
                    }
                )
            kind = classify_candidate_kind(labels, missing_for_span, raw_labels)
            reasons = ["body_numeric_citation_requires_manual_review"]
            if len(labels) > 1:
                reasons.append("multi_reference_citation_span")
            if "-" in raw_labels:
                reasons.append("numeric_range_citation_span")
                if len(labels) <= 1:
                    range_expansion_warnings.append(
                        {
                            "sourceIndex": block.get("sourceIndex"),
                            "citationSpan": match.group(0),
                            "rawLabels": raw_labels,
                            "reason": "numeric_range_span_did_not_expand_as_expected",
                        }
                    )
            if missing_for_span:
                reasons.append("unmatched_reference_labels")
            if any(label in duplicate_reference_labels for label in labels):
                reasons.append("duplicate_reference_label_detected")
            for reason in reasons:
                if reason not in manual_review_reasons:
                    manual_review_reasons.append(reason)
            if not missing_for_span:
                fully_matched_count += 1
            elif matched_references:
                partially_matched_count += 1
            else:
                unmatched_span_count += 1
            candidates_by_kind[kind] = candidates_by_kind.get(kind, 0) + 1
            candidates.append(
                {
                    "sourceIndex": block.get("sourceIndex"),
                    "citationSpan": match.group(0),
                    "labels": labels,
                    "matchedReferences": matched_references,
                    "missingLabels": missing_for_span,
                    "kind": kind,
                    "manualReviewRequired": True,
                    "manualReviewReasons": reasons,
                    "reason": "Numeric body citations are detected, but body-text conversion is intentionally not auto-applied in this milestone.",
                    "textPreview": text[:220],
                }
            )

    configured_mode = ((rules.get("references") or {}).get("citation_mode")) or None
    bibliography_numbering = ((rules.get("references") or {}).get("bibliography_numbering")) or None
    strict_school_compliance = bool(((rules.get("defaults") or {}).get("strict_school_compliance")) or False)
    body_numeric_count = int(normalization.get("bodyNumericCitationCount") or 0)
    safe_list_strip = bool(normalization.get("safeListLabelStripCandidate"))

    return {
        "sourceType": thesis_ir.get("sourceType"),
        "configuredCitationMode": configured_mode,
        "configuredBibliographyNumbering": bibliography_numbering,
        "strictSchoolCompliance": strict_school_compliance,
        "safeListLabelStripCandidate": safe_list_strip,
        "autoConvertibleBibliographyList": bool(
            bibliography_numbering == "none" and (rules.get("references") or {}).get("strip_numeric_labels_when_unnumbered", True)
        ),
        "autoConvertibleBodyCitations": False,
        "bodyNumericCitationCount": body_numeric_count,
        "referenceEntryCount": normalization.get("entryCount"),
        "referenceNumericLabelCount": normalization.get("numericLabelCount"),
        "referenceLabelMapCount": len(reference_label_map),
        "candidateCount": len(candidates),
        "candidatesByKind": candidates_by_kind,
        "matchSummary": {
            "fullyMatchedCount": fully_matched_count,
            "partiallyMatchedCount": partially_matched_count,
            "unmatchedSpanCount": unmatched_span_count,
        },
        "manualReviewReasons": manual_review_reasons,
        "unmatchedLabels": unmatched_labels,
        "duplicateReferenceLabels": duplicate_reference_labels,
        "rangeExpansionWarnings": range_expansion_warnings[:40],
        "candidateSpans": candidates[:120],
        "recommendedAction": (
            "manual_review_required"
            if body_numeric_count
            else "no_body_citation_conversion_needed"
        ),
        "notes": [
            "This plan separates safe bibliography list denumbering from unsafe body citation rewriting.",
            "Body citation conversion remains plan-only in this milestone and should not silently rewrite thesis wording.",
            "strict-school compliance raises visibility of citation-mode mismatch but does not silently rewrite body citations.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a body citation conversion plan without rewriting thesis text.")
    parser.add_argument("--thesis-ir-json", required=True, help="Path to thesis_ir.json")
    parser.add_argument("--rules-yaml", required=True, help="Rules YAML path")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    thesis_ir = json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8"))
    rules = load_rules(args.rules_yaml)
    payload = build_plan(thesis_ir, rules)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
