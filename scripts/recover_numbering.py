#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


SCIENCE_DECIMAL_RE = re.compile(r"^(?P<prefix>\d+(?:\.\d+)*)\s+.+$")
CN_CHAPTER_RE = re.compile(r"^第[一二三四五六七八九十百千]+章")
CN_SECTION_RE = re.compile(r"^第[一二三四五六七八九十百千]+节")
CN_LIST_RE = re.compile(r"^[一二三四五六七八九十百千]+[、.．]\s*.+$")
CN_PAREN_LIST_RE = re.compile(r"^（[一二三四五六七八九十百千]+）")
ARABIC_LIST_RE = re.compile(r"^\d+[、.．)]\s*.+$")
CIRCLED_RE = re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")


def classify_prefix(text: str) -> Optional[Dict[str, Any]]:
    value = (text or "").strip()
    if not value:
        return None
    match = SCIENCE_DECIMAL_RE.match(value)
    if match:
        levels = [int(part) for part in match.group("prefix").split(".")]
        return {"family": "science_decimal", "level": len(levels), "prefix": match.group("prefix"), "numbers": levels}
    if CN_CHAPTER_RE.match(value):
        return {"family": "cn_chapter", "level": 1, "prefix": value.split()[0], "numbers": None}
    if CN_SECTION_RE.match(value):
        return {"family": "cn_chapter", "level": 2, "prefix": value.split()[0], "numbers": None}
    if CN_LIST_RE.match(value):
        return {"family": "cn_list", "level": 1, "prefix": value.split()[0], "numbers": None}
    if CN_PAREN_LIST_RE.match(value):
        return {"family": "cn_list", "level": 2, "prefix": value.split()[0], "numbers": None}
    if CIRCLED_RE.match(value):
        return {"family": "circled", "level": 3, "prefix": value[:1], "numbers": None}
    if ARABIC_LIST_RE.match(value):
        return {"family": "arabic_list", "level": 3, "prefix": value.split()[0], "numbers": None}
    return None


def analyze_candidates(items: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    candidates: List[Dict[str, Any]] = []
    family_counts: Dict[str, int] = {}
    conflicts: List[Dict[str, Any]] = []
    previous_level: Optional[int] = None
    previous_numbers: Optional[List[int]] = None
    previous_family: Optional[str] = None

    for item in items:
        text = str(item.get("text") or "").strip()
        if not text:
            continue
        classified = classify_prefix(text)
        if not classified:
            continue
        entry = {
            "sourceIndex": item.get("sourceIndex"),
            "text": text,
            "family": classified["family"],
            "level": classified["level"],
            "prefix": classified["prefix"],
        }
        candidates.append(entry)
        family = str(classified["family"])
        family_counts[family] = family_counts.get(family, 0) + 1
        current_level = int(classified["level"])
        current_numbers = classified.get("numbers")

        if previous_level is not None and current_level > previous_level + 1:
            conflicts.append(
                {
                    "kind": "level_jump",
                    "sourceIndex": item.get("sourceIndex"),
                    "text": text,
                    "previousLevel": previous_level,
                    "currentLevel": current_level,
                }
            )
        if previous_family and previous_family != family:
            conflicts.append(
                {
                    "kind": "mixed_family_transition",
                    "sourceIndex": item.get("sourceIndex"),
                    "text": text,
                    "previousFamily": previous_family,
                    "currentFamily": family,
                }
            )
        if current_numbers and previous_numbers and len(current_numbers) == len(previous_numbers):
            if current_numbers[:-1] == previous_numbers[:-1] and current_numbers[-1] > previous_numbers[-1] + 1:
                conflicts.append(
                    {
                        "kind": "science_decimal_gap",
                        "sourceIndex": item.get("sourceIndex"),
                        "text": text,
                        "previous": previous_numbers,
                        "current": current_numbers,
                    }
                )
        previous_level = current_level
        previous_numbers = current_numbers or previous_numbers
        previous_family = family

    sorted_families = sorted(family_counts.items(), key=lambda item: (-item[1], item[0]))
    families_present = [name for name, _ in sorted_families]
    dominant_family = families_present[0] if families_present else None
    if set(families_present) == {"cn_chapter", "science_decimal"}:
        dominant_family = "hybrid_chapter_decimal"
    confidence = 0.0
    if candidates:
        confidence = 0.55
        if len(families_present) == 1:
            confidence += 0.25
        if not conflicts:
            confidence += 0.15
        if len(candidates) >= 3:
            confidence += 0.05
    return {
        "familiesPresent": families_present,
        "dominantFamily": dominant_family,
        "candidates": candidates,
        "conflicts": conflicts,
        "confidence": round(min(confidence, 1.0), 2),
    }


def extract_items_from_evidence(evidence: Dict[str, Any]) -> List[Dict[str, Any]]:
    source_type = evidence.get("sourceType")
    items: List[Dict[str, Any]] = []
    if source_type == "docx":
        for paragraph in evidence.get("paragraphEvidence") or []:
            items.append({"sourceIndex": paragraph.get("index"), "text": paragraph.get("text")})
    else:
        for line in evidence.get("lineEvidence") or []:
            items.append({"sourceIndex": line.get("index"), "text": line.get("normalizedText") or line.get("text")})
    return items


def main() -> None:
    parser = argparse.ArgumentParser(description="Recover numbering families and heading-prefix conflicts from source evidence.")
    parser.add_argument("evidence_json", help="Path to source_evidence.json")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    evidence = json.loads(Path(args.evidence_json).read_text(encoding="utf-8"))
    payload = analyze_candidates(extract_items_from_evidence(evidence))
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
