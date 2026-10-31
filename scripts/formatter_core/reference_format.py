from __future__ import annotations

import re
from typing import Any


TYPE_MARKER = re.compile(r"\[(?:J|M|C|D|P|S|R|N|DB|CP|EB|DS|Z)(?:/(?:OL|CD|DK|MT))?\]", re.I)
CHINESE = re.compile(r"[\u4e00-\u9fff]")


def check_reference_format(ir: dict[str, Any], rules: dict[str, Any] | None = None) -> dict[str, Any]:
    entries = (ir.get("references") or {}).get("entries") or []
    issues = []
    configured = ((rules or {}).get("references") or {}).get("citation_mode", "preserve")
    labels = [re.match(r"^\s*\[\s*(\d+)\s*\]", str(entry.get("text") or "")) for entry in entries]
    mode = "numeric" if any(labels) else "author_year"
    if configured not in {"preserve", "auto", mode} and entries:
        issues.append({"kind": "citation_mode_mismatch", "expected": configured, "actual": mode, "severity": "manual_check"})
    blocks = ir.get("semanticBlocks") or ir.get("bodyBlocks") or []
    candidates = [b for b in blocks if re.match(r"^\s*\[\s*\d+\s*\]", str(b.get("text") or "")) and TYPE_MARKER.search(str(b.get("text") or ""))]
    body = [b for b in blocks if b.get("role") != "reference_entry" and b.get("kind") != "reference"]
    has_citations = any(re.search(r"\[\s*\d+(?:\s*[-,，]\s*\d+)*\s*\]", str(b.get("text") or "")) for b in body)
    if not entries and (candidates or has_citations or (ir.get("references") or {}).get("heading")):
        issues.append({"kind": "reference_entries_not_recognized", "candidateCount": len(candidates), "severity": "error"})
    if mode == "numeric":
        numbers = [int(m.group(1)) for m in labels if m]
        if len(numbers) != len(entries) or numbers != list(range(1, len(entries) + 1)):
            issues.append({"kind": "reference_number_sequence", "numbers": numbers, "severity": "manual_check"})
        known = set(numbers)
        for block in body:
            for citation in re.finditer(r"\[\s*(\d+(?:\s*[-–,，]\s*\d+)*)\s*\]", str(block.get("text") or "")):
                cited = set()
                for part in re.split(r"[,，]", citation.group(1)):
                    bounds = re.split(r"[-–]", part)
                    first, last = int(bounds[0]), int(bounds[-1])
                    if first <= last and last - first <= 1000:
                        cited.update(range(first, last + 1))
                if cited - known:
                    issues.append({"kind": "citation_without_reference", "blockId": block.get("id"), "numbers": sorted(cited - known), "severity": "manual_check"})
    foreign_seen = False
    previous_surname = ""
    marker_blocks = [str(block.get("id") or "") for block in ir.get("semanticBlocks") or [] if re.search(r"\[@[A-Za-z0-9]", str(block.get("text") or ""))]
    if marker_blocks:
        issues.append({"kind": "latex_citation_markers_preserved_in_word", "blockIds": marker_blocks, "severity": "manual_check"})
    for index, entry in enumerate(entries):
        text = re.sub(r"^\s*\[\d+\]\s*", "", str(entry.get("text") or ""))
        authors = re.split(r"[.．]", text, maxsplit=1)[0].strip()
        first_author = re.split(r"[,，]", authors, maxsplit=1)[0].strip()
        language = "zh" if CHINESE.search(first_author) else "en"
        codes = []
        if not TYPE_MARKER.search(text):
            codes.append("missing_document_type_marker")
        if not re.search(r"\b(?:18|19|20)\d{2}\b", text):
            codes.append("missing_or_unrecognized_publication_year")
        if language == "zh":
            if foreign_seen and mode == "author_year":
                codes.append("chinese_entry_after_foreign_group")
            named = [name.strip() for name in re.split(r"[,，、]", authors) if name.strip() and name.strip() != "等"]
            if len(named) > 3:
                codes.append("more_than_three_named_authors")
        else:
            foreign_seen = True
            surname = first_author.split()[0] if first_author.split() else ""
            if mode == "author_year" and surname and surname != surname.upper():
                codes.append("foreign_surname_not_uppercase")
            if mode == "author_year" and previous_surname and surname.upper() < previous_surname:
                codes.append("foreign_surname_order")
            previous_surname = surname.upper()
            named = [name.strip() for name in re.split(r"[,，]", authors) if name.strip() and not re.match(r"et\s+al", name.strip(), re.I)]
            if len(named) > 3:
                codes.append("more_than_three_named_authors")
            if re.search(r"\b[A-Z]\.", text.split("[", 1)[0]):
                codes.append("possible_dotted_foreign_initials")
        for code in codes:
            issues.append({"kind": code, "entryIndex": index, "sourceIndex": entry.get("sourceIndex"), "severity": "manual_check"})
    return {
        "standard": ((rules or {}).get("references") or {}).get("standard", "GB/T 7714-2015"), "citationMode": mode, "requiredCitationMode": configured, "citationManagement": False,
        "entryCount": len(entries), "issues": issues, "automaticTextChanges": False,
        "font": {"eastAsia": "宋体", "sizePt": 10.5, "latin": "宋体"},
        "manualChecks": ([
            "Chinese entries first, sorted by verified pinyin; foreign entries follow in surname order. Check names against the required author-year rules.",
            "Body author-year citations and same-year a/b suffixes must match their entries; Word does not rewrite them.",
        ] if mode == "author_year" else [
            "Keep numeric labels linked to body citations; check duplicates, gaps, ranges and absent targets before changing citation style.",
        ]) + [
            "At most three named authors; additional authors use 等 or et al. Do not invent missing authors.",
            "Check type-specific fields, publisher, place, volume/issue, pages, patent/standard identifiers and electronic access dates against the source.",
        ],
        "limitations": ["Heuristic warnings only, not a standards compliance certificate; entries and body citations are not rewritten or reordered."],
    }
