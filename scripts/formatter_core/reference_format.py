from __future__ import annotations

import re
from typing import Any


TYPE_MARKER = re.compile(r"\[(?:J|M|C|D|P|S|R|N|DB|CP|EB|DS|Z)(?:/(?:OL|CD|DK|MT))?\]", re.I)
CHINESE = re.compile(r"[\u4e00-\u9fff]")


def check_reference_format(ir: dict[str, Any]) -> dict[str, Any]:
    entries = (ir.get("references") or {}).get("entries") or []
    issues = []
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
            if foreign_seen:
                codes.append("chinese_entry_after_foreign_group")
            named = [name.strip() for name in re.split(r"[,，、]", authors) if name.strip() and name.strip() != "等"]
            if len(named) > 3:
                codes.append("more_than_three_named_authors")
        else:
            foreign_seen = True
            surname = first_author.split()[0] if first_author.split() else ""
            if surname and surname != surname.upper():
                codes.append("foreign_surname_not_uppercase")
            if previous_surname and surname.upper() < previous_surname:
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
        "standard": "GB/T 7714-2015", "citationMode": "author_year", "citationManagement": False,
        "entryCount": len(entries), "issues": issues, "automaticTextChanges": False,
        "font": {"eastAsia": "宋体", "sizePt": 10.5, "latin": "宋体"},
        "manualChecks": [
            "Chinese entries first, sorted by verified pinyin of the first author's surname; polyphonic surnames require confirmation.",
            "Foreign entries follow, sorted by first-author surname; surnames uppercase, spaced initials without dots.",
            "At most three named authors; additional authors use 等 or et al. Do not invent missing authors.",
            "Check type-specific fields, publisher, place, volume/issue, pages, patent/standard identifiers and electronic access dates against the source.",
            "Body citations use parenthesized author-year, semicolons, and year suffix a/b for same-author same-year works; Word does not rewrite or renumber them.",
        ],
        "limitations": ["Heuristic warnings only, not a standards compliance certificate; entries and body citations are not rewritten or reordered."],
    }
