"""Shared section recognition and the thesis components that formatting must check."""
from __future__ import annotations

import re


REFERENCES_HEADING_RE = re.compile(
    r"^\s*(?:(?:第[一二三四五六七八九十\d]+章|\d+(?:\.\d+)*[.、]?|[一二三四五六七八九十]+、)\s*)?"
    r"(?:参\s*考\s*文\s*献|references)(?:\s*[（(][^）)]*[）)])?\s*[:：]?\s*$",
    re.I,
)


def is_references_heading(text: str) -> bool:
    return bool(REFERENCES_HEADING_RE.fullmatch(text or ""))


def front_label(text: str):
    """Recognize labels with a colon or separating whitespace, keeping the payload."""
    match = re.match(r"^\s*(摘\s*要|中文摘要|Abstract|关键词|Key\s*words?)\s*(?:[:：]\s*|\s+)(.*)$", text or "", re.I)
    if not match:
        return None
    label = re.sub(r"\s+", "", match.group(1)).lower()
    role = {"摘要": "abstract_paragraph_cn", "中文摘要": "abstract_paragraph_cn", "abstract": "abstract_paragraph_en", "关键词": "keywords_cn", "keyword": "keywords_en", "keywords": "keywords_en"}[label]
    return role, match.group(2).strip()


def recover_front_roles(paragraphs: list[dict]) -> None:
    """Recover the abstract block from its context, without changing any text."""
    # Labels with omitted punctuation are common in incoming papers. Restrict
    # recovery to the prefix, before the first explicitly numbered body chapter.
    prefix = []
    for paragraph in paragraphs[:120]:
        text = str(paragraph.get("text") or "").strip()
        if re.match(r"^(?:1\s+|第一章|一、)", text) and not re.search(r"\.{3}|…{2}", text):
            break
        prefix.append(paragraph)
    labels = []
    for paragraph in prefix:
        match = front_label(str(paragraph.get("text") or ""))
        if match:
            paragraph["role"] = {"role": match[0], "confidence": 0.98, "signals": ["front_label_in_prefix"]}
            labels.append(paragraph)
    for language, label_role in (("cn", "abstract_paragraph_cn"), ("en", "abstract_paragraph_en")):
        abstract = next((p for p in labels if p["role"]["role"] == label_role), None)
        if abstract is None:
            continue
        before = [p for p in prefix if p["index"] < abstract["index"] and str(p.get("text") or "").strip()]
        for candidate in reversed(before[-4:]):
            text = str(candidate.get("text") or "").strip()
            if front_label(text) or re.search(r"目录|诚信|毕业论文[（(]|\d{6,}|[:：]", text):
                continue
            has_chinese = bool(re.search(r"[\u4e00-\u9fff]", text))
            if has_chinese != (language == "cn") or not 4 <= len(text) <= 160:
                continue
            if candidate.get("manualNumbering"):
                continue
            centered = (candidate.get("effectiveParagraph") or {}).get("jc") == "center"
            # Chinese title before the author and abstract; English title must
            # have a title-like centered presentation, not be the author line.
            if not centered and candidate is not prefix[0]:
                continue
            candidate["role"] = {"role": f"title_{language}", "confidence": 0.97, "signals": ["title_before_abstract"]}
            break


def inspect_requirements(audit: dict, rules: dict, ir: dict | None = None) -> dict:
    """Describe actual missing components, not blank fields on immutable covers."""
    paragraphs = audit.get("paragraphs") or []
    roles = {}
    for p in paragraphs:
        roles.setdefault((p.get("role") or {}).get("role"), []).append(p.get("index"))
    texts = [re.sub(r"\s+", "", str(p.get("text") or "")) for p in paragraphs]
    cover = any(re.fullmatch(r"(?:本科生?|本科)?毕业论文[（(]设计[）)]", t) for t in texts)
    statement = any("诚信承诺书" in t for t in texts)
    front = rules.get("front_matter") or {}
    template_pages = bool(front.get("preserve_cover_and_integrity_pages_from_template"))
    items = []

    def add(key, label, present, required=True):
        items.append({"key": key, "label": label, "present": bool(present), "required": bool(required)})

    add("cover", "模板封面", cover, template_pages)
    add("integrity_statement", "模板诚信承诺书", statement, template_pages)
    add("toc", "目录", roles.get("toc_heading") or roles.get("toc_field"), front.get("toc_before_abstract", False))
    for role, label in (("title_cn", "中文题目"), ("abstract_paragraph_cn", "中文摘要"), ("keywords_cn", "中文关键词"), ("title_en", "英文题目"), ("abstract_paragraph_en", "英文摘要"), ("keywords_en", "英文关键词")):
        add(role, label, roles.get(role))
    sections = audit.get("sections") or []
    inherited_headers, inherited_footers = {}, {}
    header_checks, footer_checks = [], []
    start = 0
    for section in sections:
        end = section.get("paragraphIndex")
        end = max((p.get("index", -1) for p in paragraphs), default=-1) if end is None else end
        content = [p for p in paragraphs if start <= p.get("index", -1) <= end]
        start = end + 1
        for ref in section.get("headerReferences") or []:
            inherited_headers[ref.get("type", "default")] = ref
        for ref in section.get("footerReferences") or []:
            inherited_footers[ref.get("type", "default")] = ref
        section_roles = {(p.get("role") or {}).get("role") for p in content}
        has_cover = any(re.fullmatch(r"(?:本科生?|本科)?毕业论文[（(]设计[）)]", re.sub(r"\s+", "", p.get("text") or "")) or "诚信承诺书" in (p.get("text") or "") for p in content)
        body_section = bool(section_roles & {"abstract_paragraph_cn", "abstract_paragraph_en", "reference_entry"}) or not has_cover and bool(section_roles & {"heading1", "heading2"})
        if body_section:
            variants = ["default", "first"] if section.get("titlePg") else ["default"]
            header_checks.append(not has_cover and all(str((inherited_headers.get(v) or {}).get("text") or "").strip() for v in variants))
            footer_checks.append(not has_cover and all(any(re.search(r"\bPAGE\b", field, re.I) for field in (inherited_footers.get(v) or {}).get("fieldInstructions", [])) for v in variants))
    add("body_header", "正文页眉", bool(header_checks) and all(header_checks), (rules.get("page") or {}).get("header_text_mode") != "preserve")
    add("page_numbers", "页码", bool(footer_checks) and all(footer_checks))
    reference_present = any(is_references_heading(p.get("text") or "") for p in paragraphs) or bool(((ir or {}).get("references") or {}).get("entries"))
    add("references", "参考文献", reference_present, bool(rules.get("references")))
    if ((ir or {}).get("references") or {}).get("detection") == "numbered_entries_fallback":
        add("references_heading", "参考文献标题", False)
    return {"items": items, "missing": [i for i in items if i["required"] and not i["present"]], "coverPolicy": "insert_verbatim_do_not_edit_or_fill"}
