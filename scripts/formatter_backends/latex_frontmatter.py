from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Callable


TEXT_FIELDS = {
    "title", "titleEn", "author", "academy", "session", "studentNumber",
    "className", "teacher", "teacherTitle", "date", "abstractCn", "abstractEn",
    "keywordsCn", "keywordsEn", "frontMatterPdf",
}
BOOLEAN_FIELDS = {"cover", "statement"}
ABSTRACT_HEADINGS = {"摘要": "Cn", "中文摘要": "Cn", "英文摘要": "En", "abstract": "En", "english abstract": "En"}
KEYWORDS = re.compile(r"^(关键词|关键字|keywords?|key words)\s*[:：]\s*(.*)$", re.I)


def load_metadata(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    metadata = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(metadata, dict) or set(metadata) - TEXT_FIELDS - BOOLEAN_FIELDS:
        raise ValueError("LaTeX metadata must be an object with documented fields only")
    for key, value in metadata.items():
        if key in TEXT_FIELDS and not isinstance(value, str):
            raise ValueError(f"LaTeX metadata {key} must be a string")
        if key in BOOLEAN_FIELDS and not isinstance(value, bool):
            raise ValueError(f"LaTeX metadata {key} must be a boolean")
    return metadata


def prepare_frontmatter(ir: dict[str, Any], metadata: dict[str, Any], escape: Callable[[str], str], render: Callable[[str, str], str]) -> dict[str, Any]:
    values = {"title": str((ir.get("frontMatter") or {}).get("title") or "")}
    abstract_blocks: dict[str, list[dict[str, Any]]] = {"Cn": [], "En": []}
    consumed: set[str] = set()
    language = None
    title_blocks = []
    for block in ir.get("semanticBlocks") or []:
        text = str(block.get("text") or "").strip()
        identifier = str(block.get("id") or "")
        if text == values["title"] and block.get("role") == "front_matter":
            title_blocks.append(identifier)
            continue
        abstract_language = ABSTRACT_HEADINGS.get(re.sub(r"\s+", " ", text).lower())
        inline_abstract = re.match(r"^(摘要|中文摘要|英文摘要|abstract)\s*[:：]\s*(.+)$", text, re.I)
        if inline_abstract:
            language = "En" if inline_abstract.group(1).lower() in {"英文摘要", "abstract"} else "Cn"
            abstract_blocks[language].append(dict(block, text=inline_abstract.group(2)))
            consumed.add(identifier)
            continue
        if abstract_language:
            language = abstract_language
            consumed.add(identifier)
            continue
        # Markdown emphasis around the label is presentation, not part of the
        # label. Consume the source block instead of copying it into abstract.
        keyword_text = re.sub(r"^\*\*([^*]+[:：])\*\*\s*", r"\1 ", text)
        keyword = KEYWORDS.match(keyword_text)
        if keyword and language:
            values[f"keywords{language}"] = keyword.group(2)
            consumed.add(identifier)
            language = None
            continue
        if block.get("kind") == "heading":
            if language or block.get("role") != "front_matter":
                break
        if language and block.get("kind") == "paragraph":
            abstract_blocks[language].append(block)
            consumed.add(identifier)
    for suffix, blocks in abstract_blocks.items():
        if blocks:
            values[f"abstract{suffix}"] = "\n\n".join(str(block.get("text") or "") for block in blocks)
    values.update(metadata)
    errors = []
    warnings = []
    if (ir.get("frontMatter") or {}).get("abstractCn") and not values.get("abstractCn"):
        warnings.append("IR did not provide separable abstract blocks; combined source paragraphs remain unchanged. Separate abstract/body paragraphs before using the CLS abstract macro.")
    if values.get("cover"):
        missing = [key for key in ("title", "author", "academy", "session", "studentNumber", "className", "teacher", "teacherTitle", "date") if not str(values.get(key) or "").strip()]
        if missing:
            errors.append({"kind": "missing_cover_metadata", "fields": missing})
    if values.get("statement") and any(not values.get(key) for key in ("title", "author", "date")):
        errors.append({"kind": "missing_statement_metadata", "message": "Explicit title, author and date are required for an unsigned statement."})
    commands = [rf"\title{{{{{escape(str(values.get('title') or ''))}}}}}", rf"\author{{{escape(str(values.get('author') or ''))}}}", rf"\date{{{escape(str(values.get('date') or ''))}}}"]
    for field, macro in (("academy", "academy"), ("session", "session"), ("studentNumber", "studentNumber"), ("className", "myClass"), ("titleEn", "entitle")):
        commands.append(rf"\{macro}{{{escape(str(values.get(field) or ''))}}}")
    commands.append(rf"\teacher{{{escape(str(values.get('teacher') or ''))}}}{{{escape(str(values.get('teacherTitle') or ''))}}}")
    commands.append(r"\setsignature{}")
    front = []
    if values.get("frontMatterPdf"):
        if values.get("cover") or values.get("statement"):
            errors.append({"kind": "duplicate_frontmatter", "message": "External front PDF cannot be combined with cover/statement macros."})
        front.append(r"\includepdf[pages=-,pagecommand={\thispagestyle{empty}},fitpaper=true]{front-matter.pdf}")
    if values.get("cover") and not errors:
        front.append(r"\customCover")
    if values.get("statement") and not errors:
        front.append(r"\makestatement")
    for suffix, macro in (("Cn", "ZhAbstract"), ("En", "EnAbstract")):
        abstract = str(values.get(f"abstract{suffix}") or "")
        if abstract:
            if f"abstract{suffix}" in metadata:
                rendered = "\n\n".join(render(part, f"metadata-abstract{suffix}") for part in abstract.split("\n\n"))
            else:
                rendered = "\n\n".join(render(str(block.get("text") or ""), str(block.get("id") or "")) for block in abstract_blocks[suffix])
            rendered_keywords = render(str(values.get(f"keywords{suffix}") or ""), f"frontmatter-keywords-{suffix}")
            if suffix == "En":
                # The pinned upstream macro sets 5pt before selecting fonts.
                # Keep the class untouched; explicitly select the adapter size.
                rendered = r"{\thispagestyle{formatterabstract}\fontsize{10.5}{15}\selectfont " + rendered + "}"
                rendered_keywords = r"{\fontsize{10.5}{15}\selectfont " + rendered_keywords + "}"
            front.append(rf"\{macro}{{{rendered}}}{{{rendered_keywords}}}")
            if not values.get(f"keywords{suffix}"):
                warnings.append(f"Missing {suffix} abstract keywords; no keywords were invented.")
            if suffix == "En" and not values.get("titleEn"):
                warnings.append("English title was not supplied; the class's English title field remains blank.")
    if front:
        consumed.update(title_blocks)
        front.append(r"\customContent")
    else:
        consumed.clear()
    return {"metadataTex": "\n".join(commands), "frontmatterTex": "\n\n".join(front), "consumedIds": consumed, "errors": errors, "warnings": warnings, "cover": bool(values.get("cover")), "statement": bool(values.get("statement")), "abstractCn": bool(values.get("abstractCn")), "abstractEn": bool(values.get("abstractEn")), "tocInserted": bool(front), "metadataFields": sorted(metadata)}
