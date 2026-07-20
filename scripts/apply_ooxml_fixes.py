#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from lxml import etree as LET

from docx_ooxml import (
    NS,
    W,
    build_repair_plan,
    detect_template_prefix_count,
    diff_audits,
    front_matter_policy_value,
    load_rules,
    normalize_rel_target_word,
    parse_manual_numbering,
    parse_document,
    pt_to_half_points,
    qn,
    safe_int,
    write_json,
)
from mathml_omml import (
    INLINE_MATH_RE,
    extract_inline_math_latex,
    extract_standalone_formula_latex,
    latex_to_omml_xml,
    paragraph_looks_like_formula,
    split_display_math_and_number,
)

for prefix, uri in NS.items():
    if prefix != "rel":
        ET.register_namespace(prefix, uri)


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_TYPE_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
REL_TYPE_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
REL_TYPE_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
SOURCE_CORE_REL_TYPES = {
    REL_TYPE_NUMBERING,
    REL_TYPE_FOOTNOTES,
    REL_TYPE_ENDNOTES,
    REL_TYPE_COMMENTS,
}
CAPTION_LABEL_TOKEN_RE = r"(?:\d+|[一二三四五六七八九十百千]+)(?:[a-zA-Z])?(?:[-—.．](?:\d+|[一二三四五六七八九十百千]+))?"
CAPTION_REF_RE = re.compile(rf"^(?P<kind>图|表)\s*(?P<label>{CAPTION_LABEL_TOKEN_RE})")
CROSS_CAPTION_RE = re.compile(rf"^(?P<kind>图|表)\s*(?P<label>{CAPTION_LABEL_TOKEN_RE})(?:\s+|[：:．。]|$)")
CROSS_REF_RE = re.compile(rf"(?P<kind>图|表)\s*(?P<label>{CAPTION_LABEL_TOKEN_RE})")
WORD_TOC_FIELD_INSTR = 'TOC \\o "1-3" \\h \\z \\u'
BODY_HEADING_RE = re.compile(
    r"^\s*(?:第?[一二三四五六七八九十百千0-9]+(?:[章节篇部])?|[0-9]+(?:[.．][0-9]+)*)(?:\s+|[、.．])+\S+"
)
ABSTRACT_CN_HEADINGS = {"中文摘要", "摘 要", "摘  要", "摘要"}
BODY_HEADING_STYLE_IDS = {"zafu_body", "Body Text", "Body Text Indent"}
REFERENCES_HEADING_RE = re.compile(r"^\s*(参考文献|References)(?:[（(].*?[)）])?\s*$")
TABLE_CAPTION_RE = re.compile(rf"^表\s*{CAPTION_LABEL_TOKEN_RE}(?:\s+|[：:．。]|$)")
DEFAULT_REFERENCE_TEMPLATE = Path(__file__).resolve().parent.parent / "浙江农林大学毕业论文模板参考.docx"


def clone_xml(node: ET.Element) -> ET.Element:
    return ET.fromstring(ET.tostring(node, encoding="utf-8"))


def read_docx_files(path: str) -> Dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def paragraph_nodes(body: ET.Element) -> List[ET.Element]:
    return [child for child in list(body) if child.tag == qn("p")]


def paragraph_text(node: ET.Element) -> str:
    return "".join(text.text or "" for text in node.findall(".//w:t", NS))


def infer_header_text_from_paragraphs(paragraphs: List[ET.Element], start_index: int) -> str:
    stop = min(len(paragraphs), max(start_index, 0) + 8)
    skip_values = {"目  录", "目录", "摘要", "ABSTRACT", "参考文献"}
    for idx in range(max(start_index, 0), stop):
        candidate = paragraph_text(paragraphs[idx]).strip()
        if candidate and candidate not in skip_values:
            return candidate
    return ""


def body_children_until_paragraph_count(body: ET.Element, paragraph_count: int) -> int:
    count = 0
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag == qn("p"):
            if count >= paragraph_count:
                return idx
            count += 1
    return len(children)


def find_child_index(body: ET.Element, target: ET.Element) -> Optional[int]:
    for idx, child in enumerate(list(body)):
        if child is target:
            return idx
    return None


def latex_to_omml(latex: str, display: str = "inline") -> Optional[ET.Element]:
    omml_xml = latex_to_omml_xml(latex, display=display)
    if omml_xml is None:
        return None
    return ET.fromstring(omml_xml)


def template_prefix_children(template_body: ET.Element, paragraph_count: int) -> List[ET.Element]:
    out = []
    seen_paragraphs = 0
    for child in list(template_body):
        if child.tag == qn("p"):
            if seen_paragraphs >= paragraph_count:
                break
            seen_paragraphs += 1
            out.append(child)
        elif child.tag == qn("sectPr"):
            break
        else:
            out.append(child)
    return out


def has_paragraph_level_sectpr(nodes: Iterable[ET.Element]) -> bool:
    for node in nodes:
        if node.tag != qn("p"):
            continue
        ppr = node.find("w:pPr", NS)
        if ppr is not None and ppr.find("w:sectPr", NS) is not None:
            return True
    return False


def extract_tail_children(source_body: ET.Element, source_prefix_paragraph_count: int) -> Tuple[List[ET.Element], Optional[ET.Element]]:
    source_children = list(source_body)
    start_index = body_children_until_paragraph_count(source_body, source_prefix_paragraph_count)
    tail_children = [clone_xml(child) for child in source_children[start_index:] if child.tag != qn("sectPr")]
    tail_sectpr = source_body.find("w:sectPr", NS)
    return tail_children, clone_xml(tail_sectpr) if tail_sectpr is not None else None


def compose_template_based_document(
    template_document_root: ET.Element,
    source_document_root: ET.Element,
    source_prefix_paragraph_count: int,
    template_prefix_paragraph_count: int,
) -> Tuple[Dict[str, Any], List[ET.Element]]:
    template_body = template_document_root.find("w:body", NS)
    source_body = source_document_root.find("w:body", NS)
    if template_body is None or source_body is None:
        return (
            {
                "changed": False,
                "mode": "template_base",
                "sourcePrefixParagraphCount": source_prefix_paragraph_count,
                "templatePrefixParagraphCount": template_prefix_paragraph_count,
                "indexShift": 0,
            },
            [],
        )

    prefix_nodes = [clone_xml(node) for node in template_prefix_children(template_body, template_prefix_paragraph_count)]
    template_tail_sectpr = template_body.find("w:sectPr", NS)
    source_tail_nodes, source_tail_sectpr = extract_tail_children(source_body, source_prefix_paragraph_count)

    for child in list(template_body):
        template_body.remove(child)

    for node in prefix_nodes:
        template_body.append(node)

    inserted_boundary_paragraph = 0
    if template_tail_sectpr is not None and not has_paragraph_level_sectpr(prefix_nodes):
        boundary_para = ET.Element(qn("p"))
        boundary_ppr = ET.SubElement(boundary_para, qn("pPr"))
        # The boundary paragraph closes the template front-matter section.
        # It must keep the template sectPr so cover/integrity-page pagination,
        # docGrid, and line geometry stay identical to the template.
        boundary_sectpr = clone_xml(template_tail_sectpr)
        boundary_ppr.append(boundary_sectpr)
        template_body.append(boundary_para)
        inserted_boundary_paragraph = 1

    for node in source_tail_nodes:
        template_body.append(node)

    final_tail = source_tail_sectpr if source_tail_sectpr is not None else (
        clone_xml(template_tail_sectpr) if template_tail_sectpr is not None else None
    )
    if final_tail is not None:
        template_body.append(final_tail)

    final_prefix_paragraph_count = template_prefix_paragraph_count + inserted_boundary_paragraph
    index_shift = final_prefix_paragraph_count - source_prefix_paragraph_count
    return (
        {
            "changed": True,
            "mode": "template_base",
            "sourcePrefixParagraphCount": source_prefix_paragraph_count,
            "templatePrefixParagraphCount": template_prefix_paragraph_count,
            "sourceTailChildCount": len(source_tail_nodes),
            "sourceTailParagraphCount": sum(1 for node in source_tail_nodes if node.tag == qn("p")),
            "sourceTailTableCount": sum(1 for node in source_tail_nodes if node.tag == qn("tbl")),
            "insertedBoundaryParagraph": inserted_boundary_paragraph,
            "finalPrefixParagraphCount": final_prefix_paragraph_count,
            "indexShift": index_shift,
        },
        source_tail_nodes + ([final_tail] if final_tail is not None else []),
    )


def paragraph_has_toc_field(paragraph: ET.Element) -> bool:
    for field in paragraph.findall(".//w:fldSimple", NS):
        instr = field.attrib.get(qn("instr")) or ""
        if "TOC" in instr.upper():
            return True
    return False


def build_toc_title_paragraph(page_break_before: bool = False) -> ET.Element:
    paragraph = ET.Element(qn("p"))
    ppr = ET.SubElement(paragraph, qn("pPr"))
    pstyle = ET.SubElement(ppr, qn("pStyle"))
    pstyle.set(qn("val"), "zafu_toc_title")
    if page_break_before:
        ET.SubElement(ppr, qn("pageBreakBefore"))
    run = ET.SubElement(paragraph, qn("r"))
    text = ET.SubElement(run, qn("t"))
    text.text = "目  录"
    return paragraph


def build_toc_field_paragraph() -> ET.Element:
    paragraph = ET.Element(qn("p"))
    ppr = ET.SubElement(paragraph, qn("pPr"))
    pstyle = ET.SubElement(ppr, qn("pStyle"))
    pstyle.set(qn("val"), "zafu_toc_body")
    fld_simple = ET.SubElement(paragraph, qn("fldSimple"))
    fld_simple.set(qn("instr"), WORD_TOC_FIELD_INSTR)
    run = ET.SubElement(fld_simple, qn("r"))
    text = ET.SubElement(run, qn("t"))
    text.text = "在 Word 中打开后更新目录"
    return paragraph


def first_run_rpr(paragraph: ET.Element) -> Optional[ET.Element]:
    first_run = paragraph.find("w:r", NS)
    if first_run is None:
        return None
    rpr = first_run.find("w:rPr", NS)
    return clone_xml(rpr) if rpr is not None else None


def append_text_run(paragraph: ET.Element, text: str, source_rpr: Optional[ET.Element] = None) -> None:
    if not text:
        return
    run = ET.SubElement(paragraph, qn("r"))
    if source_rpr is not None:
        run.append(clone_xml(source_rpr))
    node = ET.SubElement(run, qn("t"))
    node.text = text
    ensure_text_node_preserve(node, text)


def replace_paragraph_plain_text(paragraph: ET.Element, text: str) -> None:
    source_rpr = first_run_rpr(paragraph)
    ppr = paragraph.find("w:pPr", NS)
    ppr_clone = clone_xml(ppr) if ppr is not None else None
    for child in list(paragraph):
        paragraph.remove(child)
    if ppr_clone is not None:
        paragraph.append(ppr_clone)
    append_text_run(paragraph, text, source_rpr)


def append_tab_run(paragraph: ET.Element, source_rpr: Optional[ET.Element] = None) -> None:
    run = ET.SubElement(paragraph, qn("r"))
    if source_rpr is not None:
        run.append(clone_xml(source_rpr))
    ET.SubElement(run, qn("tab"))


def configure_equation_number_tabs(paragraph: ET.Element) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    tabs = ensure_child(ppr, "tabs")
    for tab in list(tabs.findall("w:tab", NS)):
        tabs.remove(tab)
    center_tab = ET.SubElement(tabs, qn("tab"))
    center_tab.set(qn("val"), "center")
    center_tab.set(qn("pos"), "4429")
    right_tab = ET.SubElement(tabs, qn("tab"))
    right_tab.set(qn("val"), "right")
    right_tab.set(qn("pos"), "8858")


def rebuild_paragraph_with_inline_math(paragraph: ET.Element, text: str) -> bool:
    matches = list(INLINE_MATH_RE.finditer(text))
    if not matches:
        return False
    source_rpr = first_run_rpr(paragraph)
    ppr = paragraph.find("w:pPr", NS)
    ppr_clone = clone_xml(ppr) if ppr is not None else None
    for child in list(paragraph):
        paragraph.remove(child)
    if ppr_clone is not None:
        paragraph.append(ppr_clone)
    cursor = 0
    inserted_math = False
    for match in matches:
        before = text[cursor : match.start()]
        append_text_run(paragraph, before, source_rpr)
        token = match.group(0)
        latex = extract_inline_math_latex(token)
        omml = latex_to_omml(latex.strip(), display="inline")
        if omml is None:
            append_text_run(paragraph, token, source_rpr)
        else:
            paragraph.append(omml)
            inserted_math = True
        cursor = match.end()
    append_text_run(paragraph, text[cursor:], source_rpr)
    return inserted_math


def rebuild_paragraph_with_display_math(paragraph: ET.Element, latex: str, equation_number: Optional[str]) -> bool:
    omml = latex_to_omml(latex.strip(), display="inline")
    if omml is None:
        return False
    source_rpr = first_run_rpr(paragraph)
    ppr = paragraph.find("w:pPr", NS)
    ppr_clone = clone_xml(ppr) if ppr is not None else None
    for child in list(paragraph):
        paragraph.remove(child)
    if ppr_clone is not None:
        paragraph.append(ppr_clone)
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    spacing = ensure_child(ppr, "spacing")
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    # Display equations must not inherit the thesis body's exact fixed line
    # spacing, or tall formulas get clipped in Word.
    spacing.set(qn("line"), "240")
    spacing.set(qn("lineRule"), "auto")
    if equation_number:
        configure_equation_number_tabs(paragraph)
        append_tab_run(paragraph, source_rpr)
        paragraph.append(omml)
        append_tab_run(paragraph, source_rpr)
        append_text_run(paragraph, equation_number, source_rpr)
    else:
        jc = ensure_child(ppr, "jc")
        jc.set(qn("val"), "center")
        paragraph.append(omml)
    return True


def _paragraph_is_standalone_display_delimiter(paragraph: ET.Element) -> bool:
    text = (element_text(paragraph) or "").strip()
    return text in {"$$", r"\[", r"\]"}


def _consume_standalone_dollar_block(paragraphs: List[ET.Element], start_index: int) -> Optional[Tuple[int, str]]:
    start = paragraphs[start_index]
    if (element_text(start) or "").strip() != "$$":
        return None
    latex_lines: List[str] = []
    cursor = start_index + 1
    while cursor < len(paragraphs):
        text = (element_text(paragraphs[cursor]) or "").strip()
        if text == "$$":
            latex = "\n".join(line for line in latex_lines if line).strip()
            if latex:
                return cursor, latex
            return None
        latex_lines.append(text)
        cursor += 1
    return None


REFERENCE_STOP_HEADINGS = {"致谢", "附录", "攻读学位期间发表的学术论文", "作者简介"}


def _looks_like_reference_entry(text: str) -> bool:
    stripped = (text or "").strip()
    if not stripped:
        return False
    if re.match(r"^\[\s*\d+\s*\]", stripped):
        return True
    if re.match(r"^[A-Z][A-Za-z\-']+(?:\s+[A-Z][A-Za-z\-']+)*[,.]", stripped):
        return True
    if re.match(r"^[\u4e00-\u9fff]{1,12}[,，]", stripped):
        return True
    return False


def split_reference_entries(text: str) -> List[str]:
    normalized = (text or "").strip()
    if not normalized:
        return []
    normalized = re.sub(r"^\s*参考文献\s*", "", normalized)
    starts = list(re.finditer(r"\[\s*\d+\s*\]", normalized))
    if not starts:
        return [normalized] if normalized else []
    entries: List[str] = []
    for idx, match in enumerate(starts):
        start = match.start()
        end = starts[idx + 1].start() if idx + 1 < len(starts) else len(normalized)
        chunk = normalized[start:end].strip()
        if chunk:
            entries.append(chunk)
    return entries


def insert_paragraph_after(body: ET.Element, anchor: ET.Element, text: str, source_rpr: Optional[ET.Element]) -> ET.Element:
    new_paragraph = ET.Element(qn("p"))
    append_text_run(new_paragraph, text, source_rpr)
    body.insert(list(body).index(anchor) + 1, new_paragraph)
    return new_paragraph


def strip_leading_numeric_reference_label(paragraph: ET.Element) -> bool:
    texts = paragraph.findall(".//w:t", NS)
    full_text = "".join(node.text or "" for node in texts)
    updated = re.sub(r"^\s*\[\s*\d+\s*\]\s*", "", full_text)
    if updated == full_text:
        return False
    for node in texts:
        node.text = ""
    if texts:
        texts[0].text = updated
        ensure_text_node_preserve(texts[0], updated)
    return True


def normalize_references_section(
    document_root: ET.Element,
    protected_prefix_end: int,
    style_specs: Dict[str, Dict[str, Any]],
    rules: Dict[str, Any],
) -> List[Dict[str, Any]]:
    body = document_root.find("w:body", NS)
    if body is None:
        return []
    paragraphs = paragraph_nodes(body)
    heading_index: Optional[int] = None
    for index, paragraph in enumerate(paragraphs):
        if index < protected_prefix_end:
            continue
        text = (element_text(paragraph) or "").strip().replace(" ", "")
        if text == "参考文献" or text.startswith("参考文献[") or text.startswith("参考文献［"):
            heading_index = index
            break
    if heading_index is None:
        return []

    changes: List[Dict[str, Any]] = []
    apply_style_to_paragraph(paragraphs[heading_index], style_specs["zafu_references_heading"])
    apply_run_defaults(paragraphs[heading_index], style_specs["zafu_references_heading"], preserve_emphasis=False)
    changes.append({"action": "normalize_references_heading", "paragraphIndex": heading_index})
    bibliography_numbering = str(((rules.get("references") or {}).get("bibliography_numbering")) or "").strip().lower()

    heading_text = (element_text(paragraphs[heading_index]) or "").strip()
    heading_text = re.sub(r"\s+", " ", heading_text)
    combined_entries = split_reference_entries(heading_text)
    if combined_entries:
        source_rpr = first_run_rpr(paragraphs[heading_index])
        replace_paragraph_plain_text(paragraphs[heading_index], "参考文献")
        anchor = paragraphs[heading_index]
        inserted = 0
        for entry in combined_entries:
            anchor = insert_paragraph_after(body, anchor, entry, source_rpr)
            inserted += 1
        changes.append(
            {
                "action": "split_combined_references_block",
                "paragraphIndex": heading_index,
                "entryCount": inserted,
            }
        )
        paragraphs = paragraph_nodes(body)

    for index in range(heading_index + 1, len(paragraphs)):
        paragraph = paragraphs[index]
        text = (element_text(paragraph) or "").strip()
        compact = text.replace(" ", "")
        if not text:
            continue
        if compact in REFERENCE_STOP_HEADINGS:
            break
        if re.match(r"^\d+(?:\.\d+)*\s+\S+", text):
            break
        extra_entries = split_reference_entries(text) if text.count("[") > 1 else []
        if extra_entries:
            source_rpr = first_run_rpr(paragraph)
            replace_paragraph_plain_text(paragraph, extra_entries[0])
            anchor = paragraph
            for entry in extra_entries[1:]:
                anchor = insert_paragraph_after(body, anchor, entry, source_rpr)
            paragraphs = paragraph_nodes(body)
            text = (element_text(paragraph) or "").strip()
        spec_key = "zafu_references_en" if re.search(r"[A-Za-z]", text) and not re.search(r"[\u4e00-\u9fff]", text) else "zafu_references_cn"
        spec = style_specs[spec_key]
        numeric_label_removed = False
        if bibliography_numbering == "none":
            numeric_label_removed = strip_leading_numeric_reference_label(paragraph)
        apply_style_to_paragraph(paragraph, spec)
        ppr = paragraph.find("w:pPr", NS)
        if ppr is not None:
            ind = ensure_child(ppr, "ind")
            for attr in ("firstLineChars", "firstLine", "hanging", "left", "leftChars"):
                key = qn(attr)
                if key in ind.attrib:
                    del ind.attrib[key]
        apply_run_defaults(paragraph, spec, preserve_emphasis=False)
        changes.append(
            {
                "action": "normalize_reference_entry",
                "paragraphIndex": index,
                "styleId": spec["styleId"],
                "referenceLike": _looks_like_reference_entry(text),
                "numericLabelRemoved": numeric_label_removed,
            }
        )
    return changes


def summarize_execution_log(execution_log: List[Dict[str, Any]]) -> Dict[str, Any]:
    auto_fixed_actions: List[str] = []
    detected_only: List[str] = []
    blocked_by_policy: List[str] = []
    for item in execution_log:
        if not isinstance(item, dict):
            continue
        action = str(item.get("action") or "")
        if not action:
            continue
        if action.startswith("preserve_"):
            blocked_by_policy.append(action)
        elif action in {"replace_header_text", "ensure_update_fields_on_open", "fix_rel_paths"}:
            auto_fixed_actions.append(action)
        else:
            auto_fixed_actions.append(action)
    return {
        "autoFixedActions": sorted(set(auto_fixed_actions)),
        "detectedOnly": sorted(set(detected_only)),
        "blockedByPolicy": sorted(set(blocked_by_policy)),
    }


def insert_toc_block_if_missing(
    document_root: ET.Element,
    source_document_root: ET.Element,
    source_prefix_paragraph_count: int,
    current_prefix_paragraph_count: int,
    inserted_boundary_paragraph: int,
) -> Dict[str, Any]:
    body = document_root.find("w:body", NS)
    source_body = source_document_root.find("w:body", NS)
    if body is None or source_body is None:
        return {"inserted": False, "reason": "missing_body", "paragraphsInserted": 0}

    source_paragraphs = paragraph_nodes(source_body)
    source_front_candidates = source_paragraphs[source_prefix_paragraph_count : source_prefix_paragraph_count + 12]
    if any(
        paragraph_has_toc_field(node)
        or normalize_front_matter_token(paragraph_text(node)) in {"目录", "在Word中打开后更新目录"}
        for node in source_front_candidates
    ):
        return {"inserted": False, "reason": "source_front_matter_already_contains_toc", "paragraphsInserted": 0}

    current_paragraphs = paragraph_nodes(body)
    for paragraph in current_paragraphs:
        if paragraph_has_toc_field(paragraph) or normalize_front_matter_token(paragraph_text(paragraph)) in {"目录", "在Word中打开后更新目录"}:
            return {"inserted": False, "reason": "document_prefix_already_contains_toc", "paragraphsInserted": 0}

    insert_at = body_children_until_paragraph_count(body, current_prefix_paragraph_count)
    boundary_index: Optional[int] = None
    if inserted_boundary_paragraph:
        candidate_paragraphs = paragraph_nodes(body)
        boundary_paragraph_index = max(current_prefix_paragraph_count - 1, 0)
        if boundary_paragraph_index < len(candidate_paragraphs):
            boundary_index = find_child_index(body, candidate_paragraphs[boundary_paragraph_index])
    if boundary_index is not None:
        # TOC must start after the template front-matter section break so it
        # receives the directory/abstract section geometry instead of cover geometry.
        insert_at = boundary_index + 1

    nodes = [build_toc_title_paragraph(page_break_before=True), build_toc_field_paragraph()]
    for offset, node in enumerate(nodes):
        body.insert(insert_at + offset, node)

    return {
        "inserted": True,
        "reason": "template_front_matter_default_toc_insertion",
        "insertAtChildIndex": insert_at,
        "paragraphsInserted": len(nodes),
    }


def replace_front_matter_from_template(
    source_document_root: ET.Element,
    template_document_root: ET.Element,
    source_prefix_paragraph_count: int,
    template_prefix_paragraph_count: int,
) -> Dict[str, Any]:
    source_body = source_document_root.find("w:body", NS)
    template_body = template_document_root.find("w:body", NS)
    if source_body is None or template_body is None:
        return {
            "changed": False,
            "sourcePrefixParagraphCount": source_prefix_paragraph_count,
            "templatePrefixParagraphCount": template_prefix_paragraph_count,
            "indexShift": 0,
        }

    source_children = list(source_body)
    delete_end = body_children_until_paragraph_count(source_body, source_prefix_paragraph_count)
    template_nodes = template_prefix_children(template_body, template_prefix_paragraph_count)
    clones = [clone_xml(node) for node in template_nodes]

    for child in source_children[:delete_end]:
        source_body.remove(child)

    insert_at = 0
    for node in clones:
        source_body.insert(insert_at, node)
        insert_at += 1

    template_tail_sectpr = template_body.find("w:sectPr", NS)
    inserted_boundary_paragraph = 0
    if template_tail_sectpr is not None and not has_paragraph_level_sectpr(clones):
        boundary_para = ET.Element(qn("p"))
        boundary_ppr = ET.SubElement(boundary_para, qn("pPr"))
        boundary_ppr.append(clone_xml(template_tail_sectpr))
        source_body.insert(insert_at, boundary_para)
        inserted_boundary_paragraph = 1

    final_prefix_paragraph_count = template_prefix_paragraph_count + inserted_boundary_paragraph
    index_shift = final_prefix_paragraph_count - source_prefix_paragraph_count
    return {
        "changed": True,
        "sourcePrefixParagraphCount": source_prefix_paragraph_count,
        "templatePrefixParagraphCount": template_prefix_paragraph_count,
        "finalPrefixParagraphCount": final_prefix_paragraph_count,
        "insertedBoundaryParagraph": inserted_boundary_paragraph,
        "indexShift": index_shift,
    }


def section_nodes(document_root: ET.Element) -> List[Dict[str, Any]]:
    body = document_root.find("w:body", NS)
    results: List[Dict[str, Any]] = []
    if body is None:
        return results
    paragraph_index = 0
    for child in list(body):
        if child.tag == qn("p"):
            ppr = child.find("w:pPr", NS)
            if ppr is not None and ppr.find("w:sectPr", NS) is not None:
                results.append({"paragraphIndex": paragraph_index, "sectPr": ppr.find("w:sectPr", NS)})
            paragraph_index += 1
        elif child.tag == qn("tbl"):
            continue
    tail = body.find("w:sectPr", NS)
    if tail is not None:
        results.append({"paragraphIndex": None, "sectPr": tail})
    return results


def is_probable_body_heading(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if not compact or compact.upper() in {"ABSTRACT"}:
        return False
    if compact in {"目录", "摘要", "致谢", "参考文献", "附录"}:
        return False
    return BODY_HEADING_RE.match(text or "") is not None


def normalize_heading_token(text: str) -> str:
    compact = re.sub(r"\s+", "", (text or "").strip())
    return compact.replace("：", ":").replace("．", ".").replace("。", ".")


def normalize_front_matter_token(text: str) -> str:
    return normalize_heading_token(text).replace("\u3000", "")


def detect_body_start_paragraph(
    body_paragraphs: List[ET.Element],
    plan: Dict[str, Any],
    source_prefix_paragraph_count: int,
    index_shift: int,
    protected_prefix_end: int,
) -> Optional[int]:
    candidates: List[int] = []
    for paragraph_index_str, style_id in (plan.get("styleMapping") or {}).items():
        if style_id != "zafu_heading1":
            continue
        index = safe_int(paragraph_index_str)
        if index is None:
            continue
        if index >= source_prefix_paragraph_count:
            index += index_shift
        if not (protected_prefix_end <= index < len(body_paragraphs)):
            continue
        if is_probable_body_heading(element_text(body_paragraphs[index]).strip()):
            candidates.append(index)
    return min(candidates) if candidates else None


def ensure_section_break_before_paragraph(
    document_root: ET.Element,
    paragraph_index: Optional[int],
    protected_prefix_end: int,
) -> List[Dict[str, Any]]:
    if paragraph_index is None or paragraph_index <= 0:
        return []
    existing_sections = section_nodes(document_root)
    if protected_prefix_end > 0 and len(existing_sections) >= 3:
        return []
    if any(
        section.get("paragraphIndex") is not None and protected_prefix_end <= section["paragraphIndex"] < paragraph_index
        for section in existing_sections
    ):
        return []
    body = document_root.find("w:body", NS)
    if body is None:
        return []
    paragraphs = paragraph_nodes(body)
    if paragraph_index >= len(paragraphs):
        return []
    previous_paragraph = paragraphs[paragraph_index - 1]
    previous_ppr = previous_paragraph.find("w:pPr", NS)
    if previous_ppr is not None and previous_ppr.find("w:sectPr", NS) is not None:
        return []
    body_sectpr = body.find("w:sectPr", NS)
    if body_sectpr is None:
        return []
    if previous_ppr is None:
        previous_ppr = ET.Element(qn("pPr"))
        previous_paragraph.insert(0, previous_ppr)
    previous_ppr.append(clone_xml(body_sectpr))
    return [
        {
            "action": "insert_section_break_before_body",
            "bodyStartParagraphIndex": paragraph_index,
            "boundaryParagraphIndex": paragraph_index - 1,
        }
    ]


def classify_section_roles(
    document_root: ET.Element,
    protected_prefix_end: int,
    body_start_paragraph: Optional[int] = None,
) -> List[Dict[str, Any]]:
    sections = section_nodes(document_root)
    if not sections:
        return []
    roles: List[Dict[str, Any]] = []
    last_index = len(sections) - 1
    for section_index, section in enumerate(sections):
        paragraph_index = section.get("paragraphIndex")
        if protected_prefix_end > 0 and len(sections) >= 3:
            if section_index == 0:
                role = "front_matter"
            elif section_index == last_index:
                role = "body"
            else:
                role = "toc_abstract"
        elif protected_prefix_end > 0 and section_index == 0:
            role = "front_matter"
        elif body_start_paragraph is not None:
            role = "body" if paragraph_index is None or paragraph_index >= body_start_paragraph else "toc_abstract"
        elif section_index == last_index:
            role = "body"
        else:
            role = "toc_abstract"
        roles.append(
            {
                "sectionIndex": section_index,
                "paragraphIndex": paragraph_index,
                "sectPr": section["sectPr"],
                "role": role,
            }
        )
    return roles


def ensure_child(parent: ET.Element, child_tag: str) -> ET.Element:
    child = parent.find(f"w:{child_tag}", NS)
    if child is None:
        child = ET.SubElement(parent, qn(child_tag))
    return child


def set_table_three_line_borders(tblpr: ET.Element, top_size: str = "12", bottom_size: str = "12") -> None:
    borders = ensure_child(tblpr, "tblBorders")
    for child in list(borders):
        borders.remove(child)
    # Keep only the outer top/bottom rules at table level. Header separators are
    # written on cells so Word does not collapse them unpredictably.
    for edge, val, sz in [
        ("top", "single", top_size),
        ("bottom", "single", bottom_size),
        ("insideH", "nil", "0"),
        ("insideV", "nil", "0"),
        ("left", "nil", "0"),
        ("right", "nil", "0"),
    ]:
        node = ET.SubElement(borders, qn(edge))
        node.set(qn("val"), val)
        node.set(qn("sz"), sz)
        node.set(qn("space"), "0")
        node.set(qn("color"), "auto")


def clear_cell_borders_and_shading(tcpr: ET.Element) -> None:
    tc_borders = tcpr.find("w:tcBorders", NS)
    if tc_borders is not None:
        tcpr.remove(tc_borders)
    shading = tcpr.find("w:shd", NS)
    if shading is not None:
        tcpr.remove(shading)


def set_cell_bottom_border(tcpr: ET.Element, size: str = "8") -> None:
    set_cell_border_edge(tcpr, "bottom", size)


def set_cell_border_edge(tcpr: ET.Element, edge: str, size: str) -> None:
    tc_borders = ensure_child(tcpr, "tcBorders")
    target = tc_borders.find(f"w:{edge}", NS)
    if target is None:
        target = ET.SubElement(tc_borders, qn(edge))
    target.set(qn("val"), "single")
    target.set(qn("sz"), str(size))
    target.set(qn("space"), "0")
    target.set(qn("color"), "auto")


def set_cell_top_border(tcpr: ET.Element, size: str = "12") -> None:
    set_cell_border_edge(tcpr, "top", size)


def set_cell_no_border(tcpr: ET.Element, edge: str) -> None:
    tc_borders = ensure_child(tcpr, "tcBorders")
    target = tc_borders.find(f"w:{edge}", NS)
    if target is None:
        target = ET.SubElement(tc_borders, qn(edge))
    target.set(qn("val"), "nil")
    target.set(qn("sz"), "0")
    target.set(qn("space"), "0")
    target.set(qn("color"), "auto")


def set_row_cant_split(trpr: ET.Element) -> None:
    if trpr.find("w:cantSplit", NS) is None:
        ET.SubElement(trpr, qn("cantSplit"))


def row_cells(row: ET.Element) -> List[ET.Element]:
    return row.findall("w:tc", NS)


def cell_is_numeric_like(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if not compact:
        return False
    if re.search(r"[\u4e00-\u9fff]", compact):
        return False
    patterns = [
        r"^[+\-]?\d+(?:[.．]\d+)?%?$",
        r"^[+\-]?\d+(?:[.．]\d+)?(?:\*{1,3})?(?:[（(][+\-]?\d+(?:[.．]\d+)?[)）])$",
        r"^[+\-]?\d+(?:[.．]\d+)?(?:\*{1,3})?[（(]p[<=>]?\d+(?:[.．]\d+)?[)）]$",
        r"^[+\-]?\d+(?:[.．]\d+)?(?:[-—~～]\d+(?:[.．]\d+)?)?$",
    ]
    for pattern in patterns:
        if re.fullmatch(pattern, compact, flags=re.IGNORECASE):
            return True
    return False


def cell_grid_span(cell: ET.Element) -> int:
    tcpr = cell.find("w:tcPr", NS)
    if tcpr is None:
        return 1
    grid_span = tcpr.find("w:gridSpan", NS)
    if grid_span is None:
        return 1
    return safe_int(grid_span.attrib.get(qn("val"))) or 1


def cell_has_vmerge(cell: ET.Element) -> bool:
    tcpr = cell.find("w:tcPr", NS)
    return tcpr is not None and tcpr.find("w:vMerge", NS) is not None


def row_has_structural_merge(row: ET.Element) -> bool:
    for cell in row_cells(row):
        if cell_grid_span(cell) > 1 or cell_has_vmerge(cell):
            return True
    return False


def row_text_profile(row: ET.Element) -> Tuple[int, int, int, int]:
    non_empty = 0
    total_len = 0
    numeric_cells = 0
    text_cells = 0
    for cell in row_cells(row):
        text = (element_text(cell) or "").strip()
        if text:
            non_empty += 1
            total_len += len(re.sub(r"\s+", "", text))
            if cell_is_numeric_like(text):
                numeric_cells += 1
            else:
                text_cells += 1
    return non_empty, total_len, numeric_cells, text_cells


def detect_header_row_count(rows: List[ET.Element], table_rules: Dict[str, Any]) -> int:
    if not rows:
        return 0
    configured = table_rules.get("header_rows")
    if isinstance(configured, int) and configured > 0:
        return min(configured, len(rows))
    if isinstance(configured, str) and configured.strip().lower() in {"strict_first_row", "single", "one"}:
        return 1
    max_header_rows = safe_int(table_rules.get("max_header_rows")) or 3
    max_header_rows = max(1, min(max_header_rows, len(rows)))
    header_count = 1
    for idx, row in enumerate(rows[:max_header_rows]):
        non_empty, total_len, numeric_cells, text_cells = row_text_profile(row)
        if idx == 0:
            if non_empty == 0:
                return 1
            continue
        if non_empty == 0:
            break
        if numeric_cells > 0 and numeric_cells >= max(1, text_cells):
            break
        if text_cells == 0:
            break
        if total_len > 180:
            break
        if row_has_structural_merge(row) and numeric_cells == 0:
            header_count = idx + 1
            continue
        header_count = idx + 1
    return max(1, header_count)


def detect_group_rule_rows(rows: List[ET.Element], header_row_count: int, table_rules: Dict[str, Any]) -> List[int]:
    if header_row_count <= 1:
        return []
    group_header_rule = table_rules.get("group_header_rule") or {}
    if not group_header_rule.get("enabled", False):
        return []
    mode = str(group_header_rule.get("mode") or "auto_when_merged_header_continues")
    if mode != "auto_when_merged_header_continues":
        return []
    split_rows: List[int] = []
    for idx in range(header_row_count - 1):
        current_non_empty, current_len, current_numeric, current_text = row_text_profile(rows[idx])
        next_non_empty, next_len, next_numeric, next_text = row_text_profile(rows[idx + 1])
        if current_non_empty == 0 or next_non_empty == 0:
            continue
        if current_numeric == 0 and next_numeric == 0 and current_text > 0 and next_text > 0:
            if current_len <= 100 and next_len <= 100:
                split_rows.append(idx)
                continue
        if row_has_structural_merge(rows[idx]) and next_numeric == 0:
            split_rows.append(idx)
    return split_rows


def normalize_table_rules(rules: Dict[str, Any]) -> Dict[str, Any]:
    table_rules = dict((rules.get("tables") or {}))
    top_rule = dict(table_rules.get("top_rule") or {})
    header_rule = dict(table_rules.get("header_rule") or {})
    group_header_rule = dict(table_rules.get("group_header_rule") or {})
    bottom_rule = dict(table_rules.get("bottom_rule") or {})
    top_rule.setdefault("enabled", True)
    top_rule.setdefault("size", 12)
    top_rule.setdefault("reinforce_on_first_row_cells", True)
    header_rule.setdefault("enabled", True)
    header_rule.setdefault("size", 8)
    # Strict three-line tables must default to only three horizontal rules:
    # top, header separator, bottom. Extra grouped-header separators are opt-in.
    group_header_rule.setdefault("enabled", False)
    group_header_rule.setdefault("size", 8)
    group_header_rule.setdefault("mode", "auto_when_merged_header_continues")
    bottom_rule.setdefault("enabled", True)
    bottom_rule.setdefault("size", 12)
    bottom_rule.setdefault("reinforce_on_last_row_cells", True)
    table_rules.setdefault("border_style", "research_three_line")
    table_rules.setdefault("caption_position", "above_table")
    table_rules.setdefault("strip_table_style", True)
    table_rules.setdefault("remove_cell_shading", True)
    table_rules.setdefault("remove_vertical_borders", True)
    table_rules.setdefault("preserve_body_row_separators", False)
    table_rules.setdefault("header_rows", "auto")
    table_rules.setdefault("max_header_rows", 3)
    table_rules.setdefault("repeat_group_header_rules", False)
    table_rules["top_rule"] = top_rule
    table_rules["header_rule"] = header_rule
    table_rules["group_header_rule"] = group_header_rule
    table_rules["bottom_rule"] = bottom_rule
    return table_rules


def table_caption_position(table_rules: Dict[str, Any]) -> str:
    position = str(table_rules.get("caption_position") or "above_table").strip().lower()
    if position in {"above", "top", "above_table", "table_above"}:
        return "above_table"
    if position in {"below", "bottom", "below_table", "table_below"}:
        return "below_table"
    return "above_table"


def element_text(node: ET.Element) -> str:
    return "".join(text.text or "" for text in node.findall(".//w:t", NS))


def cm_to_twips(cm: float) -> int:
    return int(round(cm / 2.54 * 1440))


def style_spec_from_rules(rules: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    styles = rules.get("styles") or {}

    def build(rule_key: str, style_id: str, name: str, fallback_line_pt: Optional[float] = None) -> Dict[str, Any]:
        source = styles.get(rule_key) or {}
        fonts = {}
        if source.get("ascii_font"):
            fonts["ascii"] = source["ascii_font"]
            fonts["hAnsi"] = source["ascii_font"]
            fonts["cs"] = source["ascii_font"]
        if source.get("east_asia_font"):
            fonts["eastAsia"] = source["east_asia_font"]
        return {
            "styleId": style_id,
            "name": name,
            "fonts": fonts,
            "sizePt": source.get("size_pt"),
            "bold": source.get("bold"),
            "jc": source.get("align"),
            "outlineLvl": source.get("outline_level"),
            "pageBreakBefore": source.get("page_break_before"),
            "spacing": {
                "beforePt": source.get("spacing_before_pt"),
                "beforeLines": source.get("spacing_before_lines"),
                "afterPt": source.get("spacing_after_pt"),
                "afterLines": source.get("spacing_after_lines"),
                "linePt": source.get("line_spacing_pt", fallback_line_pt),
                "lineMultiple": source.get("line_spacing_multiple"),
                "lineRule": "exact" if source.get("line_spacing_pt", fallback_line_pt) else ("auto" if source.get("line_spacing_multiple") else None),
            },
            "ind": {"firstLineChars": source.get("indent_chars")},
        }

    body_line = (rules.get("body") or {}).get("line_spacing_pt")
    caption = build("body_text", "zafu_caption", "ZAFU Caption", fallback_line_pt=body_line)
    caption["jc"] = "center"
    caption["ind"] = {"firstLineChars": 0}
    keywords_cn = build("keywords_text_cn", "zafu_keywords_cn", "ZAFU Keywords CN", fallback_line_pt=body_line)
    keywords_en = build("keywords_text_en", "zafu_keywords_en", "ZAFU Keywords EN", fallback_line_pt=body_line)
    return {
        "zafu_title_cn": build("thesis_title", "zafu_title_cn", "ZAFU Title CN"),
        "zafu_title_en": build("title_en", "zafu_title_en", "ZAFU Title EN", fallback_line_pt=body_line),
        "zafu_body": build("body_text", "zafu_body", "ZAFU Body", fallback_line_pt=body_line),
        "zafu_heading1": build("heading1", "zafu_heading1", "ZAFU Heading 1"),
        "zafu_heading2": build("heading2", "zafu_heading2", "ZAFU Heading 2"),
        "zafu_heading3": build("heading3", "zafu_heading3", "ZAFU Heading 3"),
        "zafu_abstract_text_cn": build("abstract_text_cn", "zafu_abstract_text_cn", "ZAFU Abstract CN", fallback_line_pt=body_line),
        "zafu_abstract_text_en": build("abstract_text_en", "zafu_abstract_text_en", "ZAFU Abstract EN", fallback_line_pt=body_line),
        "zafu_abstract_label_cn": build("abstract_label_cn", "zafu_abstract_label_cn", "ZAFU Abstract Label CN", fallback_line_pt=body_line),
        "zafu_abstract_label_en": build("abstract_label_en", "zafu_abstract_label_en", "ZAFU Abstract Label EN", fallback_line_pt=body_line),
        "zafu_keywords_label_cn": build("keywords_label_cn", "zafu_keywords_label_cn", "ZAFU Keywords Label CN", fallback_line_pt=body_line),
        "zafu_keywords_label_en": build("keywords_label_en", "zafu_keywords_label_en", "ZAFU Keywords Label EN", fallback_line_pt=body_line),
        "zafu_keywords_cn": keywords_cn,
        "zafu_keywords_en": keywords_en,
        "zafu_caption": caption,
        "zafu_toc_title": build("toc_title", "zafu_toc_title", "ZAFU TOC Title", fallback_line_pt=body_line),
        "zafu_toc_body": build("toc_body", "zafu_toc_body", "ZAFU TOC Body", fallback_line_pt=body_line),
        "zafu_references_heading": build("references_heading", "zafu_references_heading", "ZAFU References Heading"),
        "zafu_references_cn": build("references_cn", "zafu_references_cn", "ZAFU References CN", fallback_line_pt=body_line),
        "zafu_references_en": build("references_en", "zafu_references_en", "ZAFU References EN", fallback_line_pt=body_line),
    }


def apply_spacing(ppr: ET.Element, spacing_spec: Dict[str, Any]) -> None:
    spacing = ensure_child(ppr, "spacing")
    if spacing_spec.get("beforePt") is not None:
        spacing.set(qn("before"), str(int(round(spacing_spec["beforePt"] * 20))))
    elif qn("before") in spacing.attrib:
        del spacing.attrib[qn("before")]
    if spacing_spec.get("beforeLines") is not None:
        spacing.set(qn("beforeLines"), str(int(spacing_spec["beforeLines"])))
    elif qn("beforeLines") in spacing.attrib:
        del spacing.attrib[qn("beforeLines")]

    if spacing_spec.get("afterPt") is not None:
        spacing.set(qn("after"), str(int(round(spacing_spec["afterPt"] * 20))))
    elif qn("after") in spacing.attrib:
        del spacing.attrib[qn("after")]
    if spacing_spec.get("afterLines") is not None:
        spacing.set(qn("afterLines"), str(int(spacing_spec["afterLines"])))
    elif qn("afterLines") in spacing.attrib:
        del spacing.attrib[qn("afterLines")]

    if spacing_spec.get("lineMultiple") is not None:
        spacing.set(qn("line"), str(int(round(float(spacing_spec["lineMultiple"]) * 240))))
        spacing.set(qn("lineRule"), "auto")
    elif spacing_spec.get("linePt") is not None:
        spacing.set(qn("line"), str(int(round(spacing_spec["linePt"] * 20))))
        if spacing_spec.get("lineRule"):
            spacing.set(qn("lineRule"), spacing_spec["lineRule"])
        elif qn("lineRule") in spacing.attrib:
            del spacing.attrib[qn("lineRule")]
    else:
        if qn("line") in spacing.attrib:
            del spacing.attrib[qn("line")]
        if qn("lineRule") in spacing.attrib:
            del spacing.attrib[qn("lineRule")]


def ensure_style(styles_root: ET.Element, spec: Dict[str, Any]) -> None:
    style = None
    for node in styles_root.findall("w:style", NS):
        if node.attrib.get(qn("styleId")) == spec["styleId"]:
            style = node
            break
    if style is None:
        style = ET.SubElement(styles_root, qn("style"))
        style.set(qn("type"), "paragraph")
        style.set(qn("styleId"), spec["styleId"])

    name = ensure_child(style, "name")
    name.set(qn("val"), spec["name"])

    based_on = style.find("w:basedOn", NS)
    if based_on is None:
        based_on = ET.SubElement(style, qn("basedOn"))
    based_on.set(qn("val"), "a")

    next_node = style.find("w:next", NS)
    if next_node is None:
        next_node = ET.SubElement(style, qn("next"))
    next_node.set(qn("val"), "a")

    ppr = ensure_child(style, "pPr")
    spacing_spec = spec.get("spacing") or {}
    if spacing_spec:
        apply_spacing(ppr, spacing_spec)
    if spec.get("jc"):
        jc = ensure_child(ppr, "jc")
        jc.set(qn("val"), spec["jc"])
    if spec.get("outlineLvl") is not None:
        outline = ensure_child(ppr, "outlineLvl")
        outline.set(qn("val"), str(int(spec["outlineLvl"])))
    page_break_before = ppr.find("w:pageBreakBefore", NS)
    if spec.get("pageBreakBefore") is True:
        if page_break_before is None:
            ET.SubElement(ppr, qn("pageBreakBefore"))
    elif spec.get("pageBreakBefore") is False and page_break_before is not None:
        ppr.remove(page_break_before)
    ind_spec = spec.get("ind") or {}
    if ind_spec.get("firstLineChars") is not None:
        ind = ensure_child(ppr, "ind")
        ind.set(qn("firstLineChars"), str(ind_spec["firstLineChars"] * 100))

    rpr = ensure_child(style, "rPr")
    rfonts = ensure_child(rpr, "rFonts")
    for key, value in (spec.get("fonts") or {}).items():
        rfonts.set(qn(key), value)
    if spec.get("sizePt") is not None:
        size = ensure_child(rpr, "sz")
        size.set(qn("val"), str(pt_to_half_points(spec["sizePt"])))
        size_cs = ensure_child(rpr, "szCs")
        size_cs.set(qn("val"), str(pt_to_half_points(spec["sizePt"])))
    bold = rpr.find("w:b", NS)
    if spec.get("bold") is True:
        if bold is None:
            ET.SubElement(rpr, qn("b"))
    elif spec.get("bold") is False and bold is not None:
        rpr.remove(bold)


def apply_style_to_paragraph(paragraph: ET.Element, spec: Dict[str, Any]) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        pstyle = ET.SubElement(ppr, qn("pStyle"))
    pstyle.set(qn("val"), spec["styleId"])

    spacing_spec = spec.get("spacing") or {}
    if spacing_spec:
        apply_spacing(ppr, spacing_spec)
    if spec.get("jc"):
        jc = ensure_child(ppr, "jc")
        jc.set(qn("val"), spec["jc"])
    if spec.get("outlineLvl") is not None:
        outline = ensure_child(ppr, "outlineLvl")
        outline.set(qn("val"), str(int(spec["outlineLvl"])))
    page_break_before = ppr.find("w:pageBreakBefore", NS)
    if spec.get("pageBreakBefore") is True:
        if page_break_before is None:
            ET.SubElement(ppr, qn("pageBreakBefore"))
    elif spec.get("pageBreakBefore") is False and page_break_before is not None:
        ppr.remove(page_break_before)
    ind_spec = spec.get("ind") or {}
    if ind_spec.get("firstLineChars") is not None:
        ind = ensure_child(ppr, "ind")
        ind.set(qn("firstLineChars"), str(ind_spec["firstLineChars"] * 100))


def clear_paragraph_heading_semantics(paragraph: ET.Element) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return
    for tag in ("outlineLvl", "pageBreakBefore"):
        node = ppr.find(f"w:{tag}", NS)
        if node is not None:
            ppr.remove(node)
    ind = ppr.find("w:ind", NS)
    if ind is not None:
        for attr in ("firstLineChars", "firstLine", "left", "hanging"):
            key = qn(attr)
            if key in ind.attrib:
                del ind.attrib[key]


def paragraph_style_id(paragraph: ET.Element) -> Optional[str]:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.attrib.get(qn("val"))


def ensure_builtin_toc_styles(styles_root: ET.Element, rules: Dict[str, Any]) -> List[Dict[str, Any]]:
    toc = (rules.get("styles") or {}).get("toc_body") or {}
    fonts = {}
    if toc.get("ascii_font"):
        fonts["ascii"] = toc["ascii_font"]
        fonts["hAnsi"] = toc["ascii_font"]
        fonts["cs"] = toc["ascii_font"]
    if toc.get("east_asia_font"):
        fonts["eastAsia"] = toc["east_asia_font"]
    line_multiple = toc.get("line_spacing_multiple")
    specs = [
        {"styleId": "29", "name": "toc 1", "left": None, "bold": True},
        {"styleId": "32", "name": "toc 2", "left": "420", "bold": False},
        {"styleId": "25", "name": "toc 3", "left": "840", "bold": False},
    ]
    changes: List[Dict[str, Any]] = []
    for toc_spec in specs:
        style = None
        for node in styles_root.findall("w:style", NS):
            if node.attrib.get(qn("styleId")) == toc_spec["styleId"]:
                style = node
                break
        if style is None:
            style = ET.SubElement(styles_root, qn("style"))
            style.set(qn("type"), "paragraph")
            style.set(qn("styleId"), toc_spec["styleId"])
        name = ensure_child(style, "name")
        name.set(qn("val"), toc_spec["name"])
        based_on = ensure_child(style, "basedOn")
        based_on.set(qn("val"), "1")
        next_node = ensure_child(style, "next")
        next_node.set(qn("val"), "1")
        ppr = ensure_child(style, "pPr")
        tabs = ensure_child(ppr, "tabs")
        for tab in list(tabs.findall("w:tab", NS)):
            tabs.remove(tab)
        new_tab = ET.SubElement(tabs, qn("tab"))
        new_tab.set(qn("val"), "right")
        new_tab.set(qn("leader"), "dot")
        new_tab.set(qn("pos"), "9629")
        spacing = ensure_child(ppr, "spacing")
        if line_multiple is not None:
            spacing.set(qn("line"), str(int(round(float(line_multiple) * 240))))
            spacing.set(qn("lineRule"), "auto")
        jc = ensure_child(ppr, "jc")
        jc.set(qn("val"), "left")
        ind = ensure_child(ppr, "ind")
        if toc_spec["left"] is not None:
            ind.set(qn("left"), toc_spec["left"])
        elif qn("left") in ind.attrib:
            del ind.attrib[qn("left")]
        rpr = ensure_child(style, "rPr")
        rfonts = ensure_child(rpr, "rFonts")
        for key, value in fonts.items():
            rfonts.set(qn(key), value)
        if toc.get("size_pt") is not None:
            size = ensure_child(rpr, "sz")
            size.set(qn("val"), str(pt_to_half_points(toc["size_pt"])))
            size_cs = ensure_child(rpr, "szCs")
            size_cs.set(qn("val"), str(pt_to_half_points(toc["size_pt"])))
        bold = rpr.find("w:b", NS)
        if toc_spec["bold"]:
            if bold is None:
                ET.SubElement(rpr, qn("b"))
        elif bold is not None:
            rpr.remove(bold)
        changes.append(
            {
                "action": "ensure_builtin_toc_style",
                "styleId": toc_spec["styleId"],
                "name": toc_spec["name"],
                "left": toc_spec["left"],
                "lineMultiple": line_multiple,
            }
        )
    return changes


def apply_run_defaults(paragraph: ET.Element, spec: Dict[str, Any], preserve_emphasis: bool) -> int:
    updated = 0
    for run in paragraph.findall(".//w:r", NS):
        texts = run.findall(".//w:t", NS)
        if not any((t.text or "").strip() for t in texts):
            continue
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            rpr = ET.Element(qn("rPr"))
            run.insert(0, rpr)
        rfonts = ensure_child(rpr, "rFonts")
        for key, value in (spec.get("fonts") or {}).items():
            rfonts.set(qn(key), value)
        if spec.get("sizePt") is not None:
            size = ensure_child(rpr, "sz")
            size.set(qn("val"), str(pt_to_half_points(spec["sizePt"])))
            size_cs = ensure_child(rpr, "szCs")
            size_cs.set(qn("val"), str(pt_to_half_points(spec["sizePt"])))
        if not preserve_emphasis:
            bold = rpr.find("w:b", NS)
            if spec.get("bold") is True:
                if bold is None:
                    ET.SubElement(rpr, qn("b"))
            elif spec.get("bold") is False and bold is not None:
                rpr.remove(bold)
        updated += 1
    return updated


def apply_spec_to_run(run: ET.Element, spec: Dict[str, Any], preserve_emphasis: bool = False) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(qn("rPr"))
        run.insert(0, rpr)
    rfonts = ensure_child(rpr, "rFonts")
    for key, value in (spec.get("fonts") or {}).items():
        rfonts.set(qn(key), value)
    if spec.get("sizePt") is not None:
        size = ensure_child(rpr, "sz")
        size.set(qn("val"), str(pt_to_half_points(spec["sizePt"])))
        size_cs = ensure_child(rpr, "szCs")
        size_cs.set(qn("val"), str(pt_to_half_points(spec["sizePt"])))
    if not preserve_emphasis:
        if spec.get("bold") is True:
            if rpr.find("w:b", NS) is None:
                ET.SubElement(rpr, qn("b"))
        elif spec.get("bold") is False:
            remove_run_bold(run)


def remove_run_bold(run: ET.Element) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return
    for tag in ("b", "bCs"):
        bold = rpr.find(f"w:{tag}", NS)
        if bold is not None:
            rpr.remove(bold)


def direct_runs(paragraph: ET.Element) -> List[ET.Element]:
    return [child for child in list(paragraph) if child.tag == qn("r")]


def split_first_run_with_prefix(paragraph: ET.Element, prefix: str) -> Tuple[Optional[ET.Element], Optional[ET.Element]]:
    if not prefix:
        return None, None
    runs = direct_runs(paragraph)
    for idx, run in enumerate(runs):
        text_nodes = run.findall(".//w:t", NS)
        if not text_nodes:
            continue
        full_text = "".join(node.text or "" for node in text_nodes)
        if not full_text.startswith(prefix):
            continue
        remainder = full_text[len(prefix):]
        text_nodes[0].text = prefix
        for extra in text_nodes[1:]:
            extra.text = ""
        if not remainder:
            return run, None
        new_run = clone_xml(run)
        new_text_nodes = new_run.findall(".//w:t", NS)
        if not new_text_nodes:
            new_text_nodes = [ET.SubElement(new_run, qn("t"))]
        new_text_nodes[0].text = remainder
        for extra in new_text_nodes[1:]:
            extra.text = ""
        paragraph.insert(list(paragraph).index(run) + 1, new_run)
        return run, new_run
    return None, None


def apply_label_content_format(
    paragraph: ET.Element,
    label_prefixes: List[str],
    label_spec: Dict[str, Any],
    content_spec: Dict[str, Any],
) -> bool:
    apply_style_to_paragraph(paragraph, content_spec)
    apply_run_defaults(paragraph, content_spec, preserve_emphasis=False)
    paragraph_text = "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()
    for prefix in label_prefixes:
        if not paragraph_text.startswith(prefix):
            continue
        label_run, content_run = split_first_run_with_prefix(paragraph, prefix)
        if label_run is not None:
            apply_spec_to_run(label_run, label_spec, preserve_emphasis=False)
        if content_run is not None:
            apply_spec_to_run(content_run, content_spec, preserve_emphasis=False)
            if content_spec.get("bold") is not True:
                remove_run_bold(content_run)
        return True
    return False


def ensure_text_node_preserve(text_node: ET.Element, text: str) -> None:
    xml_space = "{http://www.w3.org/XML/1998/namespace}space"
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        text_node.set(xml_space, "preserve")
    elif xml_space in text_node.attrib:
        del text_node.attrib[xml_space]


def set_paragraph_plain_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        run = ET.Element(qn("r"))
        text_node = ET.SubElement(run, qn("t"))
        paragraph.append(run)
        text_nodes = [text_node]
    text_nodes[0].text = text
    ensure_text_node_preserve(text_nodes[0], text)
    for extra in text_nodes[1:]:
        extra.text = ""
        ensure_text_node_preserve(extra, "")


def normalize_manual_leading_whitespace(paragraph: ET.Element) -> bool:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        return False
    original_text = "".join(node.text or "" for node in text_nodes)
    normalized_text = original_text.lstrip(" \t\u3000")
    if normalized_text == original_text or not normalized_text:
        return False
    set_paragraph_plain_text(paragraph, normalized_text)
    return True


def normalize_heading_spacing(paragraph: ET.Element, style_id: str) -> bool:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        return False
    original_text = "".join(node.text or "" for node in text_nodes)
    stripped_text = original_text.strip()
    manual = parse_manual_numbering(stripped_text)
    if not manual:
        return False
    prefix = manual.get("prefix") or ""
    if not prefix or not stripped_text.startswith(prefix):
        return False
    title_text = stripped_text[len(prefix):].lstrip(" \u3000")
    lead_spaces = ""
    separator = " "
    if style_id in {"zafu_heading2", "zafu_heading3"}:
        lead_spaces = "    "
        separator = "  "
    new_text = f"{lead_spaces}{prefix}{separator}{title_text}"
    if new_text == original_text:
        return False
    set_paragraph_plain_text(paragraph, new_text)
    return True


def replace_prefix(paragraph: ET.Element, old_prefix: str, new_prefix: str) -> bool:
    text_nodes = [node for node in paragraph.findall(".//w:t", NS)]
    if not text_nodes:
        return False
    paragraph_text = "".join(node.text or "" for node in text_nodes)
    leading_ws_len = len(paragraph_text) - len(paragraph_text.lstrip())
    leading_ws = paragraph_text[:leading_ws_len]
    trimmed_text = paragraph_text[leading_ws_len:]
    old_prefix_trimmed = (old_prefix or "").strip()
    new_prefix_trimmed = (new_prefix or "").strip()
    if not old_prefix_trimmed:
        return False
    old_pattern = re.escape(old_prefix_trimmed).replace(r"\ ", r"\s+")
    match = re.match(rf"^(?P<prefix>{old_pattern})(?P<gap>\s*)(?P<tail>.*)$", trimmed_text)
    if not match:
        return False
    replacement = f"{leading_ws}{new_prefix_trimmed}"
    if match.group("tail"):
        replacement += f" {match.group('tail')}"
    updated_text = replacement
    if updated_text == paragraph_text:
        return False
    text_nodes[0].text = updated_text
    for extra in text_nodes[1:]:
        extra.text = ""
    return True


def ensure_paragraph_centered(paragraph: ET.Element) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    jc = ppr.find("w:jc", NS)
    if jc is None:
        jc = ET.SubElement(ppr, qn("jc"))
    jc.set(qn("val"), "center")


def split_mixed_drawing_paragraph(paragraph: ET.Element) -> Optional[ET.Element]:
    drawing_runs = [run for run in paragraph.findall("w:r", NS) if run.find("w:drawing", NS) is not None]
    if not drawing_runs:
        return None
    all_runs = paragraph.findall("w:r", NS)
    if len(drawing_runs) == len(all_runs):
        ensure_paragraph_centered(paragraph)
        return paragraph
    new_paragraph = ET.Element(qn("p"))
    original_ppr = paragraph.find("w:pPr", NS)
    if original_ppr is not None:
        new_ppr = clone_xml(original_ppr)
        new_paragraph.append(new_ppr)
    ensure_paragraph_centered(new_paragraph)
    for run in drawing_runs:
        paragraph.remove(run)
        new_paragraph.append(run)
    return new_paragraph


def normalize_caption_adjacent_figures(body: ET.Element, protected_prefix_end: int) -> List[Dict[str, Any]]:
    logs: List[Dict[str, Any]] = []
    changed = True
    while changed:
        changed = False
        paragraphs = paragraph_nodes(body)
        for index, paragraph in enumerate(paragraphs):
            if index < protected_prefix_end:
                continue
            text = paragraph_text(paragraph).strip()
            if not CAPTION_REF_RE.match(text):
                continue
            previous_index = index - 1
            if previous_index < 0:
                continue
            previous = paragraphs[previous_index]
            previous_drawing_count = len(previous.findall(".//w:drawing", NS))
            if previous_drawing_count > 0:
                ensure_paragraph_centered(previous)
                moved = split_mixed_drawing_paragraph(previous)
                if moved is not None and moved is not previous:
                    body_children = list(body)
                    prev_child_index = find_child_index(body, previous)
                    if prev_child_index is None:
                        continue
                    body.insert(prev_child_index + 1, moved)
                    logs.append(
                        {
                            "action": "split_mixed_drawing_paragraph",
                            "paragraphIndex": previous_index,
                            "captionIndex": index,
                            "drawingCount": previous_drawing_count,
                        }
                    )
                    changed = True
                    break
                logs.append(
                    {
                        "action": "center_figure_paragraph",
                        "paragraphIndex": previous_index,
                        "captionIndex": index,
                        "drawingCount": previous_drawing_count,
                    }
                )
                continue
            scan_index = previous_index
            while scan_index >= protected_prefix_end:
                candidate = paragraphs[scan_index]
                candidate_text = paragraph_text(candidate).strip()
                candidate_drawing_count = len(candidate.findall(".//w:drawing", NS))
                if candidate_drawing_count > 0:
                    body_children = list(body)
                    candidate_child_index = find_child_index(body, candidate)
                    caption_child_index = find_child_index(body, paragraph)
                    if candidate_child_index is None or caption_child_index is None:
                        break
                    if candidate_child_index != caption_child_index - 1:
                        body.remove(candidate)
                        body.insert(caption_child_index, candidate)
                        ensure_paragraph_centered(candidate)
                        logs.append(
                            {
                                "action": "move_figure_paragraph_before_caption",
                                "fromParagraphIndex": scan_index,
                                "captionIndex": index,
                                "drawingCount": candidate_drawing_count,
                            }
                        )
                        changed = True
                    else:
                        ensure_paragraph_centered(candidate)
                    break
                if candidate_text:
                    break
                scan_index -= 1
            if changed:
                break
    return logs


def section_relationship_targets(files: Dict[str, bytes], rel_type_suffix: str) -> Dict[str, str]:
    rels_xml = files.get("word/_rels/document.xml.rels")
    if not rels_xml:
        return {}
    root = ET.fromstring(rels_xml)
    targets: Dict[str, str] = {}
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rel_type = rel.attrib.get("Type", "")
        if not rel_type.endswith(rel_type_suffix):
            continue
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if not rel_id or not target:
            continue
        targets[normalize_rel_target_word("word/document.xml", target)] = rel_id
    return targets


def relationship_targets_by_id(files: Dict[str, bytes]) -> Dict[str, str]:
    rels_xml = files.get("word/_rels/document.xml.rels")
    if not rels_xml:
        return {}
    root = ET.fromstring(rels_xml)
    targets: Dict[str, str] = {}
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if not rel_id or not target:
            continue
        targets[rel_id] = normalize_rel_target_word("word/document.xml", target)
    return targets


def ensure_section_reference(sectpr: ET.Element, ref_tag: str, ref_type: str, rel_id: str) -> bool:
    changed = False
    target_node = None
    for node in sectpr.findall(f"w:{ref_tag}", NS):
        if node.attrib.get(qn("type"), "default") == ref_type:
            target_node = node
            break
    if target_node is None:
        insert_at = 0
        for index, child in enumerate(list(sectpr)):
            if child.tag in {qn("headerReference"), qn("footerReference")}:
                insert_at = index + 1
        target_node = ET.Element(qn(ref_tag))
        target_node.set(qn("type"), ref_type)
        sectpr.insert(insert_at, target_node)
        changed = True
    attr_name = f"{{{NS['r']}}}id"
    if target_node.attrib.get(attr_name) != rel_id:
        target_node.set(attr_name, rel_id)
        changed = True
    return changed


def normalize_front_author_paragraph(paragraph: ET.Element, body_spec: Dict[str, Any]) -> None:
    apply_style_to_paragraph(paragraph, body_spec)
    apply_run_defaults(paragraph, body_spec, preserve_emphasis=False)
    clear_paragraph_heading_semantics(paragraph)
    ensure_paragraph_centered(paragraph)


def normalize_runs_by_existing_styles(
    body_paragraphs: List[ET.Element],
    style_specs: Dict[str, Dict[str, Any]],
    protected_prefix_end: int,
) -> List[Dict[str, Any]]:
    execution_log: List[Dict[str, Any]] = []
    for index, paragraph in enumerate(body_paragraphs):
        if index < protected_prefix_end:
            continue
        style_id = paragraph_style_id(paragraph)
        if not style_id:
            continue
        spec = style_specs.get(style_id)
        if not spec:
            continue
        updated_runs = apply_run_defaults(paragraph, spec, preserve_emphasis=(style_id == "zafu_body"))
        leading_whitespace_normalized = False
        if style_id in {"zafu_body", "zafu_caption"}:
            leading_whitespace_normalized = normalize_manual_leading_whitespace(paragraph)
        if updated_runs or leading_whitespace_normalized:
            execution_log.append(
                {
                    "action": "normalize_runs_by_existing_style",
                    "paragraphIndex": index,
                    "styleId": style_id,
                    "runUpdates": updated_runs,
                    "leadingWhitespaceNormalized": leading_whitespace_normalized,
                }
            )
    return execution_log


def normalize_body_style_fallback(
    body_paragraphs: List[ET.Element],
    style_specs: Dict[str, Dict[str, Any]],
    protected_prefix_end: int,
    body_start_paragraph: Optional[int],
) -> List[Dict[str, Any]]:
    execution_log: List[Dict[str, Any]] = []
    body_spec = style_specs.get("zafu_body")
    if body_spec is None:
        return execution_log
    start = body_start_paragraph if body_start_paragraph is not None else protected_prefix_end
    for index, paragraph in enumerate(body_paragraphs):
        if index < start:
            continue
        text = (paragraph_text(paragraph) or "").strip()
        if not text:
            continue
        style_id = paragraph_style_id(paragraph)
        if style_id in {
            "zafu_heading1",
            "zafu_heading2",
            "zafu_heading3",
            "zafu_caption",
            "zafu_references_heading",
            "zafu_references_cn",
            "zafu_references_en",
            "zafu_title_cn",
            "zafu_title_en",
            "zafu_abstract_label_cn",
            "zafu_abstract_label_en",
            "zafu_keywords_label_cn",
            "zafu_keywords_label_en",
            "zafu_abstract_text_cn",
            "zafu_abstract_text_en",
            "zafu_keywords_cn",
            "zafu_keywords_en",
            "zafu_toc_title",
            "zafu_toc_body",
        }:
            continue
        if style_id == "zafu_body":
            continue
        if style_id not in BODY_HEADING_STYLE_IDS or paragraph_looks_like_formula(text) or style_id is None:
            apply_style_to_paragraph(paragraph, body_spec)
            updated_runs = apply_run_defaults(paragraph, body_spec, preserve_emphasis=False)
            leading_whitespace_normalized = normalize_manual_leading_whitespace(paragraph)
            execution_log.append(
                {
                    "action": "normalize_body_style_fallback",
                    "paragraphIndex": index,
                    "styleId": style_id or "None",
                    "runUpdates": updated_runs,
                    "leadingWhitespaceNormalized": leading_whitespace_normalized,
                }
            )
    return execution_log


def normalize_decimal_heading_prefixes(body_paragraphs: List[ET.Element], protected_prefix_end: int) -> List[Dict[str, Any]]:
    execution_log: List[Dict[str, Any]] = []
    counters = [0, 0, 0]
    style_levels = {"zafu_heading1": 1, "zafu_heading2": 2, "zafu_heading3": 3}
    for index, paragraph in enumerate(body_paragraphs):
        if index < protected_prefix_end:
            continue
        style_id = paragraph_style_id(paragraph)
        level = style_levels.get(style_id or "")
        if not level:
            continue
        text = (paragraph_text(paragraph) or "").strip()
        manual = parse_manual_numbering(text)
        if not manual or manual.get("family") != "science_decimal":
            continue
        counters[level - 1] += 1
        for deeper in range(level, len(counters)):
            counters[deeper] = 0
        expected_prefix = ".".join(str(value) for value in counters[:level] if value > 0)
        old_prefix = f"{manual.get('prefix') or ''}{manual.get('separator') or ''}"
        new_prefix = f"{expected_prefix}  "
        changed = replace_prefix(paragraph, old_prefix, new_prefix)
        if changed:
            execution_log.append(
                {
                    "action": "normalize_decimal_heading_prefix",
                    "paragraphIndex": index,
                    "styleId": style_id,
                    "oldPrefix": old_prefix,
                    "newPrefix": new_prefix,
                }
            )
    return execution_log


def remove_section_reference(sectpr: ET.Element, ref_tag: str, ref_type: str) -> bool:
    changed = False
    for node in list(sectpr.findall(f"w:{ref_tag}", NS)):
        if node.attrib.get(qn("type"), "default") == ref_type:
            sectpr.remove(node)
            changed = True
    return changed


def set_section_page_numbering(sectpr: ET.Element, start: Optional[int] = None, fmt: Optional[str] = None) -> bool:
    pg_num_type = sectpr.find("w:pgNumType", NS)
    if start is None and fmt is None:
        if pg_num_type is not None:
            sectpr.remove(pg_num_type)
            return True
        return False
    changed = False
    if pg_num_type is None:
        pg_num_type = ET.SubElement(sectpr, qn("pgNumType"))
        changed = True
    if start is None:
        if qn("start") in pg_num_type.attrib:
            del pg_num_type.attrib[qn("start")]
            changed = True
    elif pg_num_type.attrib.get(qn("start")) != str(start):
        pg_num_type.set(qn("start"), str(start))
        changed = True
    if fmt is None:
        if qn("fmt") in pg_num_type.attrib:
            del pg_num_type.attrib[qn("fmt")]
            changed = True
    elif pg_num_type.attrib.get(qn("fmt")) != fmt:
        pg_num_type.set(qn("fmt"), fmt)
        changed = True
    return changed


def set_section_title_page(sectpr: ET.Element, enabled: bool) -> bool:
    title_pg = sectpr.find("w:titlePg", NS)
    if enabled:
        if title_pg is None:
            sectpr.append(ET.Element(qn("titlePg")))
            return True
        return False
    if title_pg is not None:
        sectpr.remove(title_pg)
        return True
    return False


def set_section_type_value(sectpr: ET.Element, value: Optional[str]) -> bool:
    section_type = sectpr.find("w:type", NS)
    if not value:
        if section_type is not None:
            sectpr.remove(section_type)
            return True
        return False
    if section_type is None:
        section_type = ET.Element(qn("type"))
        sectpr.insert(0, section_type)
        section_type.set(qn("val"), value)
        return True
    if section_type.attrib.get(qn("val")) != value:
        section_type.set(qn("val"), value)
        return True
    return False


def set_section_columns(sectpr: ET.Element, count: Optional[int], space: Optional[int]) -> bool:
    cols = sectpr.find("w:cols", NS)
    if count is None and space is None:
        if cols is not None:
            sectpr.remove(cols)
            return True
        return False
    changed = False
    if cols is None:
        cols = ET.SubElement(sectpr, qn("cols"))
        changed = True
    if count is None:
        if qn("num") in cols.attrib:
            del cols.attrib[qn("num")]
            changed = True
    elif cols.attrib.get(qn("num")) != str(count):
        cols.set(qn("num"), str(count))
        changed = True
    if space is None:
        if qn("space") in cols.attrib:
            del cols.attrib[qn("space")]
            changed = True
    elif cols.attrib.get(qn("space")) != str(space):
        cols.set(qn("space"), str(space))
        changed = True
    return changed


def extract_template_section_reference_patterns(template_files: Dict[str, bytes]) -> List[Dict[str, Any]]:
    document_xml = template_files.get("word/document.xml")
    if not document_xml:
        return []
    document_root = ET.fromstring(document_xml)
    rel_targets = relationship_targets_by_id(template_files)
    patterns: List[Dict[str, Any]] = []
    for section in section_nodes(document_root):
        sectpr = section["sectPr"]
        header_refs = []
        footer_refs = []
        for ref_tag, container in [("headerReference", header_refs), ("footerReference", footer_refs)]:
            for ref in sectpr.findall(f"w:{ref_tag}", NS):
                rel_id = ref.attrib.get(f"{{{NS['r']}}}id")
                container.append(
                    {
                        "type": ref.attrib.get(qn("type"), "default"),
                        "target": rel_targets.get(rel_id or ""),
                    }
                )
        patterns.append(
            {
                "paragraphIndex": section.get("paragraphIndex"),
                "headerReferences": header_refs,
                "footerReferences": footer_refs,
            }
        )
    return patterns


def apply_template_section_reference_pattern(
    sectpr: ET.Element,
    ref_tag: str,
    desired_refs: List[Dict[str, Any]],
    target_rel_map: Dict[str, str],
) -> Dict[str, Any]:
    desired_by_type: Dict[str, str] = {}
    desired_types: Set[str] = set()
    for ref in desired_refs:
        ref_type = str(ref.get("type") or "default")
        desired_types.add(ref_type)
        target = ref.get("target")
        if target:
            rel_id = target_rel_map.get(str(target))
            if rel_id:
                desired_by_type[ref_type] = rel_id
    removed = 0
    ensured = 0
    changed = False
    for node in list(sectpr.findall(f"w:{ref_tag}", NS)):
        ref_type = node.attrib.get(qn("type"), "default")
        if ref_type not in desired_types or ref_type not in desired_by_type:
            sectpr.remove(node)
            removed += 1
            changed = True
    for ref_type, rel_id in desired_by_type.items():
        if ensure_section_reference(sectpr, ref_tag, ref_type, rel_id):
            ensured += 1
            changed = True
    return {"changed": changed, "removed": removed, "ensured": ensured}


def apply_section_page_numbering(
    document_root: ET.Element,
    protected_prefix_end: int,
    body_start_paragraph: Optional[int] = None,
) -> List[Dict[str, Any]]:
    changes: List[Dict[str, Any]] = []
    toc_started = False
    body_started = False
    for section in classify_section_roles(document_root, protected_prefix_end, body_start_paragraph):
        start: Optional[int] = None
        fmt: Optional[str] = None
        if section["role"] == "toc_abstract":
            fmt = "upperRoman"
            if not toc_started:
                start = 1
                toc_started = True
        elif section["role"] == "body":
            fmt = "decimal"
            if not body_started:
                start = 1
                body_started = True
        else:
            continue
        if set_section_page_numbering(section["sectPr"], start=start, fmt=fmt):
            changes.append(
                {
                    "action": "apply_section_page_numbering",
                    "sectionIndex": section["sectionIndex"],
                    "paragraphIndex": section["paragraphIndex"],
                    "sectionRole": section["role"],
                    "start": start,
                    "format": fmt,
                }
            )
    return changes


def ensure_body_header_references(
    document_root: ET.Element,
    files: Dict[str, bytes],
    protected_prefix_end: int,
    body_start_paragraph: Optional[int] = None,
    template_section_patterns: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    header_targets = section_relationship_targets(files, "/header")
    footer_targets = section_relationship_targets(files, "/footer")
    if template_section_patterns:
        changes: List[Dict[str, Any]] = []
        section_header_rel_id = header_targets.get("word/header2.xml") or header_targets.get("word/header1.xml")
        for section in classify_section_roles(document_root, protected_prefix_end, body_start_paragraph):
            template_pattern = template_section_info_for_role(template_section_patterns, section["role"])
            sectpr = section["sectPr"]
            header_result = apply_template_section_reference_pattern(
                sectpr,
                "headerReference",
                template_pattern.get("headerReferences") or [],
                header_targets,
            )
            footer_result = apply_template_section_reference_pattern(
                sectpr,
                "footerReference",
                template_pattern.get("footerReferences") or [],
                footer_targets,
            )
            first_header_forced = False
            default_header_forced = False
            if section["role"] != "front_matter" and section_header_rel_id:
                first_header_forced = ensure_section_reference(sectpr, "headerReference", "first", section_header_rel_id)
                default_header_forced = ensure_section_reference(sectpr, "headerReference", "default", section_header_rel_id)
            if header_result["changed"] or footer_result["changed"]:
                changes.append(
                    {
                        "action": "sync_section_refs_from_template",
                        "sectionIndex": section["sectionIndex"],
                        "paragraphIndex": section["paragraphIndex"],
                        "sectionRole": section["role"],
                        "headerRemoved": header_result["removed"],
                        "headerEnsured": header_result["ensured"],
                        "footerRemoved": footer_result["removed"],
                        "footerEnsured": footer_result["ensured"],
                    }
                )
            if first_header_forced or default_header_forced:
                changes.append(
                    {
                        "action": "force_body_header_variants",
                        "sectionIndex": section["sectionIndex"],
                        "paragraphIndex": section["paragraphIndex"],
                        "sectionRole": section["role"],
                        "forcedFirstHeader": first_header_forced,
                        "forcedDefaultHeader": default_header_forced,
                        "assignedHeader": section_header_rel_id,
                    }
                )
        return changes
    section_header_rel_id = header_targets.get("word/header2.xml") or header_targets.get("word/header1.xml")
    default_footer_rel_id = footer_targets.get("word/footer1.xml") or footer_targets.get("word/footer2.xml")
    changes: List[Dict[str, Any]] = []
    if not section_header_rel_id and not default_footer_rel_id:
        return changes
    for section in classify_section_roles(document_root, protected_prefix_end, body_start_paragraph):
        if section["role"] == "front_matter":
            continue
        sectpr = section["sectPr"]
        first_header_removed = remove_section_reference(sectpr, "headerReference", "first")
        first_footer_removed = remove_section_reference(sectpr, "footerReference", "first")
        header_changed = (
            ensure_section_reference(sectpr, "headerReference", "default", section_header_rel_id)
            if section_header_rel_id
            else False
        )
        footer_changed = (
            ensure_section_reference(sectpr, "footerReference", "default", default_footer_rel_id)
            if default_footer_rel_id
            else False
        )
        if first_header_removed or first_footer_removed or header_changed or footer_changed:
            changes.append(
                {
                    "action": "ensure_body_section_refs",
                    "sectionIndex": section["sectionIndex"],
                    "paragraphIndex": section["paragraphIndex"],
                    "sectionRole": section["role"],
                    "headerChanged": header_changed,
                    "footerChanged": footer_changed,
                    "removedFirstHeader": first_header_removed,
                    "removedFirstFooter": first_footer_removed,
                    "assignedHeader": section_header_rel_id,
                    "assignedFooter": default_footer_rel_id,
                }
            )
    return changes


def normalize_tables(document_root: ET.Element, protected_prefix_end: int, rules: Dict[str, Any]) -> List[Dict[str, Any]]:
    body = document_root.find("w:body", NS)
    if body is None:
        return []
    table_rules = normalize_table_rules(rules)
    changes: List[Dict[str, Any]] = []
    paragraph_index = 0
    table_index = 0
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag == qn("p"):
            paragraph_index += 1
            continue
        if child.tag != qn("tbl"):
            continue
        if paragraph_index < protected_prefix_end:
            table_index += 1
            continue
        tblpr = ensure_child(child, "tblPr")
        tbl_style = tblpr.find("w:tblStyle", NS)
        if table_rules.get("strip_table_style") and tbl_style is not None:
            tblpr.remove(tbl_style)
        jc = ensure_child(tblpr, "jc")
        jc.set(qn("val"), str(table_rules.get("table_alignment") or "center"))
        if table_rules.get("border_style") == "research_three_line":
            top_rule = table_rules.get("top_rule") or {}
            bottom_rule = table_rules.get("bottom_rule") or {}
            set_table_three_line_borders(
                tblpr,
                str(top_rule.get("size") or 12),
                str(bottom_rule.get("size") or 12),
            )
        rows = child.findall("w:tr", NS)
        header_row_count = detect_header_row_count(rows, table_rules)
        group_rule_rows = detect_group_rule_rows(rows, header_row_count, table_rules)
        cell_count = 0
        cell_paragraph_count = 0
        for row_index, row in enumerate(rows):
            trpr = row.find("w:trPr", NS)
            if trpr is None:
                trpr = ET.Element(qn("trPr"))
                row.insert(0, trpr)
            set_row_cant_split(trpr)
            cells = row_cells(row)
            for cell in cells:
                cell_count += 1
                tcpr = ensure_child(cell, "tcPr")
                clear_cell_borders_and_shading(tcpr)
                # Explicitly suppress vertical borders per-cell; relying only on
                # tblBorders is not stable enough across Word renderers.
                for edge in ("left", "right", "insideV"):
                    set_cell_no_border(tcpr, edge)
                valign = ensure_child(tcpr, "vAlign")
                valign.set(qn("val"), str(table_rules.get("cell_vertical_alignment") or "center"))
                if (
                    table_rules.get("border_style") == "research_three_line"
                    and row_index == 0
                    and (table_rules.get("top_rule") or {}).get("enabled")
                    and (table_rules.get("top_rule") or {}).get("reinforce_on_first_row_cells")
                ):
                    set_cell_top_border(tcpr, str((table_rules.get("top_rule") or {}).get("size") or 12))
                if (
                    table_rules.get("border_style") == "research_three_line"
                    and row_index == header_row_count - 1
                    and (table_rules.get("header_rule") or {}).get("enabled")
                ):
                    set_cell_bottom_border(tcpr, str((table_rules.get("header_rule") or {}).get("size") or 8))
                if (
                    table_rules.get("border_style") == "research_three_line"
                    and row_index in group_rule_rows
                    and (table_rules.get("group_header_rule") or {}).get("enabled")
                ):
                    set_cell_bottom_border(tcpr, str((table_rules.get("group_header_rule") or {}).get("size") or 8))
                if (
                    table_rules.get("border_style") == "research_three_line"
                    and row_index == len(rows) - 1
                    and (table_rules.get("bottom_rule") or {}).get("enabled")
                    and (table_rules.get("bottom_rule") or {}).get("reinforce_on_last_row_cells")
                ):
                    set_cell_bottom_border(tcpr, str((table_rules.get("bottom_rule") or {}).get("size") or 12))
                for paragraph in cell.findall("w:p", NS):
                    cell_paragraph_count += 1
                    ppr = paragraph.find("w:pPr", NS)
                    if ppr is None:
                        ppr = ET.Element(qn("pPr"))
                        paragraph.insert(0, ppr)
                    paragraph_jc = ensure_child(ppr, "jc")
                    paragraph_jc.set(qn("val"), str(table_rules.get("cell_horizontal_alignment") or "center"))
        caption_position = table_caption_position(table_rules)
        if idx > 0 and idx + 1 < len(children):
            prev_node = children[idx - 1]
            next_node = children[idx + 1]
            if prev_node.tag == qn("p") and next_node.tag == qn("p"):
                prev_text = (element_text(prev_node) or "").strip()
                next_text = (element_text(next_node) or "").strip()
                if caption_position == "above_table" and TABLE_CAPTION_RE.match(next_text):
                    body.remove(next_node)
                    insert_at = list(body).index(child)
                    body.insert(insert_at, next_node)
                    changes.append(
                        {
                            "action": "move_table_caption_before_table",
                            "tableIndex": table_index,
                            "captionParagraphIndex": paragraph_index + 1,
                            "captionText": next_text[:120],
                        }
                    )
                    children = list(body)
                elif caption_position == "below_table" and TABLE_CAPTION_RE.match(prev_text):
                    body.remove(prev_node)
                    insert_at = list(body).index(child) + 1
                    body.insert(insert_at, prev_node)
                    changes.append(
                        {
                            "action": "move_table_caption_after_table",
                            "tableIndex": table_index,
                            "captionParagraphIndex": paragraph_index - 1,
                            "captionText": prev_text[:120],
                        }
                    )
                    children = list(body)
        changes.append(
            {
                "action": "normalize_table_layout",
                "tableIndex": table_index,
                "paragraphCursor": paragraph_index,
                "tablePolicy": table_rules.get("border_style"),
                "headerRowCount": header_row_count,
                "groupRuleRows": group_rule_rows,
                "rowCount": len(rows),
                "cellCount": cell_count,
                "cellParagraphCount": cell_paragraph_count,
            }
        )
        table_index += 1
    return changes


def normalize_front_abstract_block(
    document_root: ET.Element,
    protected_prefix_end: int,
    body_start_paragraph: Optional[int],
    style_specs: Dict[str, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    body = document_root.find("w:body", NS)
    if body is None or body_start_paragraph is None or body_start_paragraph <= protected_prefix_end:
        return []
    paragraphs = paragraph_nodes(body)
    if body_start_paragraph > len(paragraphs):
        body_start_paragraph = len(paragraphs)
    candidate_indices = [
        index
        for index in range(protected_prefix_end, body_start_paragraph)
        if (element_text(paragraphs[index]) or "").strip()
    ]
    if not candidate_indices:
        return []

    toc_marker = next(
        (
            index
            for index in candidate_indices
            if normalize_front_matter_token(element_text(paragraphs[index])) in {"目录", "在Word中打开后更新目录"}
        ),
        None,
    )
    if toc_marker is not None:
        candidate_indices = [index for index in candidate_indices if index > toc_marker]
    if not candidate_indices:
        return []

    title_index: Optional[int] = candidate_indices[0]
    abstract_heading_index: Optional[int] = None
    abstract_paragraph_index: Optional[int] = None
    keywords_index: Optional[int] = None
    author_index: Optional[int] = None
    title_en_index: Optional[int] = None
    author_en_index: Optional[int] = None
    abstract_paragraph_en_index: Optional[int] = None
    keywords_en_index: Optional[int] = None

    for index in candidate_indices:
        token = normalize_front_matter_token(element_text(paragraphs[index]))
        if token in ABSTRACT_CN_HEADINGS:
            abstract_heading_index = index
            break

    if abstract_heading_index is not None:
        following = [index for index in candidate_indices if index > abstract_heading_index]
        if following:
            abstract_paragraph_index = following[0]
        if title_index == abstract_heading_index:
            title_index = None
    else:
        for index in candidate_indices:
            text = (element_text(paragraphs[index]) or "").strip()
            compact = normalize_front_matter_token(text)
            if text.startswith("摘要：") or text.startswith("摘要:") or compact.startswith("摘要:"):
                abstract_paragraph_index = index
                if title_index == abstract_paragraph_index:
                    title_index = None
                break

    if abstract_paragraph_index is None:
        return []

    for index in candidate_indices:
        if index <= abstract_paragraph_index:
            continue
        text = (element_text(paragraphs[index]) or "").strip()
        compact = normalize_front_matter_token(text)
        if text.startswith("关键词：") or text.startswith("关键词:") or compact.startswith("关键词:"):
            keywords_index = index
            break

    changes: List[Dict[str, Any]] = []

    if title_index is not None and title_index < body_start_paragraph:
        title_paragraph = paragraphs[title_index]
        apply_style_to_paragraph(title_paragraph, style_specs["zafu_title_cn"])
        apply_run_defaults(title_paragraph, style_specs["zafu_title_cn"], preserve_emphasis=False)
        if title_index + 1 < len(paragraphs):
            subtitle_candidate = paragraphs[title_index + 1]
            subtitle_text = (element_text(subtitle_candidate) or "").strip()
            if subtitle_text.startswith("——"):
                merged_title = f"{(element_text(title_paragraph) or '').strip()} {subtitle_text}"
                replace_paragraph_plain_text(title_paragraph, merged_title)
                apply_style_to_paragraph(title_paragraph, style_specs["zafu_title_cn"])
                apply_run_defaults(title_paragraph, style_specs["zafu_title_cn"], preserve_emphasis=False)
                body.remove(subtitle_candidate)
        ppr = title_paragraph.find("w:pPr", NS)
        if ppr is None:
            ppr = ET.Element(qn("pPr"))
            title_paragraph.insert(0, ppr)
        if ppr.find("w:pageBreakBefore", NS) is None:
            ET.SubElement(ppr, qn("pageBreakBefore"))
        changes.append(
            {
                "action": "normalize_front_title_block",
                "paragraphIndex": title_index,
                "text": (element_text(title_paragraph) or "").strip(),
            }
        )
        for index in candidate_indices:
            if title_index < index < abstract_paragraph_index:
                author_index = index
                break

    if author_index is not None and author_index < body_start_paragraph:
        normalize_front_author_paragraph(paragraphs[author_index], style_specs["zafu_body"])
        changes.append(
            {
                "action": "normalize_front_author_paragraph",
                "paragraphIndex": author_index,
                "language": "cn",
            }
        )

    abstract_paragraph = paragraphs[abstract_paragraph_index]
    abstract_text = (element_text(abstract_paragraph) or "").strip()
    abstract_text = re.sub(r"^摘\s*要[：:]", "摘要：", abstract_text.replace("\u3000", ""))
    if not (abstract_text.startswith("摘要：") or abstract_text.startswith("摘要:") or normalize_front_matter_token(abstract_text).startswith("摘要:")):
        abstract_text = f"摘要：{abstract_text}"
    replace_paragraph_plain_text(abstract_paragraph, abstract_text)
    special_format_applied = apply_label_content_format(
        abstract_paragraph,
        ["摘要：", "摘要:"],
        style_specs["zafu_abstract_label_cn"],
        style_specs["zafu_abstract_text_cn"],
    )
    if not special_format_applied:
        apply_style_to_paragraph(abstract_paragraph, style_specs["zafu_abstract_text_cn"])
        apply_run_defaults(abstract_paragraph, style_specs["zafu_abstract_text_cn"], preserve_emphasis=False)
    changes.append(
        {
            "action": "normalize_front_abstract_paragraph",
            "paragraphIndex": abstract_paragraph_index,
            "mergedFromHeading": abstract_heading_index is not None,
        }
    )

    if keywords_index is not None and keywords_index < len(paragraphs):
        keywords_paragraph = paragraphs[keywords_index]
        special_format_applied = apply_label_content_format(
            keywords_paragraph,
            ["关键词：", "关键词:"],
            style_specs["zafu_keywords_label_cn"],
            style_specs["zafu_keywords_cn"],
        )
        if not special_format_applied:
            apply_style_to_paragraph(keywords_paragraph, style_specs["zafu_keywords_cn"])
            apply_run_defaults(keywords_paragraph, style_specs["zafu_keywords_cn"], preserve_emphasis=False)
        changes.append(
            {
                "action": "normalize_front_keywords_paragraph",
                "paragraphIndex": keywords_index,
            }
        )

    if keywords_index is not None:
        remaining_indices = [index for index in candidate_indices if index > keywords_index]
        for index in remaining_indices:
            text = (element_text(paragraphs[index]) or "").strip()
            if not text:
                continue
            if title_en_index is None:
                title_en_index = index
                continue
            if author_en_index is None and not (text.startswith("Abstract") or text.startswith("ABSTRACT") or text.startswith("Keywords")):
                author_en_index = index
                continue
            if abstract_paragraph_en_index is None and (text.startswith("Abstract") or text.startswith("ABSTRACT")):
                abstract_paragraph_en_index = index
                continue
            if keywords_en_index is None and text.startswith("Keywords"):
                keywords_en_index = index
                break

    if title_en_index is not None:
        title_paragraph_en = paragraphs[title_en_index]
        apply_style_to_paragraph(title_paragraph_en, style_specs["zafu_title_en"])
        apply_run_defaults(title_paragraph_en, style_specs["zafu_title_en"], preserve_emphasis=False)
        clear_paragraph_heading_semantics(title_paragraph_en)
        changes.append(
            {
                "action": "normalize_front_title_block_en",
                "paragraphIndex": title_en_index,
                "text": (element_text(title_paragraph_en) or "").strip(),
            }
        )

    if author_en_index is not None:
        normalize_front_author_paragraph(paragraphs[author_en_index], style_specs["zafu_body"])
        changes.append(
            {
                "action": "normalize_front_author_paragraph",
                "paragraphIndex": author_en_index,
                "language": "en",
            }
        )

    if abstract_paragraph_en_index is not None:
        abstract_paragraph_en = paragraphs[abstract_paragraph_en_index]
        abstract_text_en = (element_text(abstract_paragraph_en) or "").strip()
        if abstract_text_en.startswith("Abstract ") and not abstract_text_en.startswith("Abstract:"):
            abstract_text_en = f"Abstract:{abstract_text_en[len('Abstract'):]}"
            replace_paragraph_plain_text(abstract_paragraph_en, abstract_text_en)
        special_format_applied = apply_label_content_format(
            abstract_paragraph_en,
            ["Abstract:", "ABSTRACT:"],
            style_specs["zafu_abstract_label_en"],
            style_specs["zafu_abstract_text_en"],
        )
        if not special_format_applied:
            apply_style_to_paragraph(abstract_paragraph_en, style_specs["zafu_abstract_text_en"])
            apply_run_defaults(abstract_paragraph_en, style_specs["zafu_abstract_text_en"], preserve_emphasis=False)
        clear_paragraph_heading_semantics(abstract_paragraph_en)
        changes.append(
            {
                "action": "normalize_front_abstract_paragraph_en",
                "paragraphIndex": abstract_paragraph_en_index,
            }
        )

    if keywords_en_index is not None:
        keywords_paragraph_en = paragraphs[keywords_en_index]
        keywords_text_en = (element_text(keywords_paragraph_en) or "").strip()
        if normalize_front_matter_token(keywords_text_en).startswith("Keywords") and not keywords_text_en.startswith("Keywords:"):
            keywords_text_en = f"Keywords:{keywords_text_en[len('Keywords'):]}"
            replace_paragraph_plain_text(keywords_paragraph_en, keywords_text_en)
        special_format_applied = apply_label_content_format(
            keywords_paragraph_en,
            ["Key words:", "Key Words:", "KEY WORDS:", "Keywords:", "KEYWORDS:"],
            style_specs["zafu_keywords_label_en"],
            style_specs["zafu_keywords_en"],
        )
        if not special_format_applied:
            apply_style_to_paragraph(keywords_paragraph_en, style_specs["zafu_keywords_en"])
            apply_run_defaults(keywords_paragraph_en, style_specs["zafu_keywords_en"], preserve_emphasis=False)
        clear_paragraph_heading_semantics(keywords_paragraph_en)
        changes.append(
            {
                "action": "normalize_front_keywords_paragraph_en",
                "paragraphIndex": keywords_en_index,
            }
        )

    if abstract_heading_index is not None and abstract_heading_index != abstract_paragraph_index:
        body.remove(paragraphs[abstract_heading_index])
        changes.append(
            {
                "action": "remove_front_abstract_heading_paragraph",
                "paragraphIndex": abstract_heading_index,
            }
        )

    return changes


def normalize_plain_text_equations(document_root: ET.Element, protected_prefix_end: int) -> List[Dict[str, Any]]:
    body = document_root.find("w:body", NS)
    if body is None:
        return []
    changes: List[Dict[str, Any]] = []
    paragraphs = paragraph_nodes(body)
    paragraph_index = 0
    while paragraph_index < len(paragraphs):
        paragraph = paragraphs[paragraph_index]
        if paragraph_index < protected_prefix_end:
            paragraph_index += 1
            continue
        if paragraph.find(".//m:oMath", NS) is not None or paragraph.find(".//m:oMathPara", NS) is not None:
            paragraph_index += 1
            continue
        text = element_text(paragraph)
        if not text or not text.strip():
            paragraph_index += 1
            continue
        standalone_block = _consume_standalone_dollar_block(paragraphs, paragraph_index)
        if standalone_block is not None:
            end_index, display_latex = standalone_block
            if rebuild_paragraph_with_display_math(paragraph, display_latex, None):
                for remove_index in range(end_index, paragraph_index, -1):
                    body.remove(paragraphs[remove_index])
                changes.append(
                    {
                        "action": "convert_standalone_dollar_math_block",
                        "paragraphIndex": paragraph_index,
                        "endParagraphIndex": end_index,
                    }
                )
                paragraphs = paragraph_nodes(body)
            paragraph_index += 1
            continue
        display_latex, equation_number = split_display_math_and_number(text)
        if display_latex:
            changed = rebuild_paragraph_with_display_math(paragraph, display_latex, equation_number)
            if changed:
                changes.append(
                    {
                        "action": "convert_plain_text_display_equation",
                        "paragraphIndex": paragraph_index,
                        "equationNumber": equation_number,
                    }
                )
            paragraph_index += 1
            continue
        if paragraph_looks_like_formula(text):
            standalone_latex = extract_standalone_formula_latex(text)
            if standalone_latex and rebuild_paragraph_with_display_math(paragraph, standalone_latex, None):
                changes.append(
                    {
                        "action": "convert_plain_text_formula_like_paragraph",
                        "paragraphIndex": paragraph_index,
                        "mode": "display_formula_line",
                    }
                )
                paragraph_index += 1
                continue
            if rebuild_paragraph_with_inline_math(paragraph, text):
                changes.append(
                    {
                        "action": "convert_plain_text_formula_like_paragraph",
                        "paragraphIndex": paragraph_index,
                        "mode": "inline_math_scan",
                    }
                )
                paragraph_index += 1
                continue
        if INLINE_MATH_RE.search(text):
            changed = rebuild_paragraph_with_inline_math(paragraph, text)
            if changed:
                changes.append(
                    {
                        "action": "convert_plain_text_inline_equation",
                        "paragraphIndex": paragraph_index,
                    }
                )
        paragraph_index += 1
    return changes


def normalize_caption_label(kind: str, label: str) -> str:
    normalized = label.replace("—", "-").replace("．", ".").replace(".", "-")
    return f"{kind}{normalized}"


def next_bookmark_id(document_root: ET.Element) -> int:
    ids = []
    for node in document_root.findall(".//w:bookmarkStart", NS):
        value = safe_int(node.attrib.get(qn("id")))
        if value is not None:
            ids.append(value)
    return (max(ids) + 1) if ids else 1


def build_text_run_like(source_run: ET.Element, text: str) -> ET.Element:
    run = ET.Element(qn("r"))
    source_rpr = source_run.find("w:rPr", NS)
    if source_rpr is not None:
        run.append(clone_xml(source_rpr))
    text_node = ET.SubElement(run, qn("t"))
    text_node.text = text
    ensure_text_node_preserve(text_node, text)
    return run


def build_ref_field(token: str, bookmark_name: str, source_run: ET.Element) -> ET.Element:
    field = ET.Element(qn("fldSimple"))
    field.set(qn("instr"), f" REF {bookmark_name} \\\\h ")
    run = ET.SubElement(field, qn("r"))
    source_rpr = source_run.find("w:rPr", NS)
    if source_rpr is not None:
        run.append(clone_xml(source_rpr))
    text_node = ET.SubElement(run, qn("t"))
    text_node.text = token
    ensure_text_node_preserve(text_node, token)
    return field


def insert_caption_bookmarks(
    body_paragraphs: List[ET.Element],
    protected_prefix_end: int,
    document_root: ET.Element,
) -> Tuple[Dict[str, str], List[Dict[str, Any]]]:
    mapping: Dict[str, str] = {}
    changes: List[Dict[str, Any]] = []
    existing_names = {
        node.attrib.get(qn("name"))
        for node in document_root.findall(".//w:bookmarkStart", NS)
        if node.attrib.get(qn("name"))
    }
    bookmark_id = next_bookmark_id(document_root)
    for paragraph_index, paragraph in enumerate(body_paragraphs):
        if paragraph_index < protected_prefix_end:
            continue
        text = element_text(paragraph).strip()
        match = CAPTION_REF_RE.match(text)
        if not match:
            continue
        label_key = normalize_caption_label(match.group("kind"), match.group("label"))
        bookmark_name = f"{'fig' if match.group('kind') == '图' else 'tbl'}_{match.group('label').replace('—', '_').replace('．', '_').replace('.', '_').replace('-', '_')}"
        mapping[label_key] = bookmark_name
        if bookmark_name in existing_names:
            continue
        bookmark_start = ET.Element(qn("bookmarkStart"))
        bookmark_start.set(qn("id"), str(bookmark_id))
        bookmark_start.set(qn("name"), bookmark_name)
        bookmark_end = ET.Element(qn("bookmarkEnd"))
        bookmark_end.set(qn("id"), str(bookmark_id))
        bookmark_id += 1
        insert_at = 1 if paragraph.find("w:pPr", NS) is not None else 0
        paragraph.insert(insert_at, bookmark_start)
        paragraph.append(bookmark_end)
        existing_names.add(bookmark_name)
        changes.append(
            {
                "action": "insert_caption_bookmark",
                "paragraphIndex": paragraph_index,
                "bookmarkName": bookmark_name,
                "labelKey": label_key,
            }
        )
    return mapping, changes


def replace_cross_reference_runs(
    body_paragraphs: List[ET.Element],
    protected_prefix_end: int,
    bookmark_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    changes: List[Dict[str, Any]] = []
    for paragraph_index, paragraph in enumerate(body_paragraphs):
        if paragraph_index < protected_prefix_end:
            continue
        text = element_text(paragraph).strip()
        if not text or CAPTION_REF_RE.match(text):
            continue
        if paragraph.find(".//w:fldSimple", NS) is not None:
            continue
        paragraph_changes = []
        for child in list(paragraph):
            if child.tag != qn("r"):
                continue
            run_text = element_text(child)
            if not run_text:
                continue
            matches = list(CROSS_REF_RE.finditer(run_text))
            if not matches:
                continue
            insert_nodes: List[ET.Element] = []
            cursor = 0
            replaced = False
            for match in matches:
                label_key = normalize_caption_label(match.group("kind"), match.group("label"))
                bookmark_name = bookmark_map.get(label_key)
                if not bookmark_name:
                    continue
                if match.start() > cursor:
                    insert_nodes.append(build_text_run_like(child, run_text[cursor:match.start()]))
                insert_nodes.append(build_ref_field(match.group(0), bookmark_name, child))
                cursor = match.end()
                replaced = True
                paragraph_changes.append({"reference": match.group(0), "bookmarkName": bookmark_name})
            if not replaced:
                continue
            if cursor < len(run_text):
                insert_nodes.append(build_text_run_like(child, run_text[cursor:]))
            insertion_index = list(paragraph).index(child)
            paragraph.remove(child)
            for offset, node in enumerate(insert_nodes):
                paragraph.insert(insertion_index + offset, node)
        if paragraph_changes:
            changes.append(
                {
                    "action": "replace_cross_reference_text",
                    "paragraphIndex": paragraph_index,
                    "replacements": paragraph_changes,
                }
            )
    return changes


def update_sections(document_root: ET.Element, rules: Dict[str, Any]) -> List[Dict[str, Any]]:
    page = rules.get("page") or {}
    changes = []
    for sect_info in section_nodes(document_root):
        sect = sect_info["sectPr"]
        pgmar = sect.find("w:pgMar", NS)
        if pgmar is None:
            pgmar = ET.SubElement(sect, qn("pgMar"))
        for field, rule_key in [
            ("top", "margin_top_cm"),
            ("bottom", "margin_bottom_cm"),
            ("left", "margin_left_cm"),
            ("right", "margin_right_cm"),
            ("header", "header_cm"),
            ("footer", "footer_cm"),
        ]:
            if page.get(rule_key) is not None:
                pgmar.set(qn(field), str(cm_to_twips(float(page[rule_key]))))
        pgsz = sect.find("w:pgSz", NS)
        if pgsz is None:
            pgsz = ET.SubElement(sect, qn("pgSz"))
        pgsz.set(qn("w"), "11906")
        pgsz.set(qn("h"), "16838")
        changes.append({"action": "normalize_section_geometry", "paragraphIndex": sect_info["paragraphIndex"]})
    return changes


def page_size_from_rules(fallback_page_rules: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    page_rules = fallback_page_rules or {}
    paper = str(page_rules.get("paper") or "").strip().upper()
    if paper == "A4":
        return {"widthTwips": 11906, "heightTwips": 16838, "orient": "portrait"}
    return {}


def template_section_info_for_role(template_sections: List[Dict[str, Any]], role: str) -> Dict[str, Any]:
    if not template_sections:
        return {}
    if role == "front_matter":
        return template_sections[0]
    if role == "toc_abstract":
        index = 1 if len(template_sections) > 1 else len(template_sections) - 1
        return template_sections[index]
    index = 2 if len(template_sections) > 2 else len(template_sections) - 1
    return template_sections[index]


def template_section_page_for_role(template_sections: List[Dict[str, Any]], role: str) -> Dict[str, Any]:
    return template_section_info_for_role(template_sections, role).get("page") or {}


def apply_section_geometry_from_template(
    document_root: ET.Element,
    template_sections: List[Dict[str, Any]],
    protected_prefix_end: int,
    body_start_paragraph: Optional[int] = None,
    fallback_page_rules: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    changes = []
    forced_size = page_size_from_rules(fallback_page_rules)
    current_sections = classify_section_roles(document_root, protected_prefix_end, body_start_paragraph)
    for sect_info in current_sections:
        target_page: Dict[str, Any] = {}
        target_mode = None
        template_section = template_section_info_for_role(template_sections, sect_info["role"])
        # Only the immutable front-matter prefix should inherit template margins.
        # TOC/abstract/body sections must follow the profile rules to avoid leaking
        # the front-template's narrow page geometry into the thesis content.
        if sect_info["role"] == "front_matter":
            target_page = template_section.get("page") or {}
        if target_page:
            target_mode = "template"
        elif fallback_page_rules:
            target_page = {
                "top": {"twips": cm_to_twips(float(fallback_page_rules["margin_top_cm"]))} if fallback_page_rules.get("margin_top_cm") is not None else None,
                "bottom": {"twips": cm_to_twips(float(fallback_page_rules["margin_bottom_cm"]))} if fallback_page_rules.get("margin_bottom_cm") is not None else None,
                "left": {"twips": cm_to_twips(float(fallback_page_rules["margin_left_cm"]))} if fallback_page_rules.get("margin_left_cm") is not None else None,
                "right": {"twips": cm_to_twips(float(fallback_page_rules["margin_right_cm"]))} if fallback_page_rules.get("margin_right_cm") is not None else None,
                "header": {"twips": cm_to_twips(float(fallback_page_rules["header_cm"]))} if fallback_page_rules.get("header_cm") is not None else None,
                "footer": {"twips": cm_to_twips(float(fallback_page_rules["footer_cm"]))} if fallback_page_rules.get("footer_cm") is not None else None,
            }
            target_mode = "rules_fallback"
        else:
            continue
        sect = sect_info["sectPr"]
        pgmar = sect.find("w:pgMar", NS)
        if pgmar is None:
            pgmar = ET.SubElement(sect, qn("pgMar"))
        for field in ["top", "bottom", "left", "right", "header", "footer", "gutter"]:
            target = target_page.get(field)
            if target and target.get("twips") is not None:
                pgmar.set(qn(field), str(target["twips"]))
        pgsz = sect.find("w:pgSz", NS)
        if pgsz is None:
            pgsz = ET.SubElement(sect, qn("pgSz"))
        size = target_page.get("size") or {}
        if forced_size:
            size = {**size, **forced_size}
        if size.get("widthTwips") is not None:
            pgsz.set(qn("w"), str(size["widthTwips"]))
        if size.get("heightTwips") is not None:
            pgsz.set(qn("h"), str(size["heightTwips"]))
        orient = size.get("orient")
        if orient:
            pgsz.set(qn("orient"), orient)
        title_pg_changed = False
        section_type_changed = False
        columns_changed = False
        if template_section:
            title_pg_changed = set_section_title_page(sect, bool(template_section.get("titlePg")))
            section_type_changed = set_section_type_value(sect, template_section.get("sectionType"))
            columns = template_section.get("columns") or {}
            columns_changed = set_section_columns(
                sect,
                safe_int(columns.get("count")),
                safe_int(columns.get("space")),
            )
        changes.append(
            {
                "action": "apply_section_geometry",
                "mode": target_mode,
                "sectionIndex": sect_info["sectionIndex"],
                "paragraphIndex": sect_info["paragraphIndex"],
                "sectionRole": sect_info["role"],
                "titlePgChanged": title_pg_changed,
                "sectionTypeChanged": section_type_changed,
                "columnsChanged": columns_changed,
                "targetCm": {field: (target_page.get(field) or {}).get("cm") for field in ["top", "bottom", "left", "right", "header", "footer"]},
            }
        )
    return changes


def replace_header_text(xml_bytes: bytes, text: str) -> bytes:
    root = ET.fromstring(xml_bytes)
    paragraph = root.find("w:p", NS)
    if paragraph is None:
        paragraph = ET.SubElement(root, qn("p"))
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    pstyle = ensure_child(ppr, "pStyle")
    pstyle.set(qn("val"), "6")
    jc = ensure_child(ppr, "jc")
    jc.set(qn("val"), "center")
    pbdr = ensure_child(ppr, "pBdr")
    bottom = None
    for child in list(pbdr):
        if child.tag == qn("bottom"):
            bottom = child
            break
    if bottom is None:
        bottom = ET.SubElement(pbdr, qn("bottom"))
    bottom.set(qn("val"), "single")
    bottom.set(qn("color"), "auto")
    bottom.set(qn("sz"), "6")
    bottom.set(qn("space"), "1")

    for child in list(paragraph):
        if child.tag != qn("pPr"):
            paragraph.remove(child)

    run = ET.SubElement(paragraph, qn("r"))
    rpr = ET.SubElement(run, qn("rPr"))
    rfonts = ET.SubElement(rpr, qn("rFonts"))
    for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(slot), "宋体")
    size = ET.SubElement(rpr, qn("sz"))
    size.set(qn("val"), str(pt_to_half_points(9)))
    size_cs = ET.SubElement(rpr, qn("szCs"))
    size_cs.set(qn("val"), str(pt_to_half_points(9)))
    t = ET.SubElement(run, qn("t"))
    t.text = text
    ensure_text_node_preserve(t, text)

    for extra_paragraph in root.findall("w:p", NS)[1:]:
        for child in list(extra_paragraph):
            if child.tag != qn("pPr"):
                extra_paragraph.remove(child)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_update_fields_setting(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    node = root.find("w:updateFields", NS)
    if node is None:
        node = ET.SubElement(root, qn("updateFields"))
    node.set(qn("val"), "true")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def merge_missing_styles(base_styles_root: ET.Element, source_styles_root: ET.Element) -> int:
    existing_ids = {node.attrib.get(qn("styleId")) for node in base_styles_root.findall("w:style", NS)}
    added = 0
    for node in source_styles_root.findall("w:style", NS):
        style_id = node.attrib.get(qn("styleId"))
        if not style_id or style_id in existing_ids:
            continue
        base_styles_root.append(clone_xml(node))
        existing_ids.add(style_id)
        added += 1
    return added


def iter_relationship_attrs(node: ET.Element) -> Iterable[Tuple[ET.Element, str, str]]:
    for element in node.iter():
        for attr_name, value in list(element.attrib.items()):
            if value and attr_name.startswith(f"{{{NS['r']}}}"):
                yield element, attr_name, value


def part_rels_path(part_path: str) -> str:
    part = PurePosixPath(part_path)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def relative_target_for_part(base_part: str, target_part: str) -> str:
    base_parent = PurePosixPath(base_part).parent
    target = PurePosixPath(target_part)
    base_parts = list(base_parent.parts)
    target_parts = list(target.parts)
    common = 0
    while common < len(base_parts) and common < len(target_parts) and base_parts[common] == target_parts[common]:
        common += 1
    rel_parts = [".."] * (len(base_parts) - common) + target_parts[common:]
    return "/".join(rel_parts) if rel_parts else PurePosixPath(target_part).name


def next_relationship_id(existing_ids: Set[str]) -> str:
    index = 1
    while f"rId{index}" in existing_ids:
        index += 1
    new_id = f"rId{index}"
    existing_ids.add(new_id)
    return new_id


def unique_part_path(preferred_path: str, source_bytes: bytes, base_files: Dict[str, bytes]) -> str:
    if preferred_path not in base_files or base_files[preferred_path] == source_bytes:
        return preferred_path
    part = PurePosixPath(preferred_path)
    counter = 1
    while True:
        candidate = str(part.parent / f"{part.stem}_src{counter}{part.suffix}")
        if candidate not in base_files or base_files[candidate] == source_bytes:
            return candidate
        counter += 1


def find_content_type_root(files: Dict[str, bytes]) -> ET.Element:
    if "[Content_Types].xml" in files:
        return ET.fromstring(files["[Content_Types].xml"])
    return ET.Element(f"{{{CONTENT_TYPES_NS}}}Types")


def ensure_content_type_for_part(
    source_part: str,
    dest_part: str,
    source_content_types_root: ET.Element,
    base_content_types_root: ET.Element,
) -> bool:
    changed = False
    override_tag = f"{{{CONTENT_TYPES_NS}}}Override"
    default_tag = f"{{{CONTENT_TYPES_NS}}}Default"
    source_override = None
    for node in source_content_types_root.findall(override_tag):
        if node.attrib.get("PartName") == f"/{source_part}":
            source_override = node
            break
    if source_override is not None:
        exists = any(node.attrib.get("PartName") == f"/{dest_part}" for node in base_content_types_root.findall(override_tag))
        if not exists:
            new_override = ET.SubElement(base_content_types_root, override_tag)
            new_override.set("PartName", f"/{dest_part}")
            new_override.set("ContentType", source_override.attrib.get("ContentType", ""))
            changed = True
    else:
        extension = PurePosixPath(dest_part).suffix.lstrip(".")
        if extension:
            base_has_default = any(node.attrib.get("Extension") == extension for node in base_content_types_root.findall(default_tag))
            if not base_has_default:
                source_default = next(
                    (node for node in source_content_types_root.findall(default_tag) if node.attrib.get("Extension") == extension),
                    None,
                )
                if source_default is not None:
                    new_default = ET.SubElement(base_content_types_root, default_tag)
                    new_default.set("Extension", source_default.attrib.get("Extension", ""))
                    new_default.set("ContentType", source_default.attrib.get("ContentType", ""))
                    changed = True
    return changed


def copy_part_recursive(
    source_part: str,
    source_files: Dict[str, bytes],
    base_files: Dict[str, bytes],
    source_content_types_root: ET.Element,
    base_content_types_root: ET.Element,
    part_map: Dict[str, str],
    preferred_dest_part: Optional[str] = None,
) -> str:
    if source_part in part_map:
        return part_map[source_part]
    if source_part not in source_files:
        part_map[source_part] = preferred_dest_part or source_part
        return part_map[source_part]

    source_bytes = source_files[source_part]
    dest_part = unique_part_path(preferred_dest_part or source_part, source_bytes, base_files)
    part_map[source_part] = dest_part
    base_files[dest_part] = source_bytes
    ensure_content_type_for_part(source_part, dest_part, source_content_types_root, base_content_types_root)

    source_rels_path = part_rels_path(source_part)
    if source_rels_path not in source_files:
        return dest_part

    rels_root = ET.fromstring(source_files[source_rels_path])
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.attrib.get("TargetMode") == "External":
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        # Relationship targets are resolved relative to the source part, not
        # relative to the .rels file path.
        child_source_part = normalize_rel_target_word(source_part, target)
        child_dest_part = copy_part_recursive(
            child_source_part,
            source_files,
            base_files,
            source_content_types_root,
            base_content_types_root,
            part_map,
        )
        rel.set("Target", relative_target_for_part(dest_part, child_dest_part))

    base_files[part_rels_path(dest_part)] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    return dest_part


def merge_document_relationship_dependencies(
    nodes: List[ET.Element],
    source_files: Dict[str, bytes],
    base_files: Dict[str, bytes],
) -> Dict[str, Any]:
    source_rels_path = "word/_rels/document.xml.rels"
    source_document_part = "word/document.xml"
    if source_rels_path not in source_files:
        return {"mergedRelationships": 0, "copiedParts": 0}

    source_rels_root = ET.fromstring(source_files[source_rels_path])
    if source_rels_path in base_files:
        base_rels_root = ET.fromstring(base_files[source_rels_path])
    else:
        base_rels_root = ET.Element(f"{{{REL_NS}}}Relationships")

    source_content_types_root = find_content_type_root(source_files)
    base_content_types_root = find_content_type_root(base_files)
    existing_ids = {rel.attrib.get("Id") for rel in base_rels_root.findall(f"{{{REL_NS}}}Relationship") if rel.attrib.get("Id")}
    source_rel_by_id = {
        rel.attrib.get("Id"): rel
        for rel in source_rels_root.findall(f"{{{REL_NS}}}Relationship")
        if rel.attrib.get("Id")
    }
    part_map: Dict[str, str] = {}
    rel_id_map: Dict[str, str] = {}
    changed = 0
    copied_part_count = 0

    def find_existing_relationship(rel_type: str, target_mode: str, target_value: str, internal: bool) -> Optional[str]:
        for rel in base_rels_root.findall(f"{{{REL_NS}}}Relationship"):
            if rel.attrib.get("Type") != rel_type:
                continue
            if rel.attrib.get("TargetMode", "") != target_mode:
                continue
            base_target = rel.attrib.get("Target", "")
            if internal:
                normalized_base_target = normalize_rel_target_word(source_document_part, base_target)
                if normalized_base_target != target_value:
                    continue
                # Relationship reuse is only safe when the destination part
                # already contains the exact same payload. Two DOCX packages
                # commonly use identical media filenames such as image1.png,
                # but the bytes may differ. Reusing the existing relationship
                # by name alone would bind source body nodes to the template
                # cover asset (or vice versa).
                source_bytes = source_files.get(target_value)
                base_bytes = base_files.get(normalized_base_target)
                if source_bytes is not None and base_bytes is not None and source_bytes != base_bytes:
                    continue
                return rel.attrib.get("Id")
            elif base_target == target_value:
                return rel.attrib.get("Id")
        return None

    seen_source_rel_ids: Set[str] = set()
    attr_records: List[Tuple[ET.Element, str, str]] = []
    for node in nodes:
        for element, attr_name, rel_id in iter_relationship_attrs(node):
            attr_records.append((element, attr_name, rel_id))
            seen_source_rel_ids.add(rel_id)

    for source_rel_id in seen_source_rel_ids:
        source_rel = source_rel_by_id.get(source_rel_id)
        if source_rel is None:
            continue
        rel_type = source_rel.attrib.get("Type", "")
        target_mode = source_rel.attrib.get("TargetMode", "")
        target = source_rel.attrib.get("Target", "")
        internal = target_mode != "External"
        normalized_target = normalize_rel_target_word(source_document_part, target) if internal else target
        existing_rel_id = find_existing_relationship(rel_type, target_mode, normalized_target, internal)
        if existing_rel_id:
            rel_id_map[source_rel_id] = existing_rel_id
            continue

        dest_target = target
        if internal and normalized_target:
            prior_parts = set(part_map.values())
            dest_part = copy_part_recursive(
                normalized_target,
                source_files,
                base_files,
                source_content_types_root,
                base_content_types_root,
                part_map,
            )
            copied_part_count += len(set(part_map.values()) - prior_parts)
            dest_target = relative_target_for_part("word/document.xml", dest_part)

        new_rel = ET.SubElement(base_rels_root, f"{{{REL_NS}}}Relationship")
        new_rel_id = next_relationship_id(existing_ids)
        new_rel.set("Id", new_rel_id)
        new_rel.set("Type", rel_type)
        new_rel.set("Target", dest_target)
        if target_mode:
            new_rel.set("TargetMode", target_mode)
        rel_id_map[source_rel_id] = new_rel_id
        changed += 1

    for element, attr_name, source_rel_id in attr_records:
        mapped = rel_id_map.get(source_rel_id)
        if mapped and element.attrib.get(attr_name) != mapped:
            element.set(attr_name, mapped)

    base_files[source_rels_path] = ET.tostring(base_rels_root, encoding="utf-8", xml_declaration=True)
    base_files["[Content_Types].xml"] = ET.tostring(base_content_types_root, encoding="utf-8", xml_declaration=True)
    return {"mergedRelationships": changed, "copiedParts": copied_part_count}


def ensure_source_core_relationships(
    source_files: Dict[str, bytes],
    base_files: Dict[str, bytes],
) -> Dict[str, Any]:
    source_rels_path = "word/_rels/document.xml.rels"
    source_document_part = "word/document.xml"
    if source_rels_path not in source_files:
        return {"addedCoreRelationships": 0, "copiedCoreParts": 0}

    source_rels_root = ET.fromstring(source_files[source_rels_path])
    if source_rels_path in base_files:
        base_rels_root = ET.fromstring(base_files[source_rels_path])
    else:
        base_rels_root = ET.Element(f"{{{REL_NS}}}Relationships")

    source_content_types_root = find_content_type_root(source_files)
    base_content_types_root = find_content_type_root(base_files)
    existing_ids = {rel.attrib.get("Id") for rel in base_rels_root.findall(f"{{{REL_NS}}}Relationship") if rel.attrib.get("Id")}
    existing_types = {rel.attrib.get("Type"): rel for rel in base_rels_root.findall(f"{{{REL_NS}}}Relationship") if rel.attrib.get("Type")}
    part_map: Dict[str, str] = {}
    added = 0
    copied_part_count = 0

    for source_rel in source_rels_root.findall(f"{{{REL_NS}}}Relationship"):
        rel_type = source_rel.attrib.get("Type", "")
        if rel_type not in SOURCE_CORE_REL_TYPES or rel_type in existing_types:
            continue
        target = source_rel.attrib.get("Target")
        if not target or source_rel.attrib.get("TargetMode") == "External":
            continue
        source_part = normalize_rel_target_word(source_document_part, target)
        prior_parts = set(part_map.values())
        dest_part = copy_part_recursive(
            source_part,
            source_files,
            base_files,
            source_content_types_root,
            base_content_types_root,
            part_map,
        )
        copied_part_count += len(set(part_map.values()) - prior_parts)
        new_rel = ET.SubElement(base_rels_root, f"{{{REL_NS}}}Relationship")
        new_rel_id = next_relationship_id(existing_ids)
        new_rel.set("Id", new_rel_id)
        new_rel.set("Type", rel_type)
        new_rel.set("Target", relative_target_for_part("word/document.xml", dest_part))
        added += 1
        existing_types[rel_type] = new_rel

    base_files[source_rels_path] = ET.tostring(base_rels_root, encoding="utf-8", xml_declaration=True)
    base_files["[Content_Types].xml"] = ET.tostring(base_content_types_root, encoding="utf-8", xml_declaration=True)
    return {"addedCoreRelationships": added, "copiedCoreParts": copied_part_count}


def build_stub_part_xml(part_path: str) -> Optional[bytes]:
    if part_path == "word/comments.xml":
        root = ET.Element(qn("comments"))
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if part_path == "word/footnotes.xml":
        root = ET.Element(qn("footnotes"))
        separator = ET.SubElement(root, qn("footnote"))
        separator.set(qn("id"), "0")
        p0 = ET.SubElement(separator, qn("p"))
        r0 = ET.SubElement(p0, qn("r"))
        ET.SubElement(r0, qn("separator"))
        continuation = ET.SubElement(root, qn("footnote"))
        continuation.set(qn("id"), "1")
        p1 = ET.SubElement(continuation, qn("p"))
        r1 = ET.SubElement(p1, qn("r"))
        ET.SubElement(r1, qn("continuationSeparator"))
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if part_path == "word/endnotes.xml":
        root = ET.Element(qn("endnotes"))
        separator = ET.SubElement(root, qn("endnote"))
        separator.set(qn("id"), "0")
        p0 = ET.SubElement(separator, qn("p"))
        r0 = ET.SubElement(p0, qn("r"))
        ET.SubElement(r0, qn("separator"))
        continuation = ET.SubElement(root, qn("endnote"))
        continuation.set(qn("id"), "1")
        p1 = ET.SubElement(continuation, qn("p"))
        r1 = ET.SubElement(p1, qn("r"))
        ET.SubElement(r1, qn("continuationSeparator"))
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return None


def ensure_stub_part_content_type(part_path: str, content_types_root: ET.Element) -> bool:
    content_type_map = {
        "word/comments.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "word/footnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        "word/endnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    }
    content_type = content_type_map.get(part_path)
    if not content_type:
        return False
    override_tag = f"{{{CONTENT_TYPES_NS}}}Override"
    for node in content_types_root.findall(override_tag):
        if node.attrib.get("PartName") == f"/{part_path}":
            if node.attrib.get("ContentType") != content_type:
                node.set("ContentType", content_type)
                return True
            return False
    override = ET.SubElement(content_types_root, override_tag)
    override.set("PartName", f"/{part_path}")
    override.set("ContentType", content_type)
    return True


def ensure_internal_relationship_parts_exist(files: Dict[str, bytes]) -> Dict[str, Any]:
    rels_key = "word/_rels/document.xml.rels"
    rels_xml = files.get(rels_key)
    if not rels_xml:
        return {"createdStubParts": [], "removedDanglingRelationships": 0}
    rels_root = ET.fromstring(rels_xml)
    content_types_root = find_content_type_root(files)
    created_stub_parts: List[str] = []
    removed_dangling_relationships = 0
    for rel in list(rels_root.findall(f"{{{REL_NS}}}Relationship")):
        if rel.attrib.get("TargetMode") == "External":
            continue
        rel_type = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if not target:
            continue
        target_part = normalize_rel_target_word("word/document.xml", target)
        if target_part in files:
            continue
        if rel_type in SOURCE_CORE_REL_TYPES:
            stub = build_stub_part_xml(target_part)
            if stub is not None:
                files[target_part] = stub
                ensure_stub_part_content_type(target_part, content_types_root)
                created_stub_parts.append(target_part)
                continue
        rels_root.remove(rel)
        removed_dangling_relationships += 1
    files[rels_key] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    files["[Content_Types].xml"] = ET.tostring(content_types_root, encoding="utf-8", xml_declaration=True)
    return {
        "createdStubParts": created_stub_parts,
        "removedDanglingRelationships": removed_dangling_relationships,
    }


def resolve_reference_template_path(requested_template_docx: Optional[str], requested_template_audit: Optional[Dict[str, Any]]) -> Optional[Path]:
    if not requested_template_docx:
        return None
    requested_path = Path(requested_template_docx).resolve()
    section_count = len((requested_template_audit or {}).get("sections") or [])
    if section_count > 1:
        return requested_path

    return requested_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply deterministic OOXML formatting repairs to a thesis DOCX.")
    parser.add_argument("source_docx", help="Input DOCX")
    parser.add_argument("output_docx", help="Output DOCX")
    parser.add_argument("--plan-json", help="Optional repair plan JSON. If omitted, the plan is built automatically.")
    parser.add_argument("--template-docx", help="Optional template DOCX for baseline-aware planning")
    parser.add_argument("--rules-yaml", required=True, help="Rules YAML for target formatting")
    parser.add_argument("--style-map-yaml", help="Optional profile style-role mapping YAML")
    parser.add_argument("--front-matter-policy-yaml", help="Optional profile front-matter policy YAML")
    parser.add_argument("--validators-yaml", help="Optional profile validator selection YAML")
    parser.add_argument("--report-json", help="Write execution and diff report JSON to this path")
    args = parser.parse_args()

    rules = load_rules(args.rules_yaml)
    requested_template_audit = parse_document(args.template_docx, rules_path=args.rules_yaml) if args.template_docx else None
    reference_template_path = resolve_reference_template_path(args.template_docx, requested_template_audit)
    reference_template_docx = str(reference_template_path) if reference_template_path is not None else args.template_docx
    before_audit = parse_document(args.source_docx, template_docx=reference_template_docx, rules_path=args.rules_yaml)
    if args.plan_json:
        plan = json.loads(Path(args.plan_json).read_text(encoding="utf-8"))
    else:
        plan = build_repair_plan(
            before_audit,
            style_map_path=args.style_map_yaml,
            front_matter_policy_path=args.front_matter_policy_yaml,
        )
    editable_start_paragraph = safe_int(plan.get("editableStartParagraph")) or 0
    source_front_matter_drop_paragraph_count = safe_int(plan.get("sourceFrontMatterDropParagraphCount"))
    if source_front_matter_drop_paragraph_count is None:
        source_front_matter_drop_paragraph_count = editable_start_paragraph
    front_matter_policy = load_rules(args.front_matter_policy_yaml) if args.front_matter_policy_yaml else {}
    template_audit = parse_document(reference_template_docx, rules_path=args.rules_yaml) if reference_template_docx else None

    source_files = read_docx_files(args.source_docx)
    source_document_root = ET.fromstring(source_files["word/document.xml"])
    source_styles_root = ET.fromstring(source_files["word/styles.xml"])

    template_files: Optional[Dict[str, bytes]] = None
    template_document_root: Optional[ET.Element] = None
    template_section_patterns: List[Dict[str, Any]] = []
    template_prefix_paragraph_count = 0
    if args.template_docx:
        template_files = read_docx_files(args.template_docx)
        template_document_root = ET.fromstring(template_files["word/document.xml"])
        geometry_template_files = template_files
        if reference_template_docx and Path(reference_template_docx).resolve() != Path(args.template_docx).resolve():
            geometry_template_files = read_docx_files(reference_template_docx)
        template_section_patterns = extract_template_section_reference_patterns(geometry_template_files)
        template_prefix_paragraph_count = detect_template_prefix_count(requested_template_audit)

    force_template_front_matter = bool(front_matter_policy_value(front_matter_policy, "force_template_front_matter", False))
    use_template_front_matter = bool(
        template_files
        and template_document_root is not None
        and template_prefix_paragraph_count > 0
        and (force_template_front_matter or plan.get("useTemplateFrontMatter"))
    )

    execution_log: List[Dict[str, Any]] = []

    if use_template_front_matter:
        files = dict(template_files or {})
        document_root = template_document_root
        styles_root = ET.fromstring(files["word/styles.xml"])
        # Source body media must be imported through relationship-aware merging
        # below. Blindly copying word/media/* into the template package can
        # overwrite template cover assets when filenames collide (for example
        # image1.png in both packages), which corrupts the cover/integrity pages.
        # Preserve source numbering.xml so paragraph numbering refs remain valid
        if "word/numbering.xml" in source_files:
            files["word/numbering.xml"] = source_files["word/numbering.xml"]
    else:
        files = dict(source_files)
        document_root = source_document_root
        styles_root = ET.fromstring(files["word/styles.xml"])

    body = document_root.find("w:body", NS)
    if body is None:
        raise SystemExit("word/document.xml has no w:body")

    style_specs = style_spec_from_rules(rules)
    execution_log: List[Dict[str, Any]] = []
    source_prefix_paragraph_count = source_front_matter_drop_paragraph_count
    front_matter_log = None

    if use_template_front_matter and template_document_root is not None:
        front_matter_log, appended_source_nodes = compose_template_based_document(
            template_document_root=template_document_root,
            source_document_root=source_document_root,
            source_prefix_paragraph_count=source_prefix_paragraph_count,
            template_prefix_paragraph_count=template_prefix_paragraph_count,
        )
        execution_log.append({"action": "compose_template_based_document", **front_matter_log})
        merge_log = merge_document_relationship_dependencies(appended_source_nodes, source_files, files)
        if merge_log["mergedRelationships"] or merge_log["copiedParts"]:
            execution_log.append({"action": "merge_source_body_relationships_into_template", **merge_log})
        core_log = ensure_source_core_relationships(source_files, files)
        if core_log["addedCoreRelationships"] or core_log["copiedCoreParts"]:
            execution_log.append({"action": "ensure_source_core_relationships", **core_log})
        merged_style_count = merge_missing_styles(styles_root, source_styles_root)
        if merged_style_count:
            execution_log.append({"action": "merge_missing_source_styles", "count": merged_style_count})
        toc_insert_log = insert_toc_block_if_missing(
            document_root=document_root,
            source_document_root=source_document_root,
            source_prefix_paragraph_count=source_prefix_paragraph_count,
            current_prefix_paragraph_count=safe_int((front_matter_log or {}).get("finalPrefixParagraphCount")) or template_prefix_paragraph_count,
            inserted_boundary_paragraph=safe_int((front_matter_log or {}).get("insertedBoundaryParagraph")) or 0,
        )
        execution_log.append({"action": "insert_toc_block_if_missing", **toc_insert_log})
        inserted_toc_paragraphs = safe_int(toc_insert_log.get("paragraphsInserted")) or 0
        if front_matter_log and inserted_toc_paragraphs > 0:
            front_matter_log["finalPrefixParagraphCount"] = (
                safe_int(front_matter_log.get("finalPrefixParagraphCount")) or template_prefix_paragraph_count
            ) + inserted_toc_paragraphs
            front_matter_log["indexShift"] = (safe_int(front_matter_log.get("indexShift")) or 0) + inserted_toc_paragraphs
            front_matter_log["insertedTocParagraphs"] = inserted_toc_paragraphs
    elif template_document_root is not None and template_prefix_paragraph_count > 0:
        execution_log.append(
            {
                "action": "skip_legacy_front_matter_replace",
                "changed": False,
                "reason": "unsafe_front_matter_xml_replacement_disabled_to_preserve_template_relationship_integrity",
                "sourcePrefixParagraphCount": source_prefix_paragraph_count,
                "templatePrefixParagraphCount": template_prefix_paragraph_count,
            }
        )

    for spec in style_specs.values():
        ensure_style(styles_root, spec)
    execution_log.extend(ensure_builtin_toc_styles(styles_root, rules))

    index_shift = (front_matter_log or {}).get("indexShift", 0)
    protected_prefix_end = source_prefix_paragraph_count
    if use_template_front_matter and template_prefix_paragraph_count > 0:
        protected_prefix_end = template_prefix_paragraph_count
    if front_matter_log and safe_int(front_matter_log.get("finalPrefixParagraphCount")) is not None:
        protected_prefix_end = safe_int(front_matter_log.get("finalPrefixParagraphCount")) or protected_prefix_end

    geometry_protected_prefix_end = protected_prefix_end
    if use_template_front_matter and template_prefix_paragraph_count > 0:
        geometry_protected_prefix_end = template_prefix_paragraph_count

    body = document_root.find("w:body", NS)
    if body is None:
        raise SystemExit("word/document.xml has no w:body after template composition")
    body_paragraphs = paragraph_nodes(body)
    body_start_paragraph = detect_body_start_paragraph(
        body_paragraphs,
        plan,
        source_prefix_paragraph_count,
        index_shift,
        protected_prefix_end,
    )
    execution_log.extend(ensure_section_break_before_paragraph(document_root, body_start_paragraph, protected_prefix_end))

    if template_audit:
        execution_log.extend(
            apply_section_geometry_from_template(
                document_root,
                template_audit.get("sections") or [],
                geometry_protected_prefix_end,
                body_start_paragraph=body_start_paragraph,
                fallback_page_rules=(rules.get("page") or {}),
            )
        )
    else:
        execution_log.extend(update_sections(document_root, rules))

    execution_log.extend(
        ensure_body_header_references(
            document_root,
            files,
            protected_prefix_end,
            body_start_paragraph,
            template_section_patterns=template_section_patterns,
        )
    )
    execution_log.extend(apply_section_page_numbering(document_root, protected_prefix_end, body_start_paragraph))
    execution_log.extend(normalize_tables(document_root, protected_prefix_end, rules))

    body = document_root.find("w:body", NS)
    if body is None:
        raise SystemExit("word/document.xml has no w:body after repair planning")
    body_paragraphs = paragraph_nodes(body)
    for paragraph_index_str, style_id in (plan.get("styleMapping") or {}).items():
        index = safe_int(paragraph_index_str)
        if index is None:
            continue
        if index >= source_prefix_paragraph_count:
            index += index_shift
        if not (0 <= index < len(body_paragraphs)):
            continue
        if index < protected_prefix_end:
            execution_log.append({"action": "preserve_paragraph_region", "paragraphIndex": index, "reason": "immutable_prefix"})
            continue
        spec = style_specs.get(style_id)
        if not spec:
            continue
        paragraph = body_paragraphs[index]
        special_format_applied = False
        if style_id == "zafu_keywords_cn":
            special_format_applied = apply_label_content_format(
                paragraph,
                ["关键词：", "关键词:"],
                style_specs["zafu_keywords_label_cn"],
                spec,
            )
        elif style_id == "zafu_keywords_en":
            special_format_applied = apply_label_content_format(
                paragraph,
                ["Key words:", "Key Words:", "KEY WORDS:", "Keywords:", "KEYWORDS:"],
                style_specs["zafu_keywords_label_en"],
                spec,
            )
        elif style_id == "zafu_abstract_text_cn":
            special_format_applied = apply_label_content_format(
                paragraph,
                ["摘要：", "摘要:"],
                style_specs["zafu_abstract_label_cn"],
                spec,
            )
        elif style_id == "zafu_abstract_text_en":
            special_format_applied = apply_label_content_format(
                paragraph,
                ["Abstract:", "ABSTRACT:"],
                style_specs["zafu_abstract_label_en"],
                spec,
            )

        if special_format_applied:
            updated_runs = len(paragraph.findall(".//w:r", NS))
        else:
            apply_style_to_paragraph(paragraph, spec)
            updated_runs = apply_run_defaults(paragraph, spec, preserve_emphasis=(style_id == "zafu_body"))
        leading_whitespace_normalized = False
        if style_id in {"zafu_body", "zafu_caption"}:
            leading_whitespace_normalized = normalize_manual_leading_whitespace(paragraph)
        heading_spacing_changed = False
        if style_id in {"zafu_heading2", "zafu_heading3"}:
            heading_spacing_changed = normalize_heading_spacing(paragraph, style_id)
        execution_log.append(
            {
                "action": "apply_style_mapping",
                "paragraphIndex": index,
                "styleId": style_id,
                "runUpdates": updated_runs,
                "specialFormatApplied": special_format_applied,
                "leadingWhitespaceNormalized": leading_whitespace_normalized,
                "headingSpacingNormalized": heading_spacing_changed,
            }
        )

    for action in plan.get("numberingActions") or []:
        if action.get("action") != "replace_heading_prefix":
            continue
        index = safe_int(action.get("paragraphIndex"))
        if index is None:
            continue
        if index >= source_prefix_paragraph_count:
            index += index_shift
        if not (0 <= index < len(body_paragraphs)):
            continue
        if index < protected_prefix_end:
            execution_log.append({"action": "preserve_numbering_prefix", "paragraphIndex": index, "reason": "immutable_prefix"})
            continue
        changed = replace_prefix(body_paragraphs[index], action.get("oldPrefix", ""), action.get("newPrefix", ""))
        spacing_changed = False
        paragraph_style = None
        ppr = body_paragraphs[index].find("w:pPr", NS)
        if ppr is not None:
            pstyle = ppr.find("w:pStyle", NS)
            if pstyle is not None:
                paragraph_style = pstyle.attrib.get(qn("val"))
        if paragraph_style in {"zafu_heading2", "zafu_heading3"}:
            spacing_changed = normalize_heading_spacing(body_paragraphs[index], paragraph_style)
        execution_log.append(
            {
                "action": "replace_heading_prefix",
                "paragraphIndex": index,
                "oldPrefix": action.get("oldPrefix"),
                "newPrefix": action.get("newPrefix"),
                "changed": changed,
                "headingSpacingNormalized": spacing_changed,
            }
        )

    execution_log.extend(
        normalize_front_abstract_block(
            document_root,
            protected_prefix_end,
            body_start_paragraph,
            style_specs,
        )
    )
    body = document_root.find("w:body", NS)
    if body is None:
        raise SystemExit("word/document.xml has no w:body after front abstract normalization")
    body_paragraphs = paragraph_nodes(body)

    execution_log.extend(normalize_plain_text_equations(document_root, protected_prefix_end))
    execution_log.extend(normalize_caption_adjacent_figures(body, protected_prefix_end))
    execution_log.extend(normalize_references_section(document_root, protected_prefix_end, style_specs, rules))
    body_paragraphs = paragraph_nodes(body)

    bookmark_map, bookmark_changes = insert_caption_bookmarks(body_paragraphs, protected_prefix_end, document_root)
    execution_log.extend(bookmark_changes)
    execution_log.extend(replace_cross_reference_runs(body_paragraphs, protected_prefix_end, bookmark_map))
    execution_log.extend(normalize_decimal_heading_prefixes(body_paragraphs, protected_prefix_end))
    execution_log.extend(normalize_runs_by_existing_styles(body_paragraphs, style_specs, protected_prefix_end))
    execution_log.extend(normalize_body_style_fallback(body_paragraphs, style_specs, protected_prefix_end, body_start_paragraph))

    header_text = str(plan.get("headerText") or "").strip()
    if not header_text:
        header_text = infer_header_text_from_paragraphs(body_paragraphs, protected_prefix_end)
    target_headers = {"word/header2.xml"} if "word/header2.xml" in files else {
        name for name in files if name.startswith("word/header") and name.endswith(".xml")
    }
    if target_headers:
        for name in sorted(target_headers):
            files[name] = replace_header_text(files[name], header_text)
            execution_log.append({"action": "replace_header_text", "target": name, "text": header_text})

    if "word/settings.xml" in files:
        files["word/settings.xml"] = ensure_update_fields_setting(files["word/settings.xml"])
        execution_log.append({"action": "ensure_update_fields_on_open", "target": "word/settings.xml"})

    # Remove vanish (hidden) attributes from all runs to prevent content being invisible
    vanish_removed = 0
    for rpr in document_root.iter(f"{W}rPr"):
        vanish = rpr.find(f"{W}vanish")
        if vanish is not None:
            rpr.remove(vanish)
            vanish_removed += 1
    if vanish_removed > 0:
        execution_log.append({"action": "remove_vanish_attributes", "removed": vanish_removed})

    files["word/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
    files["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)

    # Fix incorrect _rels/ prefix in relationship target paths
    rels_key = "word/_rels/document.xml.rels"
    if rels_key in files:
        rels_root = ET.fromstring(files[rels_key])
        fixed_paths = 0
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
            target = rel.attrib.get("Target", "")
            if target.startswith("_rels/") and not target.startswith("_rels/.rels"):
                rel.attrib["Target"] = target[len("_rels/"):]
                fixed_paths += 1
        if fixed_paths > 0:
            files[rels_key] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
            execution_log.append({"action": "fix_rel_paths", "fixed": fixed_paths})
    rel_integrity_log = ensure_internal_relationship_parts_exist(files)
    if rel_integrity_log["createdStubParts"] or rel_integrity_log["removedDanglingRelationships"]:
        execution_log.append({"action": "ensure_internal_relationship_parts_exist", **rel_integrity_log})

    try:
        with zipfile.ZipFile(args.output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)
    except PermissionError as exc:
        raise SystemExit(
            f"Cannot write output DOCX '{args.output_docx}'. Close any app that is opening this file and retry."
        ) from exc

    after_audit = parse_document(args.output_docx, template_docx=reference_template_docx, rules_path=args.rules_yaml)
    report = {
        "generatedAt": after_audit["generatedAt"],
        "sourceDocx": str(Path(args.source_docx).resolve()),
        "outputDocx": str(Path(args.output_docx).resolve()),
        "plan": plan,
        "executionLog": execution_log,
        "executionSummary": summarize_execution_log(execution_log),
        "diffReport": diff_audits(before_audit, after_audit, plan=plan, execution_log=execution_log),
    }
    if args.report_json:
        write_json(report, args.report_json)


if __name__ == "__main__":
    main()


