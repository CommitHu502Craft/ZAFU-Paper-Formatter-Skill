from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from typing import Callable

from docx_ooxml import NS, qn


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def header_sections(document: ET.Element, files: dict[str, bytes]) -> list[dict]:
    relationships = ET.fromstring(files["word/_rels/document.xml.rels"])
    targets = {
        item.get("Id"): posixpath.normpath(posixpath.join("word", item.get("Target", ""))).lstrip("/")
        for item in relationships
        if item.get("Type", "").endswith("/header") and item.get("TargetMode") != "External"
    }
    body = document.find("w:body", NS)
    if body is None:
        return []
    inherited = {}
    content = []
    result = []
    previous_scope = None
    for child in body:
        content.append(child)
        section = child if child.tag == qn("sectPr") else child.find("w:pPr/w:sectPr", NS)
        if section is None:
            continue
        paragraphs = [node for node in content if node.tag == qn("p")]
        texts = ["".join(node.text or "" for node in paragraph.iter(qn("t"))) for paragraph in paragraphs]
        visible = ["".join(node.itertext()) for paragraph in paragraphs for node in paragraph.findall(".//w:t", NS)]
        has_toc = any(re.sub(r"\s+", "", text) == "目录" for text in texts) or any(
            re.search(r"\bTOC\b", node.text or "", re.IGNORECASE)
            for paragraph in content for node in paragraph.iter(qn("instrText"))
        ) or any(
            re.search(r"\bTOC\b", node.get(qn("instr"), ""), re.IGNORECASE)
            for paragraph in content for node in paragraph.iter(qn("fldSimple"))
        )
        has_heading = any(
            re.fullmatch(r"(?:zafu_heading|Heading|heading)[1-6]", node.get(qn("val"), ""))
            for paragraph, text in zip(paragraphs, texts)
            if re.sub(r"\s+", "", text).lower() not in {"目录", "摘要", "中文摘要", "abstract"}
            for node in paragraph.findall("w:pPr/w:pStyle", NS)
        )
        has_front = any("诚信承诺书" in text or "本科生毕业论文（设计）" in text for text in visible)
        has_abstract = any(re.match(r"^(?:摘\s*要|中文摘要|Abstract)(?:\s*[:：]|\s*$)", text.strip(), re.I) for text in texts)
        scope = "unknown"
        if has_toc:
            scope = "mixed" if has_heading else "toc"
        elif has_front:
            scope = "front_matter"
        elif has_heading or has_abstract or previous_scope == "body":
            scope = "body"
        for reference in section.findall("w:headerReference", NS):
            inherited[reference.get(qn("type"), "default")] = reference.get(f"{{{NS['r']}}}id")
        result.append({
            "index": len(result), "scope": scope, "sectPr": section,
            "headers": {variant: {"id": rel_id, "path": targets.get(rel_id)} for variant, rel_id in inherited.items()},
        })
        content = []
        previous_scope = scope
    return result


def apply_header_policy(document: ET.Element, files: dict[str, bytes], text: str, mode: str, replace_text: Callable[[bytes, str], bytes], *, create_missing: bool = False) -> list[dict]:
    if mode == "preserve":
        return []
    sections = header_sections(document, files)
    relationships = ET.fromstring(files["word/_rels/document.xml.rels"])
    content_types = ET.fromstring(files["[Content_Types].xml"])
    used_ids = {item.get("Id") for item in relationships}
    changes = []
    changed_variants = set()
    for section in sections:
        if create_missing and section["scope"] in {"body", "toc"} and "default" not in section["headers"]:
            section["headers"]["default"] = {"id": None, "path": None}
        for variant, header in section["headers"].items():
            path = header["path"]
            rel_id = header["id"]
            missing = not path or path not in files
            if missing and not (create_missing and (section["scope"] == "toc" or section["scope"] == "body" and text)):
                changes.append({"action": "header_manual_review", "sectionIndex": section["index"], "reason": "missing_header_part"})
                continue
            scope = section["scope"]
            source_bytes = files[path] if not missing else ET.tostring(ET.Element(qn("hdr")), encoding="utf-8")
            source = ET.fromstring(source_bytes)
            populated = bool("".join(source.itertext()).strip())
            should_change = scope == "toc" or (scope == "body" and (populated or create_missing) and bool(text))
            if scope in {"mixed", "unknown"}:
                changes.append({"action": "header_manual_review", "sectionIndex": section["index"], "reason": f"{scope}_section_preserved"})
            if should_change:
                number = 1
                while f"word/header_formatter_{number}.xml" in files or f"rIdFormatterHeader{number}" in used_ids:
                    number += 1
                destination = f"word/header_formatter_{number}.xml"
                rel_id = f"rIdFormatterHeader{number}"
                used_ids.add(rel_id)
                if scope == "toc":
                    blank = ET.Element(qn("hdr"))
                    ET.SubElement(blank, qn("p"))
                    files[destination] = ET.tostring(blank, encoding="utf-8", xml_declaration=True)
                else:
                    files[destination] = replace_text(source_bytes, text)
                ET.SubElement(relationships, f"{{{REL_NS}}}Relationship", {
                    "Id": rel_id, "Type": f"{NS['r']}/header", "Target": destination.removeprefix("word/"),
                })
                ET.SubElement(content_types, f"{{{CT_NS}}}Override", {
                    "PartName": f"/{destination}",
                    "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                })
                changed_variants.add(variant)
                changes.append({"action": "clear_toc_header" if scope == "toc" else "replace_header_text", "sectionIndex": section["index"], "variant": variant, "target": destination})
            if should_change or variant in changed_variants:
                reference = next((node for node in section["sectPr"].findall("w:headerReference", NS) if node.get(qn("type"), "default") == variant), None)
                if reference is None:
                    reference = ET.Element(qn("headerReference"), {qn("type"): variant})
                    section["sectPr"].insert(0, reference)
                reference.set(f"{{{NS['r']}}}id", rel_id)
    files["word/_rels/document.xml.rels"] = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
    files["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    return changes


def initialize_fresh_footers(document: ET.Element, files: dict[str, bytes], latin_font: str = "Times New Roman") -> list[dict]:
    """Give unambiguous generated non-cover sections independent PAGE fields."""
    relationships = ET.fromstring(files["word/_rels/document.xml.rels"])
    content_types = ET.fromstring(files["[Content_Types].xml"])
    used_ids = {item.get("Id") for item in relationships}
    changes = []
    for section in header_sections(document, files):
        if section["scope"] not in {"body", "toc"}:
            continue
        number = 1
        while f"word/footer_formatter_{number}.xml" in files or f"rIdFormatterFooter{number}" in used_ids:
            number += 1
        path = f"word/footer_formatter_{number}.xml"
        rel_id = f"rIdFormatterFooter{number}"
        used_ids.add(rel_id)
        root = ET.Element(qn("ftr"))
        paragraph = ET.SubElement(root, qn("p"))
        ppr = ET.SubElement(paragraph, qn("pPr"))
        ET.SubElement(ppr, qn("jc"), {qn("val"): "center"})
        field = ET.SubElement(paragraph, qn("fldSimple"), {qn("instr"): " PAGE "})
        run = ET.SubElement(field, qn("r"))
        rpr = ET.SubElement(run, qn("rPr"))
        ET.SubElement(rpr, qn("rFonts"), {qn("ascii"): latin_font, qn("hAnsi"): latin_font, qn("eastAsia"): "宋体"})
        ET.SubElement(rpr, qn("sz"), {qn("val"): "18"})
        ET.SubElement(run, qn("t")).text = "1"
        files[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        ET.SubElement(relationships, f"{{{REL_NS}}}Relationship", {
            "Id": rel_id, "Type": f"{NS['r']}/footer", "Target": path.removeprefix("word/"),
        })
        ET.SubElement(content_types, f"{{{CT_NS}}}Override", {
            "PartName": f"/{path}", "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        })
        sectpr = section["sectPr"]
        for node in list(sectpr.findall("w:footerReference", NS)) + list(sectpr.findall("w:titlePg", NS)):
            sectpr.remove(node)
        reference = ET.Element(qn("footerReference"), {qn("type"): "default", f"{{{NS['r']}}}id": rel_id})
        # References precede page geometry in sectPr.
        insert_at = len(sectpr.findall("w:headerReference", NS))
        sectpr.insert(insert_at, reference)
        changes.append({"action": "initialize_fresh_footer", "sectionIndex": section["index"], "target": path})
    files["word/_rels/document.xml.rels"] = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
    files["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    return changes


def check_header_policy(document: ET.Element, files: dict[str, bytes], text: str, mode: str) -> list[dict]:
    if mode == "preserve":
        return []
    issues = []
    for section in header_sections(document, files):
        scope = section["scope"]
        if scope in {"mixed", "unknown"} and section["headers"]:
            issues.append({"type": "header_scope_needs_review", "sectionIndex": section["index"], "scope": scope})
        for variant, header in section["headers"].items():
            path = header["path"]
            if not path or path not in files:
                continue
            root = ET.fromstring(files[path])
            actual = "".join(node.text or "" for node in root.iter(qn("t"))).strip()
            if scope == "toc" and (actual or root.find(".//w:pBdr", NS) is not None or root.find(".//w:drawing", NS) is not None):
                issues.append({"type": "toc_header_not_empty", "sectionIndex": section["index"], "variant": variant})
            elif scope == "body" and actual and text and actual != text:
                issues.append({"type": "header_text_mismatch", "sectionIndex": section["index"], "variant": variant, "expected": text, "actual": actual})
    return issues
