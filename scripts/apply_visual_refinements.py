#!/usr/bin/env python3
"""Apply a whitelisted visual refinement plan to a DOCX.

Reads visual_refinement_plan.json produced by an agent after reviewing the
rendered contact sheet / critical pages, and applies only a fixed whitelist
of layout-level actions. Body text is never rewritten; unknown actions and
unresolvable targets are reported and skipped, never fatal.

Plan format:
{
  "actions": [
    {
      "type": "set_keep_with_next",
      "target": {"paragraphId": "p_0012", "textPrefix": "第三章"},
      "reason": "一级标题落在页面末尾",
      "confidence": 0.92
    }
  ]
}

Targets:
  paragraph actions: {"paragraphId": "p_NNNN"} and/or {"paragraphIndex": N},
    optional "textPrefix" for verification (recommended; mismatch = skipped).
  table actions:     {"tableIndex": N}

Whitelisted action types:
  set_keep_with_next            paragraph keepNext
  set_keep_lines                paragraph keepLines
  set_page_break_before         paragraph pageBreakBefore
  clear_page_break_before       remove pageBreakBefore
  remove_empty_paragraph        delete paragraph only when it has no text/drawing
  remove_duplicate_page_break   drop redundant <w:br type="page"> runs beyond the first
  set_widow_control             paragraph widowControl
  set_spacing                   paragraph spacing beforePt/afterPt (0..72)
  center_paragraph              jc=center
  scale_image_to_width          shrink oversized inline image to maxWidthCm keeping ratio
  set_table_header_repeat       tblHeader on first N rows
  set_table_rows_no_split       cantSplit on non-header rows
  center_table                  tblPr jc=center
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"w": W_NS, "wp": WP_NS, "a": A_NS}

EMU_PER_CM = 360000

ALLOWED_ACTIONS = {
    "set_keep_with_next",
    "set_keep_lines",
    "set_page_break_before",
    "clear_page_break_before",
    "remove_empty_paragraph",
    "remove_duplicate_page_break",
    "set_widow_control",
    "set_spacing",
    "center_paragraph",
    "scale_image_to_width",
    "set_table_header_repeat",
    "set_table_rows_no_split",
    "center_table",
}


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def register_namespaces() -> None:
    for prefix, uri in {
        "w": W_NS,
        "wp": WP_NS,
        "a": A_NS,
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "v": "urn:schemas-microsoft-com:vml",
        "o": "urn:schemas-microsoft-com:office:office",
        "w10": "urn:schemas-microsoft-com:office:word",
    }.items():
        ET.register_namespace(prefix, uri)


def paragraph_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS))


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def set_flag(ppr: ET.Element, tag: str, value: bool) -> None:
    node = ppr.find(f"w:{tag}", NS)
    if value:
        if node is None:
            ET.SubElement(ppr, qn(tag))
    elif node is not None:
        ppr.remove(node)


class RefinementEngine:
    def __init__(self, document_root: ET.Element):
        self.root = document_root
        self.body = document_root.find("w:body", NS)
        if self.body is None:
            raise ValueError("word/document.xml has no <w:body>")
        self.paragraphs: List[ET.Element] = list(document_root.iter(qn("p")))
        self.tables: List[ET.Element] = list(document_root.iter(qn("tbl")))

    # -- target resolution ------------------------------------------------
    def resolve_paragraph(self, target: Dict[str, Any]) -> Optional[ET.Element]:
        index: Optional[int] = None
        used_block_id = False
        paragraph_id = target.get("paragraphId")
        if not isinstance(paragraph_id, str):
            paragraph_id = target.get("blockId")
            used_block_id = isinstance(paragraph_id, str)
        if isinstance(paragraph_id, str):
            match = re.search(r"(\d+)", paragraph_id)
            if match:
                index = int(match.group(1))
        if index is None and isinstance(target.get("paragraphIndex"), int):
            index = int(target["paragraphIndex"])
        if index is None or not (0 <= index < len(self.paragraphs)):
            return None
        node = self.paragraphs[index]
        prefix = target.get("textPrefix")
        if used_block_id and not (isinstance(prefix, str) and prefix.strip()):
            # blockId ordinals (block-NNNNN) are NOT paragraph indexes; without
            # a textPrefix to verify, refusing is safer than guessing. Use the
            # docxParagraphIndex field from thesis_ir.json semanticBlocks instead.
            return None
        if isinstance(prefix, str) and prefix.strip():
            actual = re.sub(r"\s+", "", paragraph_text(node))
            wanted = re.sub(r"\s+", "", prefix)
            if not actual.startswith(wanted[:24]):
                return None
        return node

    def resolve_table(self, target: Dict[str, Any]) -> Optional[ET.Element]:
        index = target.get("tableIndex")
        if isinstance(index, int) and 0 <= index < len(self.tables):
            return self.tables[index]
        return None

    def remove_paragraph(self, node: ET.Element) -> bool:
        for parent in self.root.iter():
            children = list(parent)
            if node in children:
                parent.remove(node)
                return True
        return False

    # -- actions ----------------------------------------------------------
    def apply(self, action: Dict[str, Any]) -> Dict[str, Any]:
        action_type = str(action.get("type") or "")
        target = action.get("target") or {}
        entry: Dict[str, Any] = {
            "type": action_type,
            "target": target,
            "reason": action.get("reason"),
            "status": "skipped",
            "detail": None,
        }
        if action_type not in ALLOWED_ACTIONS:
            entry["detail"] = "action_not_in_whitelist"
            return entry

        if action_type in {"set_table_header_repeat", "set_table_rows_no_split", "center_table"}:
            table = self.resolve_table(target)
            if table is None:
                entry["detail"] = "table_target_not_resolved"
                return entry
            handler = getattr(self, f"_do_{action_type}")
            return handler(table, action, entry)

        paragraph = self.resolve_paragraph(target)
        if paragraph is None:
            entry["detail"] = "paragraph_target_not_resolved_or_text_mismatch"
            return entry
        handler = getattr(self, f"_do_{action_type}")
        return handler(paragraph, action, entry)

    def _do_set_keep_with_next(self, paragraph, action, entry):
        set_flag(ensure_ppr(paragraph), "keepNext", True)
        entry["status"] = "applied"
        return entry

    def _do_set_keep_lines(self, paragraph, action, entry):
        set_flag(ensure_ppr(paragraph), "keepLines", True)
        entry["status"] = "applied"
        return entry

    def _do_set_page_break_before(self, paragraph, action, entry):
        set_flag(ensure_ppr(paragraph), "pageBreakBefore", True)
        entry["status"] = "applied"
        return entry

    def _do_clear_page_break_before(self, paragraph, action, entry):
        set_flag(ensure_ppr(paragraph), "pageBreakBefore", False)
        entry["status"] = "applied"
        return entry

    def _do_set_widow_control(self, paragraph, action, entry):
        ppr = ensure_ppr(paragraph)
        node = ppr.find("w:widowControl", NS)
        if node is None:
            node = ET.SubElement(ppr, qn("widowControl"))
        node.attrib.pop(qn("val"), None)
        entry["status"] = "applied"
        return entry

    def _do_center_paragraph(self, paragraph, action, entry):
        ppr = ensure_ppr(paragraph)
        jc = ppr.find("w:jc", NS)
        if jc is None:
            jc = ET.SubElement(ppr, qn("jc"))
        jc.set(qn("val"), "center")
        entry["status"] = "applied"
        return entry

    def _do_set_spacing(self, paragraph, action, entry):
        before = action.get("beforePt")
        after = action.get("afterPt")
        if before is None and after is None:
            entry["detail"] = "set_spacing_requires_beforePt_or_afterPt"
            return entry
        ppr = ensure_ppr(paragraph)
        spacing = ppr.find("w:spacing", NS)
        if spacing is None:
            spacing = ET.SubElement(ppr, qn("spacing"))
        for key, attr in ((before, "before"), (after, "after")):
            if key is None:
                continue
            points = max(0.0, min(float(key), 72.0))
            spacing.set(qn(attr), str(int(round(points * 20))))
        entry["status"] = "applied"
        return entry

    def _do_remove_empty_paragraph(self, paragraph, action, entry):
        if paragraph_text(paragraph).strip():
            entry["detail"] = "paragraph_not_empty"
            return entry
        if paragraph.find(".//w:drawing", NS) is not None or paragraph.find(".//w:pict", NS) is not None:
            entry["detail"] = "paragraph_contains_drawing"
            return entry
        ppr = paragraph.find("w:pPr", NS)
        if ppr is not None and ppr.find("w:sectPr", NS) is not None:
            entry["detail"] = "paragraph_carries_section_break"
            return entry
        if self.remove_paragraph(paragraph):
            entry["status"] = "applied"
        else:
            entry["detail"] = "paragraph_parent_not_found"
        return entry

    def _do_remove_duplicate_page_break(self, paragraph, action, entry):
        removed = 0
        seen = False
        for run in list(paragraph.findall("w:r", NS)):
            for br in list(run.findall("w:br", NS)):
                if br.get(qn("type")) != "page":
                    continue
                if seen:
                    run.remove(br)
                    removed += 1
                else:
                    seen = True
            if not list(run) or all(child.tag == qn("rPr") for child in run):
                if removed:
                    paragraph.remove(run)
        if removed:
            entry["status"] = "applied"
            entry["detail"] = f"removed_{removed}_duplicate_page_breaks"
        else:
            entry["detail"] = "no_duplicate_page_break_found"
        return entry

    def _do_scale_image_to_width(self, paragraph, action, entry):
        max_width_cm = float(action.get("maxWidthCm") or 14.6)
        max_emu = int(max_width_cm * EMU_PER_CM)
        scaled = 0
        for extent in paragraph.findall(".//wp:extent", NS):
            cx = int(extent.get("cx") or 0)
            cy = int(extent.get("cy") or 0)
            if cx > max_emu and cx > 0:
                ratio = max_emu / cx
                extent.set("cx", str(max_emu))
                extent.set("cy", str(int(cy * ratio)))
                scaled += 1
        # Also scale the drawing's inner a:ext to match, keeping Word happy.
        if scaled:
            for ext in paragraph.findall(".//a:xfrm/a:ext", NS):
                cx = int(ext.get("cx") or 0)
                cy = int(ext.get("cy") or 0)
                if cx > max_emu and cx > 0:
                    ratio = max_emu / cx
                    ext.set("cx", str(max_emu))
                    ext.set("cy", str(int(cy * ratio)))
            entry["status"] = "applied"
            entry["detail"] = f"scaled_{scaled}_images_to_{max_width_cm}cm"
        else:
            entry["detail"] = "no_oversized_inline_image_found"
        return entry

    def _do_set_table_header_repeat(self, table, action, entry):
        header_rows = max(1, min(int(action.get("headerRows") or 1), 3))
        rows = table.findall("w:tr", NS)
        if not rows:
            entry["detail"] = "table_has_no_rows"
            return entry
        for row in rows[:header_rows]:
            trpr = row.find("w:trPr", NS)
            if trpr is None:
                trpr = ET.Element(qn("trPr"))
                row.insert(0, trpr)
            if trpr.find("w:tblHeader", NS) is None:
                ET.SubElement(trpr, qn("tblHeader"))
        entry["status"] = "applied"
        return entry

    def _do_set_table_rows_no_split(self, table, action, entry):
        rows = table.findall("w:tr", NS)
        changed = 0
        for row in rows:
            trpr = row.find("w:trPr", NS)
            if trpr is None:
                trpr = ET.Element(qn("trPr"))
                row.insert(0, trpr)
            if trpr.find("w:cantSplit", NS) is None:
                ET.SubElement(trpr, qn("cantSplit"))
                changed += 1
        entry["status"] = "applied" if changed else "skipped"
        entry["detail"] = f"cantSplit_added_to_{changed}_rows"
        return entry

    def _do_center_table(self, table, action, entry):
        tblpr = table.find("w:tblPr", NS)
        if tblpr is None:
            tblpr = ET.Element(qn("tblPr"))
            table.insert(0, tblpr)
        jc = tblpr.find("w:jc", NS)
        if jc is None:
            jc = ET.SubElement(tblpr, qn("jc"))
        jc.set(qn("val"), "center")
        entry["status"] = "applied"
        return entry


def apply_plan(input_docx: Path, output_docx: Path, plan: Dict[str, Any]) -> Dict[str, Any]:
    register_namespaces()
    with zipfile.ZipFile(input_docx) as zf:
        files = {name: zf.read(name) for name in zf.namelist()}
    if "word/document.xml" not in files:
        raise SystemExit("word/document.xml missing; cannot apply visual refinements")

    root = ET.fromstring(files["word/document.xml"])
    engine = RefinementEngine(root)
    actions = plan.get("actions") or []
    log: List[Dict[str, Any]] = []
    for action in actions:
        if not isinstance(action, dict):
            log.append({"type": None, "status": "skipped", "detail": "action_not_an_object"})
            continue
        try:
            log.append(engine.apply(action))
        except Exception as exc:  # a single bad action must never sink the pass
            log.append({
                "type": action.get("type"),
                "target": action.get("target"),
                "status": "error",
                "detail": f"{type(exc).__name__}: {exc}",
            })

    files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    applied = sum(1 for item in log if item.get("status") == "applied")
    return {
        "inputDocx": str(input_docx),
        "outputDocx": str(output_docx),
        "actionCount": len(actions),
        "appliedCount": applied,
        "skippedCount": sum(1 for item in log if item.get("status") == "skipped"),
        "errorCount": sum(1 for item in log if item.get("status") == "error"),
        "executionLog": log,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply a whitelisted visual refinement plan to a DOCX.")
    parser.add_argument("input_docx", help="DOCX to refine (typically repaired_pass1.docx)")
    parser.add_argument("output_docx", help="Refined DOCX output path")
    parser.add_argument("--plan-json", required=True, help="visual_refinement_plan.json path")
    parser.add_argument("--report-json", help="Write execution report to this path")
    args = parser.parse_args()

    plan = json.loads(Path(args.plan_json).read_text(encoding="utf-8"))
    report = apply_plan(Path(args.input_docx).resolve(), Path(args.output_docx).resolve(), plan)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report_json:
        Path(args.report_json).write_text(text, encoding="utf-8")
    print(json.dumps({"applied": report["appliedCount"], "skipped": report["skippedCount"], "errors": report["errorCount"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
