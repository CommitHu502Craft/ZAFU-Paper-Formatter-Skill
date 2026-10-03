from __future__ import annotations

import copy
import sys
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from apply_ooxml_fixes import replace_header_text
from docx_ooxml import NS, qn
from formatter_core.word_headers import apply_header_policy, check_header_policy, header_sections
from formatter_core.word_semantics import apply_ir_corrections
from thesis_format import PROFILES, docx_commands, text_commands
from thesis_ir import apply_semantic_overrides


class SemanticCorrectionTests(unittest.TestCase):
    def make_ir(self):
        return {
            "sourceType": "docx",
            "semanticBlocks": [{
                "id": "block-00042", "role": "heading_1", "kind": "heading",
                "level": 1, "text": "Original wording", "docxParagraphIndex": 42,
                "sourceAnchor": {"kind": "docx_paragraph", "index": 42},
            }],
            "headingTree": [{"text": "Original wording", "sourceIndex": 42}],
        }

    def test_unknown_id_is_not_guessed_by_number(self):
        ir = self.make_ir()
        apply_semantic_overrides(ir, {"overrides": [{"blockId": "wrong-42", "role": "body"}]})
        self.assertEqual(ir["semanticOverrides"]["appliedCount"], 0)
        self.assertEqual(ir["semanticOverrides"]["rejectedCount"], 1)
        self.assertEqual(ir["semanticBlocks"][0]["role"], "heading_1")

    def test_body_correction_clears_last_heading(self):
        ir = self.make_ir()
        apply_semantic_overrides(ir, {"overrides": [{"blockId": "block-00042", "role": "body"}]})
        self.assertEqual(ir["headingTree"], [])
        self.assertEqual(ir["semanticBlocks"][0]["kind"], "paragraph")
        self.assertIsNone(ir["semanticBlocks"][0]["level"])
        self.assertEqual(ir["semanticBlocks"][0]["text"], "Original wording")

    def test_correction_reaches_in_place_style_and_removes_stale_numbering(self):
        ir = self.make_ir()
        apply_semantic_overrides(ir, {"overrides": [{"blockId": "block-00042", "role": "body"}]})
        plan = {
            "profileStyleTargets": {"body_text": "zafu_body"},
            "styleMapping": {"42": "zafu_heading1"}, "paragraphRoles": {},
            "numberingActions": [{"paragraphIndex": 42}, {"paragraphIndex": 43}],
        }
        apply_ir_corrections(plan, {"paragraphs": [{"index": 42}]}, ir)
        self.assertEqual(plan["styleMapping"]["42"], "zafu_body")
        self.assertEqual(plan["numberingActions"], [{"paragraphIndex": 43}])
        self.assertEqual(len(plan["semanticCorrectionReport"]["applied"]), 1)


class HeaderPolicyTests(unittest.TestCase):
    def make_package(self, mixed=False):
        document = ET.Element(qn("document"))
        body = ET.SubElement(document, qn("body"))
        toc = ET.SubElement(body, qn("p"))
        properties = ET.SubElement(toc, qn("pPr"))
        ET.SubElement(properties, qn("pStyle"), {qn("val"): "zafu_heading1"})
        ET.SubElement(ET.SubElement(toc, qn("r")), qn("t")).text = "目录"
        if not mixed:
            section = ET.SubElement(properties, qn("sectPr"))
            ET.SubElement(section, qn("headerReference"), {qn("type"): "default", f"{{{NS['r']}}}id": "rId7"})
        paragraph = ET.SubElement(body, qn("p"))
        properties = ET.SubElement(paragraph, qn("pPr"))
        ET.SubElement(properties, qn("pStyle"), {qn("val"): "zafu_heading1"})
        ET.SubElement(ET.SubElement(paragraph, qn("r")), qn("t")).text = "1 Introduction"
        section = ET.SubElement(body, qn("sectPr"))
        if mixed:
            ET.SubElement(section, qn("headerReference"), {qn("type"): "default", f"{{{NS['r']}}}id": "rId7"})
        header = ET.Element(qn("hdr"))
        ET.SubElement(ET.SubElement(ET.SubElement(header, qn("p")), qn("r")), qn("t")).text = "Old title"
        files = {
            "word/_rels/document.xml.rels": (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header7.xml"/>'
                '</Relationships>'
            ).encode(),
            "[Content_Types].xml": b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
            "word/header7.xml": ET.tostring(header),
        }
        return document, files

    def test_shared_header_isolated_and_toc_blank(self):
        document, files = self.make_package()
        original_header = files["word/header7.xml"]
        apply_header_policy(document, files, "School thesis", "fixed", replace_header_text)
        sections = header_sections(document, files)
        self.assertEqual([item["scope"] for item in sections], ["toc", "body"])
        paths = [item["headers"]["default"]["path"] for item in sections]
        self.assertNotEqual(paths[0], paths[1])
        self.assertFalse("".join(ET.fromstring(files[paths[0]]).itertext()).strip())
        self.assertIn(b"School thesis", files[paths[1]])
        self.assertEqual(files["word/header7.xml"], original_header)
        self.assertEqual(check_header_policy(document, files, "School thesis", "fixed"), [])

    def test_mixed_scope_preserved_and_reported(self):
        document, files = self.make_package(mixed=True)
        original = ET.tostring(document)
        changes = apply_header_policy(document, files, "School thesis", "fixed", replace_header_text)
        self.assertEqual(ET.tostring(document), original)
        self.assertTrue(any(item["action"] == "header_manual_review" for item in changes))
        self.assertEqual(check_header_policy(document, files, "School thesis", "fixed")[0]["type"], "header_scope_needs_review")

    def test_preserve_mode_is_noop(self):
        document, files = self.make_package()
        original_files = copy.deepcopy(files)
        original_document = ET.tostring(document)
        self.assertEqual(apply_header_policy(document, files, "", "preserve", replace_header_text), [])
        self.assertEqual(files, original_files)
        self.assertEqual(ET.tostring(document), original_document)

    def test_missing_header_is_not_created(self):
        document, files = self.make_package()
        for section in document.iter(qn("sectPr")):
            for reference in list(section):
                section.remove(reference)
        apply_header_policy(document, files, "School thesis", "fixed", replace_header_text)
        self.assertEqual(list(document.iter(qn("headerReference"))), [])
        self.assertEqual(len([name for name in files if name.startswith("word/header")]), 1)


class CommandBoundaryTests(unittest.TestCase):
    def test_original_docx_requires_explicit_structural_permission(self):
        commands = docx_commands(Path("source.docx"), PROFILES["zafu_2022"], "conservative-repair", Path("work"), thesis_ir=Path("work/thesis_ir.json"))
        planner = next(command for command in commands if "scripts/plan_docx_repairs.py" in command)
        self.assertIn("--thesis-ir-json", planner)
        self.assertTrue(all("--allow-structural-rebuild" not in command for command in commands))
        approved = docx_commands(Path("source.docx"), PROFILES["zafu_2022"], "conservative-repair", Path("work"), allow_structural_rebuild=True)
        self.assertEqual(sum("--allow-structural-rebuild" in command for command in approved), 2)

    def test_text_generation_keeps_template_permission(self):
        commands = text_commands(Path("source.md"), PROFILES["zafu_2022"], Path("work"))
        self.assertEqual(sum("--allow-structural-rebuild" in command for command in commands), 2)

    def test_audit_only_does_not_apply_repairs(self):
        commands = docx_commands(Path("source.docx"), PROFILES["zafu_2022"], "audit-only", Path("work"), allow_structural_rebuild=True)
        self.assertFalse(any("scripts/apply_ooxml_fixes.py" in command for command in commands))


if __name__ == "__main__":
    unittest.main()
