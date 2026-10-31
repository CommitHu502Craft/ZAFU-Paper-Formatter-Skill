from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import xml.etree.ElementTree as ET

from docx import Document
from docx.enum.style import WD_STYLE_TYPE

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from apply_ooxml_fixes import (
    looks_like_english_title,
    normalize_front_abstract_block,
    split_english_front_label,
    style_spec_from_rules,
)
from docx_ooxml import NS, analyze_numbering_tree, load_rules, parse_document, parse_manual_numbering, qn
from extract_source_evidence import extract_docx_evidence
from formatter_core.sections import inspect_requirements
from thesis_ir import build_thesis_ir
from validate_visual_contracts import blank_paragraph_count

RULES = ROOT / "references" / "zafu_2022_rules.yaml"


class FrontMatterRegressionTests(unittest.TestCase):
    def test_english_front_label_split(self) -> None:
        self.assertEqual(split_english_front_label("Abstract"), ("abstract", "Abstract", ""))
        self.assertEqual(split_english_front_label("Abstract:"), ("abstract", "Abstract", ""))
        self.assertEqual(split_english_front_label("Abstract: body"), ("abstract", "Abstract", "body"))
        self.assertEqual(split_english_front_label("Key words: a; b"), ("keywords", "Key words", "a; b"))
        self.assertEqual(split_english_front_label("Keywords：a"), ("keywords", "Keywords", "a"))
        self.assertIsNone(split_english_front_label("English title")[0])
        self.assertTrue(looks_like_english_title("Olfaction in insects"))
        self.assertFalse(looks_like_english_title("Insects rely on a sensitive system. " * 3))

    def test_zafu_style_specs_match_rule_documents(self) -> None:
        specs = style_spec_from_rules(load_rules(str(RULES)))
        self.assertEqual(specs["zafu_caption"]["sizePt"], 9)
        self.assertEqual(specs["zafu_toc_body"]["spacing"]["linePt"], 20)
        self.assertEqual(specs["zafu_toc_body"]["spacing"]["lineRule"], "exact")
        self.assertEqual(specs["zafu_references_en"]["fonts"]["ascii"], "Times New Roman")
        self.assertEqual(specs["zafu_abstract_label_cn"]["fonts"]["eastAsia"], "黑体")

    def test_bare_abstract_label_merges_with_following_paragraph(self) -> None:
        root = ET.Element(qn("document"))
        body = ET.SubElement(root, qn("body"))

        def paragraph(text: str):
            node = ET.SubElement(body, qn("p"))
            run = ET.SubElement(node, qn("r"))
            ET.SubElement(run, qn("t")).text = text
            return node

        paragraph("测试论文题目")
        paragraph("摘要")
        paragraph("中文摘要正文内容。")
        paragraph("关键词：甲；乙")
        paragraph("Abstract")
        paragraph("English abstract body text.")
        paragraph("Key words: alpha; beta")
        rules = load_rules(str(RULES))
        specs = style_spec_from_rules(rules)
        changes = normalize_front_abstract_block(root, 0, 8, specs)
        self.assertTrue(any(item.get("action") == "merge_english_abstract_label" for item in changes))
        merged = [p for p in body.findall("w:p", NS) if "".join(t.text or "" for t in p.iter(qn("t"))).startswith("Abstract:")]
        self.assertEqual(len(merged), 1)
        text = "".join(t.text or "" for t in merged[0].iter(qn("t")))
        self.assertEqual(text, "Abstract: English abstract body text.")
        style_id = merged[0].find("w:pPr/w:pStyle", NS).get(qn("val"))
        self.assertEqual(style_id, "zafu_abstract_text_en")
        first_run = merged[0].find("w:r", NS)
        self.assertIsNotNone(first_run.find("w:rPr/w:b", NS))
        self.assertEqual(first_run.find("w:rPr/w:rFonts", NS).get(qn("eastAsia")), "Times New Roman")
        texts = ["".join(t.text or "" for t in p.iter(qn("t"))) for p in body.findall("w:p", NS)]
        self.assertNotIn("English abstract body text.", texts)

    def test_bare_labels_are_recovered_and_inventory_is_accurate(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "bare-labels.docx"
            document = Document()
            document.add_paragraph("测试论文题目")
            document.add_paragraph("摘要")
            document.add_paragraph("中文摘要正文内容足够长，用于识别。")
            document.add_paragraph("关键词：甲；乙")
            document.add_paragraph("Abstract")
            document.add_paragraph("English abstract body text for recognition.")
            document.add_paragraph("Key words: alpha; beta")
            document.save(source)
            audit = parse_document(str(source), rules_path=str(RULES), detail="semantic")
            roles = {(item.get("role") or {}).get("role") for item in audit["paragraphs"]}
            self.assertIn("title_cn", roles)
            self.assertIn("abstract_paragraph_cn", roles)
            self.assertIn("abstract_paragraph_en", roles)
            self.assertIn("keywords_en", roles)
            ir = build_thesis_ir(extract_docx_evidence(source, None, str(RULES), audit=audit))
            self.assertEqual(ir["frontMatter"]["title"], "测试论文题目")
            self.assertIn("中文摘要", ir["frontMatter"]["abstractCn"])
            self.assertIn("English abstract", ir["frontMatter"]["abstractEn"])
            missing = {item["key"] for item in inspect_requirements(audit, load_rules(str(RULES)), ir)["missing"]}
            for present_role in ("title_cn", "abstract_paragraph_cn", "keywords_cn", "abstract_paragraph_en", "keywords_en"):
                self.assertNotIn(present_role, missing)
            self.assertIn("title_en", missing)

    def test_unnumbered_introduction_does_not_shift_chapter_numbers(self) -> None:
        def heading(index: int, text: str):
            return {
                "index": index,
                "text": text,
                "role": {"role": "heading1", "confidence": 0.95, "family": "science_decimal"},
                "manualNumbering": parse_manual_numbering(text),
            }

        paragraphs = [
            heading(0, "引言"),
            heading(1, "1 研究背景"),
            heading(2, "2 研究方法"),
        ]
        result = analyze_numbering_tree(paragraphs)
        self.assertEqual(result["anomalies"], [])
        self.assertTrue(any(item.get("unnumbered") for item in result["headingTree"]))

    def test_blank_paragraph_count_ignores_template_front_matter(self) -> None:
        document = Document()
        document.styles.add_style("ZAFU Title CN", WD_STYLE_TYPE.PARAGRAPH)
        for _ in range(45):
            document.add_paragraph("")
        document.add_paragraph("题目", style="ZAFU Title CN")
        document.add_paragraph("摘要：正文。")
        document.add_paragraph("")
        self.assertEqual(blank_paragraph_count(document), 1)

    def test_chapter_page_break_skips_section_start(self) -> None:
        from apply_ooxml_fixes import ensure_chapter_page_breaks, paragraph_nodes

        root = ET.Element(qn("document"))
        body = ET.SubElement(root, qn("body"))

        def paragraph(text: str, style: str | None = None, sect: bool = False):
            node = ET.SubElement(body, qn("p"))
            ppr = ET.SubElement(node, qn("pPr"))
            if style:
                ET.SubElement(ppr, qn("pStyle"), {qn("val"): style})
            if sect:
                ET.SubElement(ppr, qn("sectPr"))
            run = ET.SubElement(node, qn("r"))
            ET.SubElement(run, qn("t")).text = text
            return node

        paragraph("", sect=True)
        paragraph("引言", "zafu_heading1")
        paragraph("1 研究背景", "zafu_heading1")
        body_paragraphs = paragraph_nodes(body)
        log = ensure_chapter_page_breaks(body_paragraphs, 0, None)
        self.assertIsNone(body_paragraphs[1].find("w:pPr/w:pageBreakBefore", NS))
        self.assertIsNotNone(body_paragraphs[2].find("w:pPr/w:pageBreakBefore", NS))
        self.assertEqual(len(log), 1)

    def test_introduction_moves_body_section_boundary(self) -> None:
        from apply_ooxml_fixes import promote_unnumbered_introduction, paragraph_nodes

        root = ET.Element(qn("document"))
        body = ET.SubElement(root, qn("body"))

        def paragraph(text: str, style: str | None = None, sect: bool = False):
            node = ET.SubElement(body, qn("p"))
            ppr = ET.SubElement(node, qn("pPr"))
            if style:
                ET.SubElement(ppr, qn("pStyle"), {qn("val"): style})
            if sect:
                ET.SubElement(ppr, qn("sectPr"))
            run = ET.SubElement(node, qn("r"))
            ET.SubElement(run, qn("t")).text = text
            return node

        paragraph("")
        paragraph("引言", "zafu_body")
        paragraph("", sect=True)
        paragraph("1 研究背景", "zafu_heading1")
        specs = style_spec_from_rules(load_rules(str(RULES)))
        log = promote_unnumbered_introduction(root, 0, specs)
        paragraphs = paragraph_nodes(body)
        self.assertIn("move_body_section_break_before_introduction", [item["action"] for item in log])
        self.assertIsNotNone(paragraphs[0].find("w:pPr/w:sectPr", NS))
        self.assertIsNone(paragraphs[2].find("w:pPr/w:sectPr", NS))
        self.assertEqual(paragraphs[1].find("w:pPr/w:pStyle", NS).get(qn("val")), "zafu_heading1")


if __name__ == "__main__":
    unittest.main()

