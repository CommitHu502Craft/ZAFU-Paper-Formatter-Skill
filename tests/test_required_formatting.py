from __future__ import annotations

import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, NumberObject, DecodedStreamObject, ArrayObject

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from docx_ooxml import NS, qn, parse_document, build_repair_plan, load_rules, normalize_rel_target_word
from extract_source_evidence import extract_docx_evidence
from thesis_ir import build_thesis_ir, extract_references_section
from formatter_core.sections import is_references_heading, front_label, inspect_requirements
from formatter_core.production import normalize_docx
from formatter_core.word_headers import check_header_policy, apply_header_policy, separate_toc_and_abstract, header_sections
from formatter_core.integrity import verify
from formatter_core.reference_format import check_reference_format
from formatter_core.page_layout import extract_image_geometry, caption_geometry_findings
from render_validate_docx import analyze_text_layout
from apply_ooxml_fixes import replace_header_text, normalize_references_section, style_spec_from_rules, apply_section_geometry_from_template

RULES = ROOT / "references" / "zafu_2022_rules.yaml"


class RequiredFormattingTests(unittest.TestCase):
    def test_reference_heading_variations(self):
        for text in ("参考文献：", "参考文献:", "参 考 文 献", "References:", "6 参考文献", "第六章 参考文献"):
            with self.subTest(text=text):
                self.assertTrue(is_references_heading(text))
        self.assertFalse(is_references_heading("参考文献的选择方法"))
        self.assertEqual(front_label("Abstract Catalysts remove pollutants."), ("abstract_paragraph_en", "Catalysts remove pollutants."))
        self.assertEqual(front_label("Keywords single atom catalysts"), ("keywords_en", "single atom catalysts"))
        self.assertEqual(normalize_rel_target_word("word/_rels/document.xml.rels", "header1.xml"), "word/header1.xml")
        self.assertEqual(normalize_rel_target_word("word/_rels/document.xml.rels", "/word/footer1.xml"), "word/footer1.xml")

    def test_source_like_paper_has_real_missing_list_and_21_references(self):
        with tempfile.TemporaryDirectory() as directory:
            source, target = Path(directory) / "source.docx", Path(directory) / "formatted.docx"
            doc = Document()
            doc.add_paragraph("铁基单原子催化剂在水污染治理中的应用", "Heading 2").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("某同学 2020123456")
            doc.add_paragraph("摘要：本文介绍铁基催化剂。")
            doc.add_paragraph("关键词：催化剂；环境")
            doc.add_paragraph("Application of catalysts in water pollution treatment").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Abstract Single-atom catalysts treat pollutants.")
            doc.add_paragraph("Keywords single atom catalysts, water treatment")
            doc.add_paragraph("1 引言", "Heading 1")
            doc.add_paragraph("已有研究[1]显示催化剂有效。")
            doc.add_paragraph("参考文献：")
            for number in range(1, 22):
                doc.add_paragraph(f"[{number}] Smith A. Catalyst study {number}[J]. Journal, 2020, 1: 1-9.")
            doc.sections[0].different_first_page_header_footer = True
            doc.settings.odd_and_even_pages_header_footer = True
            doc.save(source)
            audit = parse_document(str(source), rules_path=str(RULES), detail="semantic")
            ir = build_thesis_ir(extract_docx_evidence(source, None, str(RULES), audit=audit))
            self.assertEqual(len(ir["references"]["entries"]), 21)
            self.assertEqual(ir["frontMatter"]["title"], "铁基单原子催化剂在水污染治理中的应用")
            self.assertIn("Application of catalysts", ir["frontMatter"]["titleEn"])
            self.assertIn("Single-atom", ir["frontMatter"]["abstractEn"])
            rules = load_rules(str(RULES))
            missing = {i["key"] for i in inspect_requirements(audit, rules, ir)["missing"]}
            self.assertEqual(missing, {"cover", "integrity_statement", "toc", "body_header", "page_numbers"})
            plan = build_repair_plan(audit)
            self.assertEqual(plan["styleMapping"]["10"], "zafu_references_en")
            normalize_docx(source, target, audit, plan, rules)
            self.assertTrue(verify(source, target)["passed"])
            with zipfile.ZipFile(target) as package:
                files = {name: package.read(name) for name in package.namelist()}
            root = ET.fromstring(files["word/document.xml"])
            header_text = str(plan.get("headerText") or "")
            self.assertTrue(header_text)
            self.assertEqual(check_header_policy(root, files, header_text, rules["page"]["header_text_mode"]), [])
            references = root.findall("w:body/w:sectPr/w:headerReference", NS)
            self.assertEqual({ref.get(qn("type")) for ref in references}, {"default", "first", "even"})
            self.assertTrue(any(b'PAGE' in data for name, data in files.items() if name.startswith("word/footer_formatter_")))
            self.assertIn("[21] Smith", "".join(n.text or "" for n in root.iter(qn("t"))))
            after = parse_document(str(target), rules_path=str(RULES), detail="semantic")
            remaining = {i["key"] for i in inspect_requirements(after, rules)["missing"]}
            self.assertNotIn("body_header", remaining)
            self.assertNotIn("page_numbers", remaining)

    def test_reference_formatting_preserves_field_spanning_multiple_entries(self):
        root = ET.Element(qn("document"))
        body = ET.SubElement(root, qn("body"))
        def paragraph(text):
            p = ET.SubElement(body, qn("p"))
            ET.SubElement(ET.SubElement(p, qn("r")), qn("t")).text = text
            return p
        paragraph("正文已有引用[1]。")
        paragraph("参考文献：")
        first = paragraph("[1] Author. First[J]. Journal, 2020.")
        begin = ET.Element(qn("r"))
        ET.SubElement(begin, qn("fldChar"), {qn("fldCharType"): "begin"})
        first.insert(0, begin)
        separate = ET.Element(qn("r"))
        ET.SubElement(separate, qn("fldChar"), {qn("fldCharType"): "separate"})
        first.insert(1, separate)
        paragraph("[2] Author. Second[J]. Journal, 2021.")
        last = paragraph("")
        ET.SubElement(ET.SubElement(last, qn("r")), qn("fldChar"), {qn("fldCharType"): "end"})
        rules = load_rules(str(RULES))
        normalize_references_section(root, 0, style_spec_from_rules(rules), rules)
        self.assertEqual(len(body.findall("w:p", NS)), 5)
        self.assertEqual([n.get(qn("fldCharType")) for n in root.iter(qn("fldChar"))], ["begin", "separate", "end"])
        self.assertIn("[1] Author", "".join(n.text or "" for n in root.iter(qn("t"))))

    def test_absent_and_empty_headers_are_detected_then_created(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            doc = Document()
            doc.add_paragraph("1 引言", "Heading 1")
            doc.add_paragraph("正文。")
            doc.save(source)
            with zipfile.ZipFile(source) as package:
                files = {name: package.read(name) for name in package.namelist()}
            root = ET.fromstring(files["word/document.xml"])
            self.assertTrue(any(i["type"] == "missing_body_header" for i in check_header_policy(root, files, "学校规定页眉", "fixed")))
            apply_header_policy(root, files, "学校规定页眉", "fixed", replace_header_text)
            self.assertFalse(check_header_policy(root, files, "学校规定页眉", "fixed"))
            for name in list(files):
                if name.startswith("word/header_formatter_"):
                    files[name] = ET.tostring(ET.Element(qn("hdr")))
            self.assertTrue(any(i["type"] == "header_text_mismatch" for i in check_header_policy(root, files, "学校规定页眉", "fixed")))

    def test_numbered_bibliography_fallback_and_zero_entry_failure(self):
        evidence = {"sourceType": "docx", "paragraphEvidence": [{"index": n, "text": f"[{n+1}] Author. Title[J]. Journal, 2020."} for n in range(2)]}
        refs = extract_references_section(evidence)
        self.assertEqual(refs["detection"], "numbered_entries_fallback")
        self.assertEqual(len(refs["entries"]), 2)
        inventory = inspect_requirements({"paragraphs": []}, {"references": {"citation_mode": "numeric"}}, {"references": refs})
        self.assertIn("references_heading", {i["key"] for i in inventory["missing"]})
        report = check_reference_format({"references": {"entries": []}, "bodyBlocks": [{"text": "正文引用[1]。"}]})
        self.assertIn("reference_entries_not_recognized", [i["kind"] for i in report["issues"]])

    def test_numeric_linkage_is_checked_without_author_year_sort_rules(self):
        ir = {"references": {"entries": [{"text": "[1] Smith A. Title[J]. Journal, 2020."}, {"text": "[2] 张某. 标题[J]. 期刊, 2021."}]}, "bodyBlocks": [{"id": "body1", "text": "研究结果[1-3]。"}]}
        report = check_reference_format(ir, {"references": {"citation_mode": "author_year"}})
        kinds = {i["kind"] for i in report["issues"]}
        self.assertIn("citation_mode_mismatch", kinds)
        self.assertIn("citation_without_reference", kinds)
        self.assertNotIn("foreign_surname_not_uppercase", kinds)
        self.assertNotIn("chinese_entry_after_foreign_group", kinds)

    def test_template_blanks_are_not_missing_personal_information(self):
        template = ROOT / "assets" / "zafu_front_matter_template.docx"
        audit = parse_document(str(template), rules_path=str(RULES), detail="semantic")
        result = inspect_requirements(audit, load_rules(str(RULES)))
        self.assertNotIn("cover", {i["key"] for i in result["missing"]})
        self.assertNotIn("integrity_statement", {i["key"] for i in result["missing"]})
        self.assertNotIn("personal", str(result))

    def test_normalizing_body_preserves_template_front_pages_and_assets(self):
        template = ROOT / "assets" / "zafu_front_matter_template.docx"
        with tempfile.TemporaryDirectory() as directory:
            source, target = Path(directory) / "covered.docx", Path(directory) / "formatted.docx"
            doc = Document(template)
            doc.add_section(WD_SECTION_START.NEW_PAGE)
            doc.add_paragraph("摘要：本文测试封面保护。")
            doc.add_paragraph("1 引言", "Heading 1")
            doc.add_paragraph("正文内容不得改变。")
            doc.save(source)
            audit = parse_document(str(source), template_docx=str(template), rules_path=str(RULES), detail="semantic")
            plan = build_repair_plan(audit)
            prefix = plan["editableStartParagraph"]
            self.assertGreater(prefix, 20)
            normalize_docx(source, target, audit, plan, load_rules(str(RULES)))
            with zipfile.ZipFile(source) as old, zipfile.ZipFile(target) as new:
                old_body = ET.fromstring(old.read("word/document.xml")).find("w:body", NS)
                new_body = ET.fromstring(new.read("word/document.xml")).find("w:body", NS)
                old_prefix = old_body.findall("w:p", NS)[:prefix]
                new_prefix = new_body.findall("w:p", NS)[:prefix]
                self.assertEqual([ET.tostring(p) for p in old_prefix], [ET.tostring(p) for p in new_prefix])
                for name in old.namelist():
                    if name.startswith("word/media/"):
                        self.assertEqual(old.read(name), new.read(name))
            self.assertTrue(verify(source, target)["passed"])


class PageGeometryTests(unittest.TestCase):
    def test_toc_and_abstract_can_get_different_headers_without_touching_cover(self):
        root = ET.Element(qn("document"))
        body = ET.SubElement(root, qn("body"))
        def paragraph(text, style=None):
            p = ET.SubElement(body, qn("p"))
            props = ET.SubElement(p, qn("pPr"))
            if style:
                ET.SubElement(props, qn("pStyle"), {qn("val"): style})
            ET.SubElement(ET.SubElement(p, qn("r")), qn("t")).text = text
            return p, props
        cover, props = paragraph("本科生毕业论文（设计）")
        ET.SubElement(props, qn("sectPr"))
        before = ET.tostring(cover)
        paragraph("目录")
        toc, _ = paragraph("目录域结果")
        paragraph("中文题目", "zafu_title_cn")
        paragraph("摘要：摘要内容。")
        end = ET.SubElement(body, qn("sectPr"))
        ET.SubElement(end, qn("pgNumType"), {qn("fmt"): "upperRoman", qn("start"): "1"})
        self.assertTrue(separate_toc_and_abstract(root, 1))
        self.assertEqual(ET.tostring(cover), before)
        self.assertIsNotNone(toc.find("w:pPr/w:sectPr", NS))
        self.assertNotIn(qn("start"), end.find("w:pgNumType", NS).attrib)
        files = {"word/_rels/document.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'}
        self.assertEqual([s["scope"] for s in header_sections(root, files)], ["front_matter", "toc", "body"])
        self.assertEqual(separate_toc_and_abstract(root, 1), [])
        apply_section_geometry_from_template(root, [], 1, fallback_page_rules={"paper": "A4", "margin_top_cm": 2.5})
        self.assertEqual(ET.tostring(cover), before)

    def test_image_transform_in_nested_form_is_measured(self):
        with tempfile.TemporaryDirectory() as directory:
            writer = PdfWriter()
            page = writer.add_blank_page(600, 800)
            image = DecodedStreamObject()
            image.set_data(b'\0\0\0')
            image.update({NameObject("/Type"): NameObject("/XObject"), NameObject("/Subtype"): NameObject("/Image"), NameObject("/Width"): NumberObject(1), NameObject("/Height"): NumberObject(1), NameObject("/BitsPerComponent"): NumberObject(8), NameObject("/ColorSpace"): NameObject("/DeviceRGB")})
            resources = DictionaryObject({NameObject("/XObject"): DictionaryObject({NameObject("/Im1"): writer._add_object(image)})})
            form = DecodedStreamObject()
            form.set_data(b'q 100 0 0 80 10 20 cm /Im1 Do Q')
            form.update({NameObject("/Type"): NameObject("/XObject"), NameObject("/Subtype"): NameObject("/Form"), NameObject("/BBox"): ArrayObject([NumberObject(n) for n in (0, 0, 200, 200)]), NameObject("/Resources"): resources})
            page[NameObject("/Resources")] = DictionaryObject({NameObject("/XObject"): DictionaryObject({NameObject("/Fm1"): writer._add_object(form)})})
            content = DecodedStreamObject()
            content.set_data(b'q 2 0 0 2 50 100 cm /Fm1 Do Q')
            page[NameObject("/Contents")] = writer._add_object(content)
            path = Path(directory) / "figure.pdf"
            writer.write(path)
            measured = extract_image_geometry(path)
            self.assertEqual(measured[0]["images"][0]["boundingBox"], [70, 140, 270, 300])

    def test_caption_position_uses_image_boxes_even_with_running_header(self):
        pages = [{"fragments": [{"text": "学校页眉", "boundingBox": [60, 760, 400, 775], "coordinateAccuracy": "pdf_text_bbox"}, {"text": "图1 实验结果", "boundingBox": [180, 610, 370, 624], "coordinateAccuracy": "pdf_text_bbox"}]}]
        images = [{"images": [{"boundingBox": [100, 350, 450, 600]}]}]
        findings = caption_geometry_findings(pages, images)
        self.assertEqual(findings[0]["type"], "figure_caption_above_image")
        pages[0]["fragments"][1]["boundingBox"] = [180, 325, 370, 339]
        self.assertEqual(caption_geometry_findings(pages, images), [])
        self.assertEqual(caption_geometry_findings(pages, [{"images": []}]), [])
        self.assertEqual(analyze_text_layout(["前页正文", "学校页眉\n图1 实验结果\n正文"], {}, None), [])


if __name__ == "__main__":
    unittest.main()
