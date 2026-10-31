from __future__ import annotations
import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from docx_ooxml import parse_document, load_rules, build_repair_plan, qn
from formatter_core.production import inspect_risks, normalize_docx
from formatter_core.session import analysis_session
from formatter_core.integrity import verify, relationship_issues, verify_source_assets
from formatter_core.visual import anchor_findings, batch_plan, merge_actions
import render_validate_docx as render
from thesis_format import parse_args

class ProductionTests(unittest.TestCase):
    def test_new_section_boundary_is_not_reused_for_heading(self):
        from docx.enum.section import WD_SECTION_START
        from build_docx_from_markdown import acquire_output_paragraph
        document = Document()
        document.add_paragraph("Abstract")
        document.add_section(WD_SECTION_START.NEW_PAGE)
        boundary = document.paragraphs[-1]._p
        heading = acquire_output_paragraph(document, {"reuse_last_empty_paragraph": True}, style="Heading 1")
        heading.add_run("1 Introduction")
        self.assertIsNot(boundary, heading._p)
        self.assertIsNotNone(boundary.find(qn("pPr") + "/" + qn("sectPr")))
        self.assertIsNone(heading._p.find(qn("pPr") + "/" + qn("sectPr")))

    def test_quality_defaults_and_legacy_check_override(self):
        for flags, expected in (([], "visual"), (["--quality", "fast"], "structural"), (["--check", "layout"], "layout"), (["--check", "structural"], "structural")):
            with patch.object(sys, "argv", ["thesis_format.py", "paper.docx", *flags]):
                self.assertEqual(parse_args().check, expected)

    def test_batch_preserves_text_emphasis_native_region_and_section(self):
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory) / "input.docx", Path(directory) / "out.docx"
            document = Document()
            paragraph = document.add_paragraph("1 Introduction")
            paragraph = document.add_paragraph()
            run = paragraph.add_run("H2O biological sample")
            run.italic = True
            run.bold = True
            run.font.subscript = True
            run.font.size = Pt(19)
            formula = document.add_paragraph("Native formula")
            math = OxmlElement("m:oMath")
            formula._p.append(math)
            document.add_paragraph("Following body should normalize")
            document.add_table(rows=2, cols=2).cell(0, 0).text = "Data"
            document.save(source)
            rules_path = ROOT / "references/zafu_2022_rules.yaml"
            audit = parse_document(str(source), rules_path=str(rules_path), detail="semantic")
            plan = build_repair_plan(audit)
            plan["styleMapping"].update({"1": "zafu_body", "3": "zafu_body"})
            original_math = ET.canonicalize(ET.tostring(inspect_risks(source)[1].find(qn("body"))[2]), rewrite_prefixes=True)
            result = normalize_docx(source, output, audit, plan, load_rules(str(rules_path)))
            self.assertTrue(result["preservedRegions"])
            self.assertTrue(verify(source, output)["passed"])
            self.assertEqual(original_math, ET.canonicalize(ET.tostring(inspect_risks(output)[1].find(qn("body"))[2]), rewrite_prefixes=True))
            actual = Document(output)
            run = actual.paragraphs[1].runs[0]
            self.assertTrue(run.italic)
            self.assertTrue(run.bold)
            self.assertTrue(run.font.subscript)
            self.assertNotEqual(run.font.size, Pt(19))
            self.assertEqual(actual.paragraphs[3].style.style_id, "zafu_body")
            self.assertEqual(relationship_issues(output), [])

    def test_comment_range_interior_is_untouched(self):
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory) / "in.docx", Path(directory) / "out.docx"
            doc = Document()
            for text in ("Range start", "Interior", "Range end", "Outside"):
                doc.add_paragraph(text)
            for index, tag in ((0, "commentRangeStart"), (2, "commentRangeEnd")):
                marker = OxmlElement("w:" + tag)
                marker.set(qn("id"), "42")
                doc.paragraphs[index]._p.append(marker)
            doc.save(source)
            original = [ET.tostring(c) for c in inspect_risks(source)[1].find(qn("body"))][:3]
            audit = parse_document(str(source), detail="semantic")
            plan = build_repair_plan(audit)
            plan["styleMapping"] = {str(i): "zafu_body" for i in range(4)}
            normalize_docx(source, output, audit, plan, load_rules(str(ROOT / "references/zafu_2022_rules.yaml")))
            final = [ET.tostring(c) for c in inspect_risks(output)[1].find(qn("body"))][:3]
            self.assertEqual(original, final)

    def test_cache_is_content_based_and_mutations_are_isolated(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "rules.yaml"
            source.write_text("value: one\n", encoding="utf-8")
            with analysis_session(Path(directory) / "cache") as session:
                first = load_rules(str(source))
                first["value"] = "mutated"
                self.assertEqual(load_rules(str(source))["value"], "one")
                source.write_text("value: two\n", encoding="utf-8")
                self.assertEqual(load_rules(str(source))["value"], "two")
                self.assertEqual(session.metrics["cacheHits"], 1)

    def test_object_loss_fails_even_when_text_is_identical(self):
        with tempfile.TemporaryDirectory() as directory:
            first, second = Path(directory) / "a.docx", Path(directory) / "b.docx"
            document = Document()
            document.add_paragraph("Same text")
            document.add_table(rows=1, cols=1)
            document.save(first)
            document._element.body.remove(document.tables[0]._tbl)
            document.save(second)
            self.assertFalse(verify(first, second)["passed"])

    def test_original_cannot_be_used_as_normalization_destination(self):
        with self.assertRaises(ValueError):
            normalize_docx(ROOT / "input.docx", ROOT / "input.docx", {}, {}, {})

    def test_missing_source_image_is_an_integrity_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "built.docx"
            Document().save(source)
            result = verify_source_assets({"semanticBlocks": [{"kind": "image"}]}, source)
            self.assertFalse(result["passed"])
            self.assertEqual(result["issues"][0]["expected"], 1)

    def test_template_cache_survives_equivalent_rule_paths_and_invalidates_changes(self):
        import docx_ooxml
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            template = base / "template.docx"
            Document().save(template)
            one, two = base / "one.yaml", base / "two.yaml"
            one.write_text("page: {}", encoding="utf-8")
            two.write_bytes(one.read_bytes())
            with analysis_session(base / "cache"):
                docx_ooxml.extract_template_baseline(str(template), str(one))
            with analysis_session(base / "cache") as session:
                docx_ooxml.extract_template_baseline(str(template), str(two))
                self.assertEqual(session.metrics["cacheHits"], 1)
            two.write_text("page: {paper: A4}", encoding="utf-8")
            with analysis_session(base / "cache") as session:
                docx_ooxml.extract_template_baseline(str(template), str(two))
                self.assertGreater(session.metrics["cacheMisses"], 0)

    def test_empty_native_formula_cannot_be_removed_as_blank_paragraph(self):
        from apply_visual_refinements import apply_plan
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory) / "input.docx", Path(directory) / "output.docx"
            document = Document()
            paragraph = document.add_paragraph()
            paragraph._p.append(OxmlElement("m:oMath"))
            document.save(source)
            result = apply_plan(source, output, {"actions": [{"type": "remove_empty_paragraph", "target": {"paragraphIndex": 0}}]})
            self.assertEqual(result["appliedCount"], 0)
            self.assertTrue(verify(source, output)["passed"])

    def test_stale_visual_plan_is_rejected_before_mutation(self):
        from apply_visual_refinements import apply_plan
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory) / "input.docx", Path(directory) / "out.docx"
            Document().save(source)
            with self.assertRaises(ValueError):
                apply_plan(source, output, {"candidateSha256": "wrong", "actions": []})
            self.assertFalse(output.exists())

    def test_visual_review_requires_current_candidate_and_all_inspected_images(self):
        from record_visual_review import record_review
        from formatter_core.workspace import Workspace, file_digest
        from PIL import Image
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.docx"
            Document().save(source)
            workspace = Workspace.create(source, "zafu_2022")
            candidate = workspace.deliverables / "candidate.docx"
            candidate.write_bytes(source.read_bytes())
            image = workspace.work / "critical.png"
            Image.new("RGB", (20, 20), "white").save(image)
            render_path = workspace.reports / "render.json"
            render_path.write_text(json.dumps({"candidateSha256": file_digest(candidate), "contactSheets": [str(image)], "criticalPages": [{"imagePath": str(image), "status": "exported"}], "visualFindings": []}), encoding="utf-8")
            summary = workspace.reports / "summary.json"
            summary.write_text("{}", encoding="utf-8")
            workspace.write_manifest({"deliverables": {"docx": str(candidate)}, "artifacts": {"renderValidationReport": str(render_path), "summary": str(summary)}})
            record = {"reviewer": "unit-test-simulation", "candidateSha256": file_digest(candidate), "reviewedImages": []}
            with self.assertRaises(ValueError):
                record_review(workspace.root, record)
            record["reviewedImages"] = [{"path": str(image), "sha256": file_digest(image)}]
            with patch("generate_review_report.main"):
                self.assertTrue(record_review(workspace.root, record)["visualReviewed"])
            candidate.write_bytes(b"changed")
            with self.assertRaises(ValueError):
                record_review(workspace.root, record)

    def test_ambiguous_anchor_never_becomes_an_auto_action(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "a.docx"
            document = Document()
            for _ in range(2):
                document.add_paragraph("1 Repeated heading")
            document.add_paragraph("1 引言", style="Heading 1")
            document.save(source)
            findings = anchor_findings(source, [{"type": "heading_orphan_at_page_bottom_geometry", "page": 1, "headingText": "1 Repeated heading", "confidence": 0.99}], [])
            self.assertEqual(findings[0]["mappingStatus"], "ambiguous")
            short = anchor_findings(source, [{"type": "heading_orphan_at_page_bottom_geometry", "page": 1, "headingText": "1 引言"}], [])
            self.assertEqual(short[0]["docxParagraphIndex"], 2)
            self.assertEqual(short[0]["mappingStatus"], "unique")
            self.assertEqual(batch_plan(findings, "hash")["actions"], [])

    def test_conflicting_actions_are_deferred_and_duplicates_merged(self):
        target = {"paragraphIndex": 1, "textPrefix": "Heading"}
        plan = merge_actions({"actions": [{"type": "set_page_break_before", "target": target}, {"type": "clear_page_break_before", "target": target}, *[{"type": "set_keep_lines", "target": target}] * 2]})
        self.assertEqual(len(plan["actions"]), 1)
        self.assertEqual(len(plan["conflictingActions"]), 2)

    def test_changed_candidate_invalidates_pdf_conversion(self):
        from pypdf import PdfWriter
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source = base / "a.docx"
            Document().save(source)
            def convert(docx, out, **kwargs):
                out.mkdir(parents=True, exist_ok=True)
                pdf = out / "a.pdf"
                writer = PdfWriter()
                writer.add_blank_page(width=595, height=842)
                writer.write(pdf)
                return {"available": True, "pdfPath": str(pdf), "backend": "fake"}
            with patch.object(render, "convert_docx_to_pdf", side_effect=convert) as conversion:
                first = render.run_render_validation(str(source), str(base / "render"), analyze_layout=False)
                render.run_render_validation(str(source), str(base / "render"), analyze_layout=False)
                self.assertEqual(conversion.call_count, 1)
                document = Document(source)
                document.add_paragraph("Changed")
                document.save(source)
                second = render.run_render_validation(str(source), str(base / "render"), analyze_layout=False)
                self.assertEqual(conversion.call_count, 2)
                self.assertNotEqual(first["candidateSha256"], second["candidateSha256"])
                self.assertFalse(second["visualReviewed"])

if __name__ == "__main__":
    unittest.main()
