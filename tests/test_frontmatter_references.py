from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfWriter


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from build_docx_from_markdown import add_reference_paragraph_block
from extract_source_evidence import extract_text_evidence
from thesis_ir import build_thesis_ir
from formatter_backends.latex import build_project, compile_project, escape_text, render_inline
from formatter_backends.latex_frontmatter import load_metadata, prepare_frontmatter
from formatter_backends.latex_bibliography import prepare_bibliography, render_citations
from formatter_core.reference_format import check_reference_format


class FrontmatterTests(unittest.TestCase):
    def prepare(self, text: str, metadata: dict | None = None) -> dict:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.md"
            source.write_text(text, encoding="utf-8")
            thesis = build_thesis_ir(extract_text_evidence(source))
            return prepare_frontmatter(thesis, metadata or {}, escape_text, lambda content, identifier: render_inline(content, [], identifier))

    def test_bilingual_abstracts_call_original_macros(self) -> None:
        result = self.prepare("# 测试论文\n\n# 摘要\n\n中文摘要。\n\n关键词：中文\n\n# Abstract\n\nEnglish summary.\n\nKeywords: English\n\n# 1 引言\n\n正文。", {"titleEn": "English title"})
        self.assertIn(r"\ZhAbstract{中文摘要。}{中文}", result["frontmatterTex"])
        self.assertIn(r"\EnAbstract{English summary.}{English}", result["frontmatterTex"])
        self.assertEqual(result["frontmatterTex"].count(r"\customContent"), 1)
        self.assertNotIn("正文", result["frontmatterTex"])
        self.assertNotIn(r"\customCover", result["frontmatterTex"])
        self.assertNotIn(r"\makestatement", result["frontmatterTex"])

    def test_inline_abstract_label_is_supported(self) -> None:
        result = self.prepare("测试论文\n\n摘要：中文摘要。\n\n关键词：测试\n\n1 引言\n\n正文。")
        self.assertIn(r"\ZhAbstract{中文摘要。}{测试}", result["frontmatterTex"])

    def test_missing_cover_metadata_blocks_generation(self) -> None:
        result = self.prepare("# 论文\n\n# 1 引言\n\n正文。", {"cover": True})
        self.assertEqual(result["errors"][0]["kind"], "missing_cover_metadata")
        self.assertNotIn(r"\customCover", result["frontmatterTex"])

    def test_explicit_statement_is_unsigned_and_metadata_escaped(self) -> None:
        result = self.prepare("# 论文\n\n# 1 引言\n\n正文。", {"statement": True, "author": r"张\input{secret}", "date": "2026-10-01"})
        self.assertEqual(result["errors"], [])
        self.assertIn(r"\makestatement", result["frontmatterTex"])
        self.assertIn(r"\setsignature{}", result["metadataTex"])
        self.assertNotIn(r"\input{secret}", result["metadataTex"])
        self.assertIn(r"\textbackslash{}input", result["metadataTex"])

    def test_metadata_schema_rejects_unknown_fields_and_types(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "metadata.json"
            for value in ({"cover": "yes"}, {"author": 123}, {"signature": "guessed"}, []):
                path.write_text(json.dumps(value), encoding="utf-8")
                with self.assertRaises(ValueError):
                    load_metadata(path)

    def test_frontmatter_removed_from_body_without_placeholder_expansion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.md"
            source.write_text("# 测试论文\n\n# 摘要\n\n唯一摘要。\n\n关键词：测试\n\n# 1 引言\n\n正文 @@METADATA@@。", encoding="utf-8")
            project = Path(directory) / "project"
            with patch("formatter_backends.latex.install_external_class", return_value={"modified": False}):
                report = build_project(build_thesis_ir(extract_text_evidence(source)), source, {}, ROOT / "profiles/zafu_2022", project)
            tex = (project / "main.tex").read_text(encoding="utf-8")
            self.assertTrue(report["passed"])
            self.assertEqual(tex.count("唯一摘要"), 1)
            self.assertIn("正文 @@METADATA@@", tex)
            self.assertNotIn(r"\section*{摘要}", tex)
            self.assertNotIn(r"\makestatement", tex)


class BibliographyTests(unittest.TestCase):
    def prepare(self, content: str) -> dict:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.bib"
            path.write_text(content, encoding="utf-8")
            original = path.read_bytes()
            project = Path(directory) / "project"
            project.mkdir()
            result = prepare_bibliography(path, project)
            self.assertEqual(path.read_bytes(), original)
            return result

    def test_plain_unicode_bibtex_input_is_preserved(self) -> None:
        result = self.prepare("@book{chen1989, author={陈守常 and 曾大鹏}, title={油茶病害}, year={1989}, key={chen shou chang}}")
        self.assertEqual(result["errors"], [])
        self.assertEqual(result["keys"], ["chen1989"])

    def test_chinese_authors_require_verified_pinyin_key(self) -> None:
        result = self.prepare("@book{chen1989, author={陈守常}, title={油茶病害}, year={1989}}")
        self.assertIn("missing_chinese_sort_key", [issue["kind"] for issue in result["errors"]])

    def test_bibtex_injection_and_preamble_are_rejected(self) -> None:
        result = self.prepare(r"@book{bad, author={Smith, Alice}, title={\input{secret}}, year={2020}}")
        self.assertIn("unsafe_bibliography_command", [issue["kind"] for issue in result["errors"]])
        result = self.prepare('@preamble{"unsafe"}\n@book{good,author={Smith},title={Title},year={2020}}')
        self.assertIn("unsafe_bibliography_preamble", [issue["kind"] for issue in result["errors"]])

    def test_duplicate_keys_and_unparseable_entries_are_rejected(self) -> None:
        result = self.prepare("@book{same,title={One}}\n@book{same,title={Two}}")
        self.assertIn("invalid_or_duplicate_citation_key", [issue["kind"] for issue in result["errors"]])
        self.assertTrue(self.prepare("@book{broken, title={unterminated}")["errors"])

    def test_citation_markers_are_explicit_and_missing_keys_fail(self) -> None:
        used = set()
        issues = []
        rendered = render_citations("研究 [@chen; @reich]。", {"chen", "reich"}, used, escape_text, issues, "b1")
        self.assertIn(r"\citep{chen,reich}", rendered)
        self.assertEqual(used, {"chen", "reich"})
        render_citations("缺失 [@unknown]。", set(), used, escape_text, issues, "b2")
        self.assertEqual(issues[0]["kind"], "missing_citation_key")

    def test_citation_marker_in_code_is_not_executed(self) -> None:
        issues = []
        rendered = render_citations("`[@chen]`", {"chen"}, set(), lambda content: render_inline(content, issues, "b1"), issues, "b1")
        self.assertNotIn(r"\citep", rendered)
        self.assertIn(r"\texttt{[@chen]}", rendered)

    def test_citations_preserve_emphasis_and_do_not_execute_math_content(self) -> None:
        issues = []
        rendered = render_citations("**强调 [@chen]** 和 $[@chen]$", {"chen"}, set(), lambda content: render_inline(content, issues, "b1"), issues, "b1")
        self.assertIn(r"\textbf{强调 \citep{chen}}", rendered)
        self.assertEqual(rendered.count(r"\citep"), 1)

    def test_markers_without_bibliography_block_compilation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.md"
            source.write_text("# 论文\n\n# 1 引言\n\n研究 [@unknown]。", encoding="utf-8")
            with patch("formatter_backends.latex.install_external_class", return_value={"modified": False}):
                report = build_project(build_thesis_ir(extract_text_evidence(source)), source, {}, ROOT / "profiles/zafu_2022", Path(directory) / "project")
            self.assertFalse(report["passed"])
            self.assertIn("citation_requires_bibliography", [issue["kind"] for issue in report["errors"]])

    def test_xelatex_runs_bibtex_between_tex_passes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            (project / "references.bib").write_text("@book{key,title={Title}}", encoding="utf-8")
            def run_engine(command, **kwargs):
                if command[0] == "xelatex":
                    writer = PdfWriter()
                    writer.add_blank_page(width=200, height=200)
                    with (project / "build/main.pdf").open("wb") as destination:
                        writer.write(destination)
                return subprocess.CompletedProcess(command, 0, "", "")
            with patch("formatter_backends.latex.shutil.which", side_effect=lambda name: name if name in {"xelatex", "bibtex"} else None), patch("formatter_backends.latex.subprocess.run", side_effect=run_engine) as process:
                result = compile_project(project, engine_name="xelatex")
            self.assertEqual(result["status"], "ok")
            self.assertEqual([call.args[0][0] for call in process.call_args_list], ["xelatex", "bibtex", "xelatex"])
            self.assertEqual(process.call_args_list[1].kwargs["cwd"], str(project / "build"))

    def test_xelatex_does_not_publish_without_bibtex(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            (project / "references.bib").write_text("@book{key,title={Title}}", encoding="utf-8")
            with patch("formatter_backends.latex.shutil.which", side_effect=lambda name: "xelatex" if name == "xelatex" else None):
                result = compile_project(project, engine_name="xelatex")
            self.assertEqual(result["status"], "failed")
            self.assertIsNone(result["pdfPath"])

    def test_tectonic_bibtex_error_is_not_hidden_by_zero_exit(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            (project / "references.bib").write_text("@book{key,title={Title}}", encoding="utf-8")
            with patch("formatter_backends.latex.shutil.which", side_effect=lambda name: "tectonic" if name == "tectonic" else None), patch("formatter_backends.latex.subprocess.run", return_value=subprocess.CompletedProcess([], 0, "", "warning: errors were issued by BibTeX, but were ignored")):
                result = compile_project(project, engine_name="tectonic")
            self.assertEqual(result["status"], "failed")
            self.assertIsNone(result["pdfPath"])


class WordReferenceTests(unittest.TestCase):
    def test_markdown_references_are_extracted_without_acknowledgements(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.md"
            source.write_text("# 论文\n\n# 1 引言\n\n正文。\n\n# 参考文献\n\n陈守常. 油茶病害[M]. 北京：出版社，1989.\n\n# 致谢\n\n感谢指导教师。", encoding="utf-8")
            ir = build_thesis_ir(extract_text_evidence(source))
            report = check_reference_format(ir)
            self.assertEqual(report["entryCount"], 1)
            self.assertNotIn("感谢指导教师", str(ir["references"]["entries"]))

    def test_reference_text_unchanged_and_song_five_point(self) -> None:
        document = Document()
        text = "陈守常，曾大鹏. 油茶病害及其防治[M]. 北京：中国林业出版社，1989."
        paragraph = add_reference_paragraph_block(document, text, {})
        self.assertEqual(paragraph.text, text)
        for run in paragraph.runs:
            self.assertEqual(run.font.size.pt, 10.5)
            self.assertEqual(run.font.name, "宋体")
            self.assertEqual(run._element.rPr.rFonts.get(qn("w:eastAsia")), "宋体")

    def test_reference_checker_is_non_mutating_and_detects_rules(self) -> None:
        ir = {"references": {"entries": [{"text": "Smith A. Title[J]. Journal, 2020."}, {"text": "陈守常. 油茶病害[M]. 北京：出版社，1989."}]}}
        original = copy.deepcopy(ir)
        result = check_reference_format(ir)
        self.assertEqual(ir, original)
        self.assertFalse(result["citationManagement"])
        self.assertFalse(result["automaticTextChanges"])
        kinds = {issue["kind"] for issue in result["issues"]}
        self.assertIn("foreign_surname_not_uppercase", kinds)
        self.assertIn("chinese_entry_after_foreign_group", kinds)


if __name__ == "__main__":
    unittest.main()
