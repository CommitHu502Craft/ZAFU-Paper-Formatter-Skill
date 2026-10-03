from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pypdf import PdfWriter


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from formatter_core.checks import CheckPolicy
from formatter_core.workspace import Workspace, safe_name
from formatter_backends.latex import build_project, compile_project, escape_text, render_inline, run_latex, safe_math
from extract_source_evidence import extract_text_evidence
from thesis_ir import build_thesis_ir
from thesis_format import PROFILES, docx_commands, text_commands
from thesis_format import finalize_outputs, run_visual_refinement_pass
import render_validate_docx as rendering
from generate_review_report import _rel


class WorkspaceTests(unittest.TestCase):
    def test_names_are_safe_and_preserve_unicode(self) -> None:
        self.assertEqual(safe_name("CON"), "thesis-CON")
        self.assertEqual(safe_name(" ../论文:最终? "), "-论文-最终-")
        self.assertEqual(safe_name("..."), "thesis")
        self.assertNotIn("/", safe_name("a/b"))
        self.assertLessEqual(len(safe_name("中" * 100)), 64)

    def test_runs_are_isolated_and_artifacts_are_relative(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "论文.md"
            source.write_text("原文", encoding="utf-8")
            first = Workspace.create(source, "zafu_2022")
            second = Workspace.create(source, "zafu_2022")
            self.assertNotEqual(first.root, second.root)
            self.assertEqual(first.root.parent.parent, source.parent / "thesis-output")
            candidate = first.work / "candidate.docx"
            candidate.write_bytes(b"example")
            destination = first.publish(candidate, "word", "docx")
            self.assertFalse(candidate.exists())
            self.assertEqual(destination.name, "论文__zafu_2022__word.docx")
            first.write_manifest({"status": "completed", "artifacts": {"docx": str(destination)}})
            manifest = json.loads(first.manifest_path.read_text(encoding="utf-8"))
            self.assertFalse(Path(manifest["artifacts"]["docx"]).is_absolute())
            self.assertEqual(manifest["inputSha256"], hashlib.sha256(source.read_bytes()).hexdigest())
            self.assertFalse(manifest["visualReviewed"])
            self.assertEqual(source.read_text(encoding="utf-8"), "原文")

    def test_move_rejects_workspace_escape_and_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("original", encoding="utf-8")
            workspace = Workspace.create(source, "profile")
            with self.assertRaises(ValueError):
                workspace.publish(source, "word", "docx")
            candidate = workspace.work / "candidate"
            candidate.write_text("candidate", encoding="utf-8")
            with self.assertRaises(ValueError):
                workspace.move(candidate, source)
            target = workspace.deliverables / "existing"
            target.write_text("existing", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                workspace.move(candidate, target)
            self.assertTrue(candidate.exists())

    def test_relocation_rewrites_report_references(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("source", encoding="utf-8")
            workspace = Workspace.create(source, "profile")
            old_path = workspace.work / "old.pdf"
            new_path = workspace.deliverables / "new.pdf"
            report_path = workspace.reports / "report.json"
            report_path.write_text(json.dumps({"nested": [{"pdf": str(old_path)}]}), encoding="utf-8")
            workspace.rewrite_references({str(old_path): str(new_path)})
            self.assertEqual(json.loads(report_path.read_text(encoding="utf-8"))["nested"][0]["pdf"], str(new_path))

    def test_relocation_preserves_lone_pdf_surrogates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("source", encoding="utf-8")
            workspace = Workspace.create(source, "profile")
            report_path = workspace.reports / "report.json"
            report_path.write_text(json.dumps({"text": "glyph\ud800", "pdf": "old"}), encoding="utf-8")
            workspace.rewrite_references({"old": "new"})
            rewritten = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(rewritten["text"], "glyph\ud800")
            self.assertEqual(rewritten["pdf"], "new")

    def test_html_links_are_relative_to_the_report_directory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.assertEqual(_rel(str(root / "deliverables/paper.pdf"), root / "reports"), "../deliverables/paper.pdf")

    def test_hard_failed_word_candidate_is_not_published(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("source", encoding="utf-8")
            workspace = Workspace.create(source, "zafu_2022")
            candidate = workspace.work / "repaired.docx"
            candidate.write_bytes(b"candidate")
            manifest = {"runDir": str(workspace.root), "projectName": workspace.project, "profile": workspace.profile, "input": str(source), "backend": "word", "qualityGate": {"passed": False}}
            with patch("thesis_format.run_step"):
                outputs = finalize_outputs(workspace.work, source, workspace, manifest)
            self.assertNotIn("docx", outputs)
            self.assertTrue(candidate.is_file())
            self.assertEqual(list(workspace.deliverables.iterdir()), [])


class CheckPolicyTests(unittest.TestCase):
    def test_structural_is_the_default(self) -> None:
        self.assertFalse(CheckPolicy().render)
        self.assertFalse(CheckPolicy().summary()["visualReviewed"])
        self.assertTrue(CheckPolicy(export_pdf=True).render)
        self.assertTrue(CheckPolicy("layout").render)
        self.assertTrue(CheckPolicy("visual").render)
        with self.assertRaises(ValueError):
            CheckPolicy("invalid")

    def test_word_command_wires_explicit_check_policy(self) -> None:
        commands = docx_commands(Path("input.docx"), PROFILES["zafu_2022"], "conservative-repair", Path("work"), CheckPolicy("layout", True))
        validation = commands[-1]
        self.assertEqual(validation[validation.index("--check") + 1], "layout")
        self.assertIn("--export-pdf", validation)
        text = text_commands(Path("input.md"), PROFILES["zafu_2022"], Path("work"))
        self.assertIn("--build-only", text[0])

    def test_pdf_only_never_extracts_text_or_images(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            pdf = output / "input.pdf"
            with patch.object(rendering, "count_pdf_pages", return_value=2), patch.object(rendering, "extract_page_texts") as extract, patch.object(rendering, "export_all_page_thumbnails") as thumbnails:
                report = rendering.analyze_pdf_document(pdf, output / "preview", analyze_layout=False)
            self.assertEqual(report["status"], "ok")
            self.assertFalse(report["layoutAnalyzed"])
            extract.assert_not_called()
            thumbnails.assert_not_called()

    def test_layout_never_generates_png(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            with patch.object(rendering, "count_pdf_pages", return_value=1), patch.object(rendering, "extract_page_texts", return_value=["正文"]), patch.object(rendering, "extract_page_text_boxes", return_value=[]), patch.object(rendering, "export_critical_pages") as critical, patch.object(rendering, "export_all_page_thumbnails") as thumbnails, patch.object(rendering, "build_contact_sheets") as sheets:
                report = rendering.analyze_pdf_document(output / "input.pdf", output / "preview")
            critical.assert_not_called()
            thumbnails.assert_not_called()
            sheets.assert_not_called()
            self.assertTrue(report["layoutAnalyzed"])
            self.assertEqual(report["contactSheets"], [])
            self.assertFalse(report["visualReviewed"])
            self.assertEqual(list(output.rglob("*.png")), [])

    def test_visual_previews_are_explicit_not_a_review(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            with patch.object(rendering, "count_pdf_pages", return_value=1), patch.object(rendering, "extract_page_texts", return_value=["正文"]), patch.object(rendering, "extract_page_text_boxes", return_value=[]), patch.object(rendering, "export_critical_pages", return_value=[]) as critical, patch.object(rendering, "export_all_page_thumbnails", return_value=[]) as thumbnails:
                report = rendering.analyze_pdf_document(output / "input.pdf", output / "preview", export_previews=True)
            critical.assert_called_once()
            thumbnails.assert_called_once()
            self.assertFalse(report["visualReviewed"])

    def test_refinement_executor_does_not_run_a_separate_render(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            work = Path(directory)
            (work / "repaired.docx").write_bytes(b"candidate")
            plan = work / "plan.json"
            plan.write_text('{"actions": []}', encoding="utf-8")
            with patch("thesis_format.subprocess.run", return_value=subprocess.CompletedProcess([], 0)) as run:
                result = run_visual_refinement_pass(work, plan)
            self.assertEqual(result["status"], "ok")
            self.assertEqual(run.call_count, 1)
            self.assertIn("scripts/apply_visual_refinements.py", run.call_args.args[0])


class LatexTests(unittest.TestCase):
    def make_ir(self, source: Path) -> dict:
        return build_thesis_ir(extract_text_evidence(source))

    def test_text_escaping_and_math_safety(self) -> None:
        self.assertEqual(escape_text("a&b_10%"), r"a\&b\_10\%")
        self.assertTrue(safe_math(r"\frac{x_1}{2} + \alpha"))
        self.assertFalse(safe_math(r"\input{secret}"))
        self.assertFalse(safe_math(r"\begin{document}x\end{document}"))
        self.assertFalse(safe_math(r"\csname input\endcsname"))
        self.assertFalse(safe_math("{unbalanced"))
        issues = []
        rendered = render_inline(r"**重点** and $\alpha_1$ and $\input{secret}$", issues, "block")
        self.assertIn(r"\textbf{重点}", rendered)
        self.assertIn(r"\(\alpha_1\)", rendered)
        self.assertIn(r"\textbackslash{}input", rendered)
        self.assertEqual(issues[0]["kind"], "unsafe_or_unsupported_math")

    def test_inline_images_are_reported_instead_of_silently_omitted(self) -> None:
        issues = []
        rendered = render_inline("正文 ![图](figure.png) 之后", issues, "block")
        self.assertIn("figure.png", rendered)
        self.assertEqual(issues[0]["kind"], "inline_image_requires_own_block")
        literal_issues = []
        render_inline("`![图](figure.png)`", literal_issues, "block")
        self.assertEqual(literal_issues, [])

    def test_build_project_preserves_labels_tables_assets_and_source(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "论文.md"
            Image.new("RGB", (8, 8), "white").save(source.parent / "figure.png")
            source.write_text("# 测试论文\n\n# 摘要\n\n本文摘要。\n\n关键词：测试\n\n# 3 引言\n\n正文 $x_1$。\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n![示意图](figure.png)\n\n# 参考文献\n\n[1] Author. Title.\n", encoding="utf-8")
            original = source.read_bytes()
            report = build_project(self.make_ir(source), source, {"page": {"margin_top_cm": 3.1}}, ROOT / "profiles/zafu_2022", source.parent / "project")
            tex = (source.parent / "project/main.tex").read_text(encoding="utf-8")
            self.assertTrue(report["passed"], report["errors"])
            self.assertIn(r"\section*{3 引言}", tex)
            self.assertIn(r"\(x_1\)", tex)
            self.assertIn(r"\toprule", tex)
            self.assertIn("[1] Author. Title.", tex)
            self.assertIn(r"\documentclass{ZafuThesis}", tex)
            self.assertNotIn("top=3.1cm", tex)
            self.assertFalse(report["externalClass"]["modified"])
            self.assertTrue((source.parent / "project/ZafuThesis.cls").is_file())
            self.assertEqual(len(report["assets"]), 1)
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(report["profileCompliance"], "draft")

    def test_unsafe_image_and_unknown_blocks_are_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.md"
            source.write_text("# 论文\n\n# 1 引言\n\n![外部](../secret.png)\n", encoding="utf-8")
            ir = self.make_ir(source)
            ir["semanticBlocks"].append({"id": "unsupported", "kind": "ole", "text": "object"})
            report = build_project(ir, source, {}, ROOT / "profiles/zafu_2022", source.parent / "project")
            self.assertFalse(report["passed"])
            self.assertIn("unsupported_or_missing_image", [issue["kind"] for issue in report["errors"]])
            self.assertIn("unsupported_block_kind", [issue["kind"] for issue in report["errors"]])

    def test_docx_ir_is_not_silently_converted(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError):
                build_project({"source": {"format": "docx"}}, Path(directory) / "input.docx", {}, ROOT / "profiles/zafu_2022", Path(directory) / "project")

    def test_missing_engine_and_explicit_no_compile(self) -> None:
        with tempfile.TemporaryDirectory() as directory, patch("formatter_backends.latex.shutil.which", return_value=None) as detect:
            project = Path(directory)
            self.assertEqual(compile_project(project)["status"], "unavailable")
            detect.reset_mock()
            self.assertEqual(compile_project(project, no_compile=True)["status"], "skipped")
            detect.assert_not_called()

    def test_successful_compile_requires_pdf_and_disables_shell_escape(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            def fake_engine(command: list[str], **kwargs: object) -> subprocess.CompletedProcess:
                writer = PdfWriter()
                writer.add_blank_page(width=595, height=842)
                writer.write(project / "build/main.pdf")
                return subprocess.CompletedProcess(command, 0, "ok", "")

            with patch("formatter_backends.latex.shutil.which", return_value="xelatex"), patch("formatter_backends.latex.subprocess.run", side_effect=fake_engine) as run:
                result = compile_project(project, timeout=7)
            self.assertEqual(result["status"], "ok")
            self.assertEqual(result["pageCount"], 1)
            self.assertEqual(run.call_count, 2)
            self.assertIn("-no-shell-escape", run.call_args.args[0])
            self.assertEqual(run.call_args.kwargs["timeout"], 7)

    def test_failed_compile_never_reuses_existing_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            (project / "build").mkdir()
            (project / "build/main.pdf").write_bytes(b"stale")
            with patch("formatter_backends.latex.shutil.which", return_value="xelatex"), patch("formatter_backends.latex.subprocess.run", return_value=subprocess.CompletedProcess([], 1, "error", "")):
                result = compile_project(project)
            self.assertEqual(result["status"], "failed")
            self.assertIsNone(result["pdfPath"])

    def test_tectonic_uses_untrusted_mode(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            def fake_engine(command: list[str], **kwargs: object) -> subprocess.CompletedProcess:
                writer = PdfWriter()
                writer.add_blank_page(width=595, height=842)
                writer.write(project / "build/main.pdf")
                return subprocess.CompletedProcess(command, 0, "ok", "")

            with patch("formatter_backends.latex.shutil.which", side_effect=lambda name: "tectonic" if name == "tectonic" else None), patch("formatter_backends.latex.subprocess.run", side_effect=fake_engine) as run:
                result = compile_project(project)
            self.assertEqual(result["status"], "ok")
            self.assertEqual(result["engineName"], "tectonic")
            self.assertIn("--untrusted", run.call_args.args[0])
            self.assertEqual(run.call_count, 1)

    def test_compile_timeout_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory, patch("formatter_backends.latex.shutil.which", return_value="xelatex"), patch("formatter_backends.latex.subprocess.run", side_effect=subprocess.TimeoutExpired("xelatex", 1)):
            result = compile_project(Path(directory), timeout=1)
            self.assertEqual(result["status"], "failed")
            self.assertIn("timed out", result["reason"])

    def test_source_zip_excludes_build_files_and_reports_no_compile(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("论文\n摘要\n摘要内容。\n关键词：测试\n1 引言\n正文。", encoding="utf-8")
            workspace = Workspace.create(source, "zafu_2022")
            result = run_latex(self.make_ir(source), source, {}, ROOT / "profiles/zafu_2022", workspace, CheckPolicy(), no_compile=True)
            self.assertEqual(result["compilation"]["status"], "skipped")
            self.assertNotIn("latexPdf", result["deliverables"])
            with zipfile.ZipFile(result["deliverables"]["latexSource"]) as archive:
                self.assertIn("main.tex", archive.namelist())
                self.assertIn("ZafuThesis.cls", archive.namelist())
                self.assertIn("UPSTREAM.json", archive.namelist())
                provenance = json.loads(archive.read("UPSTREAM.json"))
                self.assertEqual(provenance["sha256"], hashlib.sha256(archive.read("ZafuThesis.cls")).hexdigest())
                self.assertFalse(any(name.startswith("build/") for name in archive.namelist()))
                self.assertEqual(json.loads(archive.read("conversion.json"))["mainTex"], "main.tex")

    def test_published_pdf_updates_latex_layout_references(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("论文\n摘要\n摘要内容。\n1 引言\n正文。", encoding="utf-8")
            workspace = Workspace.create(source, "zafu_2022")
            def fake_compile(project: Path, *arguments: object) -> dict:
                build = project / "build"
                build.mkdir()
                writer = PdfWriter()
                writer.add_blank_page(width=595, height=842)
                writer.write(build / "main.pdf")
                return {"status": "ok", "pdfPath": str(build / "main.pdf")}

            def fake_layout(pdf: Path, output: Path, **arguments: object) -> dict:
                output.mkdir()
                manifest_path = output / "layout.json"
                manifest_path.write_text(json.dumps({"pdfPath": str(pdf)}), encoding="utf-8")
                return {"pdfPath": str(pdf), "manifestPath": str(manifest_path)}

            with patch("formatter_backends.latex.compile_project", side_effect=fake_compile), patch("render_validate_docx.analyze_pdf_document", side_effect=fake_layout):
                result = run_latex(self.make_ir(source), source, {}, ROOT / "profiles/zafu_2022", workspace, CheckPolicy("layout"))
            published = result["deliverables"]["latexPdf"]
            self.assertEqual(result["layoutValidation"]["pdfPath"], published)
            layout = json.loads((workspace.work / "latex-preview/layout.json").read_text(encoding="utf-8"))
            self.assertEqual(layout["pdfPath"], published)
            self.assertFalse((workspace.work / "latex/build/main.pdf").exists())


class CliAcceptanceTests(unittest.TestCase):
    def run_cli(self, source: Path, output: Path, *arguments: str) -> tuple[subprocess.CompletedProcess, dict, Path]:
        command = [sys.executable, str(ROOT / "scripts/thesis_format.py"), str(source), "--output-dir", str(output), *arguments]
        result = subprocess.run(command, cwd=ROOT, text=True, encoding="utf-8", errors="replace", capture_output=True, timeout=90)
        manifests = list(output.glob("*/*/manifest.json"))
        self.assertTrue(manifests, result.stdout + result.stderr)
        path = max(manifests, key=lambda candidate: candidate.stat().st_mtime_ns)
        return result, json.loads(path.read_text(encoding="utf-8")), path

    def test_both_routes_share_ir_and_default_has_no_pdf_or_png(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "论文.md"
            source.write_text("# 测试论文\n\n# 摘要\n\n本文摘要。\n\n关键词：测试\n\n# 1 引言\n\n正文。\n\n# 参考文献\n\n[1] Example.\n", encoding="utf-8")
            original = source.read_bytes()
            result, manifest, path = self.run_cli(source, Path(directory) / "output", "--backend", "both", "--no-compile")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(manifest["status"], "completed_with_warnings")
            root = path.parent
            self.assertEqual(len(list(root.rglob("thesis_ir.json"))), 1)
            self.assertEqual(sum("scripts/thesis_ir.py" in command for command in manifest["commands"]), 1)
            self.assertEqual(list(root.rglob("*.pdf")), [])
            self.assertEqual(list(root.rglob("*.png")), [])
            self.assertTrue((root / manifest["deliverables"]["docx"]).is_file())
            self.assertTrue((root / manifest["deliverables"]["latexSource"]).is_file())
            self.assertFalse((root / "work/repaired.docx").exists())
            self.assertFalse((root / "final").exists())
            self.assertFalse((root / "debug").exists())
            self.assertEqual(source.read_bytes(), original)
            summary = json.loads((root / manifest["artifacts"]["summary"]).read_text(encoding="utf-8"))
            self.assertTrue(summary["inputPreserved"])
            self.assertFalse(summary["visualReviewed"])
            for value in manifest["artifacts"].values():
                self.assertFalse(Path(value).is_absolute(), value)

    def test_latex_only_does_not_build_word(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("测试论文\n摘要\n本文摘要。\n关键词：测试\n1 引言\n正文。", encoding="utf-8")
            result, manifest, path = self.run_cli(source, Path(directory) / "output", "--backend", "latex", "--no-compile")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertEqual(manifest["status"], "completed_source_only")
            self.assertEqual(list(path.parent.rglob("*.docx")), [])
            self.assertEqual(manifest["backendResults"]["latex"]["compilation"]["status"], "skipped")

    def test_bad_input_leaves_failed_manifest_without_deliverables(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "broken.docx"
            source.write_bytes(b"not a docx")
            result, manifest, path = self.run_cli(source, Path(directory) / "output")
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(manifest["status"], "failed")
            self.assertEqual(list((path.parent / "deliverables").iterdir()), [])

    def test_unsafe_formula_reports_hard_failure_without_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.md"
            source.write_text("# 论文\n\n# 1 引言\n\n$\\input{secret}$", encoding="utf-8")
            result, manifest, path = self.run_cli(source, Path(directory) / "output", "--backend", "latex")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertEqual(manifest["status"], "completed_with_hard_failures")
            self.assertFalse(manifest["qualityGate"]["passed"])
            self.assertTrue((path.parent / manifest["deliverables"]["latexSource"]).is_file())
            self.assertNotIn("latexPdf", manifest["deliverables"])

    def test_dry_run_isolated_and_does_not_publish(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.txt"
            source.write_text("source", encoding="utf-8")
            output = Path(directory) / "output"
            first_result, first, first_path = self.run_cli(source, output, "--dry-run")
            second_result, second, second_path = self.run_cli(source, output, "--dry-run")
            self.assertEqual(first_result.returncode, 0)
            self.assertEqual(second_result.returncode, 0)
            self.assertEqual(first["status"], "dry-run")
            self.assertNotEqual(first_path, second_path)
            self.assertNotEqual(first["runId"], second["runId"])
            self.assertEqual(list((first_path.parent / "deliverables").iterdir()), [])


if __name__ == "__main__":
    unittest.main()
