#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Tuple


ROOT = Path(__file__).resolve().parents[1]
PYTHON = sys.executable
W_NS = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
DOCX_W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scaffold regression fixture runner for thesis formatter.")
    parser.add_argument("--profile", default="zafu_2022", help="Profile name")
    parser.add_argument("--fixtures-dir", default="tests/fixtures", help="Fixture root directory")
    parser.add_argument("--output-dir", default="test_output", help="Output directory for fixture reports")
    parser.add_argument("--include-tag", action="append", default=[], help="Only run fixtures containing this tag; repeatable")
    return parser.parse_args()


def discover_fixtures(fixtures_dir: Path) -> List[Dict[str, object]]:
    fixtures: List[Dict[str, object]] = []
    if not fixtures_dir.exists():
        return fixtures
    for child in sorted(fixtures_dir.iterdir()):
        if not child.is_dir():
            continue
        fixtures.append(
            {
                "name": child.name,
                "path": str(child),
                "expectedSpec": str(child / "expected.json"),
                "inputCandidates": [str(p) for p in sorted(child.glob("*")) if p.is_file()],
            }
        )
    return fixtures


def read_json(path: Path) -> Dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def normalize_tags(value: object) -> List[str]:
    if not isinstance(value, list):
        return []
    return [str(item) for item in value if isinstance(item, str) and item.strip()]


def normalize_required_commands(value: object) -> List[str]:
    if not isinstance(value, list):
        return []
    return [str(item) for item in value if isinstance(item, str) and item.strip()]


def detect_available_commands(commands: List[str]) -> Dict[str, Optional[str]]:
    detected: Dict[str, Optional[str]] = {}
    for command in commands:
        detected[command] = shutil.which(command)
    return detected


def should_run_fixture(expected_spec: Dict[str, object], include_tags: List[str]) -> Tuple[bool, Optional[str], Dict[str, Optional[str]]]:
    fixture_tags = normalize_tags(expected_spec.get("tags"))
    if include_tags and not set(include_tags).issubset(set(fixture_tags)):
        return False, "tag_filtered_out", {}

    required_commands = normalize_required_commands(expected_spec.get("requiresCommands"))
    if not required_commands:
        return True, None, {}

    detected = detect_available_commands(required_commands)
    missing = [command for command, resolved in detected.items() if not resolved]
    if missing:
        return False, f"missing_commands:{','.join(missing)}", detected
    return True, None, detected


def assert_equal(mismatches: List[str], field: str, actual: object, expected_value: object) -> None:
    if actual != expected_value:
        mismatches.append(f"{field}: expected {expected_value!r}, got {actual!r}")


def assert_contains_keys(mismatches: List[str], field: str, payload: object, required_keys: List[str]) -> None:
    if not isinstance(payload, dict):
        mismatches.append(f"{field}: expected dict, got {type(payload).__name__}")
        return
    for key in required_keys:
        if key not in payload:
            mismatches.append(f"{field}: missing key {key!r}")


def assert_dict_contains_values(mismatches: List[str], field: str, payload: object, expected_values: Dict[str, object]) -> None:
    if not isinstance(payload, dict):
        mismatches.append(f"{field}: expected dict, got {type(payload).__name__}")
        return
    for key, expected_value in expected_values.items():
        actual_value = payload.get(key)
        if actual_value != expected_value:
            mismatches.append(f"{field}.{key}: expected {expected_value!r}, got {actual_value!r}")


def load_optional_json(path_value: object) -> Tuple[Optional[Path], Optional[Dict[str, object]]]:
    if not isinstance(path_value, str) or not path_value.strip():
        return None, None
    path = Path(path_value).resolve()
    if not path.exists():
        return path, None
    return path, read_json(path)


def load_optional_path(path_value: object) -> Optional[Path]:
    if not isinstance(path_value, str) or not path_value.strip():
        return None
    return Path(path_value).resolve()


def count_docx_omml_nodes(docx_path: Path) -> int:
    if not docx_path.exists():
        return 0
    with zipfile.ZipFile(docx_path) as zf:
        try:
            document_xml = zf.read("word/document.xml")
        except KeyError:
            return 0
    root = ET.fromstring(document_xml)
    return len(root.findall(".//m:oMath", W_NS)) + len(root.findall(".//m:oMathPara", W_NS))


def read_docx_document_root(docx_path: Path) -> Optional[ET.Element]:
    if not docx_path.exists():
        return None
    with zipfile.ZipFile(docx_path) as zf:
        try:
            document_xml = zf.read("word/document.xml")
        except KeyError:
            return None
    return ET.fromstring(document_xml)


def count_docx_dollar_blocks(docx_path: Path) -> int:
    root = read_docx_document_root(docx_path)
    if root is None:
        return 0
    text = "".join(node.text or "" for node in root.findall(".//w:t", DOCX_W_NS))
    return text.count("$$")


def docx_paragraph_texts(docx_path: Path) -> List[str]:
    root = read_docx_document_root(docx_path)
    if root is None:
        return []
    return [
        "".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_W_NS)).strip()
        for paragraph in root.findall(".//w:body/w:p", DOCX_W_NS)
    ]


def docx_table_bottom_border_row_counts(docx_path: Path) -> List[List[int]]:
    root = read_docx_document_root(docx_path)
    if root is None:
        return []
    tables = root.findall(".//w:tbl", DOCX_W_NS)
    result: List[List[int]] = []
    for table in tables:
        row_counts: List[int] = []
        for row in table.findall("w:tr", DOCX_W_NS):
            count = 0
            for tcpr in row.findall("w:tc/w:tcPr", DOCX_W_NS):
                tc_borders = tcpr.find("w:tcBorders", DOCX_W_NS)
                if tc_borders is not None and tc_borders.find("w:bottom", DOCX_W_NS) is not None:
                    count += 1
            row_counts.append(count)
        result.append(row_counts)
    return result


def check_result_artifacts(
    manifest: Dict[str, object],
    expected: Dict[str, object],
    mismatches: List[str],
) -> Dict[str, object]:
    artifacts = manifest.get("artifacts") or {}
    if not isinstance(artifacts, dict):
        artifacts = {}

    repair_plan_path, repair_plan = load_optional_json(artifacts.get("repairPlan"))
    repair_execution_path, repair_execution = load_optional_json(artifacts.get("repairExecution"))
    repaired_docx_path = load_optional_path(artifacts.get("repairedDocx"))
    repaired_pdf_path = load_optional_path(artifacts.get("repairedPdf"))

    requires_repair_plan = expected.get("requiresRepairPlan")
    if requires_repair_plan is not None:
        assert_equal(
            mismatches,
            "requiresRepairPlan",
            repair_plan_path is not None and repair_plan is not None,
            bool(requires_repair_plan),
        )
    elif repair_plan_path is not None and repair_plan is None:
        mismatches.append(f"repairPlan: declared but file not found at {repair_plan_path}")

    if repair_plan is not None:
        assert_contains_keys(
            mismatches,
            "repairPlan",
            repair_plan,
            ["actions", "schemaGate", "blockedAutoRepairs", "manualReview", "confirmationRequests"],
        )

    requires_repair_execution = expected.get("requiresRepairExecution")
    if requires_repair_execution is not None:
        assert_equal(
            mismatches,
            "requiresRepairExecution",
            repair_execution_path is not None and repair_execution is not None,
            bool(requires_repair_execution),
        )
    elif repair_execution_path is not None and repair_execution is None:
        mismatches.append(f"repairExecution: declared but file not found at {repair_execution_path}")

    if repair_execution is not None:
        expected_repair_execution = expected.get("repairExecution")
        if isinstance(expected_repair_execution, dict):
            for key, expected_value in expected_repair_execution.items():
                assert_equal(mismatches, f"repairExecution.{key}", repair_execution.get(key), expected_value)
        if "repairExecutionLogActionIncludes" in expected:
            execution_log = repair_execution.get("executionLog") or []
            action_names = {item.get("action") for item in execution_log if isinstance(item, dict)}
            for action_name in expected.get("repairExecutionLogActionIncludes") or []:
                if action_name not in action_names:
                    mismatches.append(f"repairExecution.executionLog missing expected action {action_name!r}")

    requires_repaired_docx = expected.get("requiresRepairedDocx")
    if requires_repaired_docx is not None:
        has_repaired_docx = repaired_docx_path is not None and repaired_docx_path.exists()
        assert_equal(mismatches, "requiresRepairedDocx", has_repaired_docx, bool(requires_repaired_docx))
    elif repaired_docx_path is not None and not repaired_docx_path.exists():
        mismatches.append(f"repairedDocx: declared but file not found at {repaired_docx_path}")
    if repaired_docx_path is not None and repaired_docx_path.exists() and "repairedDocxMinOmmlCount" in expected:
        omml_count = count_docx_omml_nodes(repaired_docx_path)
        if omml_count < int(expected.get("repairedDocxMinOmmlCount") or 0):
            mismatches.append(
                f"repairedDocxMinOmmlCount: expected >= {expected.get('repairedDocxMinOmmlCount')!r}, got {omml_count!r}"
            )
    if repaired_docx_path is not None and repaired_docx_path.exists() and "repairedDocxDollarCountMax" in expected:
        dollar_count = count_docx_dollar_blocks(repaired_docx_path)
        maximum = int(expected.get("repairedDocxDollarCountMax") or 0)
        if dollar_count > maximum:
            mismatches.append(f"repairedDocxDollarCountMax: expected <= {maximum!r}, got {dollar_count!r}")
    if repaired_docx_path is not None and repaired_docx_path.exists() and "repairedDocxParagraphTextsInclude" in expected:
        paragraph_texts = docx_paragraph_texts(repaired_docx_path)
        merged = "\n".join(paragraph_texts)
        for snippet in expected.get("repairedDocxParagraphTextsInclude") or []:
            if snippet not in merged:
                mismatches.append(f"repairedDocxParagraphTexts missing expected snippet {snippet!r}")
    if repaired_docx_path is not None and repaired_docx_path.exists() and "repairedDocxTableBottomBorderRowCounts" in expected:
        actual_counts = docx_table_bottom_border_row_counts(repaired_docx_path)
        assert_equal(
            mismatches,
            "repairedDocxTableBottomBorderRowCounts",
            actual_counts,
            expected.get("repairedDocxTableBottomBorderRowCounts"),
        )

    requires_repaired_pdf = expected.get("requiresRepairedPdf")
    if requires_repaired_pdf is not None:
        has_repaired_pdf = repaired_pdf_path is not None and repaired_pdf_path.exists()
        assert_equal(mismatches, "requiresRepairedPdf", has_repaired_pdf, bool(requires_repaired_pdf))
    elif repaired_pdf_path is not None and not repaired_pdf_path.exists():
        mismatches.append(f"repairedPdf: declared but file not found at {repaired_pdf_path}")

    return {
        "repairPlan": str(repair_plan_path) if repair_plan_path is not None else None,
        "repairExecution": str(repair_execution_path) if repair_execution_path is not None else None,
        "repairedDocx": str(repaired_docx_path) if repaired_docx_path is not None else None,
        "repairedPdf": str(repaired_pdf_path) if repaired_pdf_path is not None else None,
    }


def check_validation_reports(
    manifest: Dict[str, object],
    expected: Dict[str, object],
    mismatches: List[str],
) -> Dict[str, object]:
    artifacts = manifest.get("artifacts") or {}
    if not isinstance(artifacts, dict):
        artifacts = {}

    validation_path, validation_report = load_optional_json(artifacts.get("validationReport"))
    render_path, render_report = load_optional_json(artifacts.get("renderValidationReport"))
    render_preview_dir = artifacts.get("renderPreviewDir")
    preview_manifest_path: Optional[Path] = None
    preview_manifest: Optional[Dict[str, object]] = None

    requires_validation_report = expected.get("requiresValidationReport")
    if requires_validation_report is not None:
        assert_equal(
            mismatches,
            "requiresValidationReport",
            validation_path is not None and validation_report is not None,
            bool(requires_validation_report),
        )
    elif validation_path is not None and validation_report is None:
        mismatches.append(f"validationReport: declared but file not found at {validation_path}")

    requires_render_report = expected.get("requiresRenderValidationReport")
    if requires_render_report is not None:
        assert_equal(
            mismatches,
            "requiresRenderValidationReport",
            render_path is not None and render_report is not None,
            bool(requires_render_report),
        )
    elif render_path is not None and render_report is None:
        mismatches.append(f"renderValidationReport: declared but file not found at {render_path}")

    if validation_report is not None:
        assert_contains_keys(
            mismatches,
            "validationReport",
            validation_report,
            ["structuralValidation", "ruleValidation", "renderValidation", "validationSummary", "passed"],
        )
        assert_contains_keys(
            mismatches,
            "validationReport.structuralValidation",
            validation_report.get("structuralValidation"),
            ["checks", "issues"],
        )
        assert_contains_keys(
            mismatches,
            "validationReport.ruleValidation",
            validation_report.get("ruleValidation"),
            ["checks", "issues"],
        )
        assert_contains_keys(
            mismatches,
            "validationReport.validationSummary",
            validation_report.get("validationSummary"),
            ["structuralPassed", "rulePassed", "renderPassed", "renderStatus"],
        )

        expected_validation_summary = expected.get("validationSummary")
        if isinstance(expected_validation_summary, dict):
            actual_summary = validation_report.get("validationSummary") or {}
            for key, expected_value in expected_validation_summary.items():
                assert_equal(
                    mismatches,
                    f"validationSummary.{key}",
                    actual_summary.get(key) if isinstance(actual_summary, dict) else None,
                    expected_value,
                )

    if render_report is not None:
        assert_contains_keys(
            mismatches,
            "renderValidationReport",
            render_report,
            ["status", "available", "pdfPath", "pageCount", "previewDir", "suggestedReviewPages", "previewExport"],
        )
        assert_contains_keys(
            mismatches,
            "renderValidationReport.suggestedReviewPages",
            render_report.get("suggestedReviewPages"),
            ["cover", "toc", "abstract", "firstBodyPage", "references"],
        )
        assert_contains_keys(
            mismatches,
            "renderValidationReport.previewExport",
            render_report.get("previewExport"),
            ["available", "backend", "backendPath", "manifestPath", "images", "reason"],
        )

        preview_export = render_report.get("previewExport")
        if isinstance(preview_export, dict):
            preview_manifest_path, preview_manifest = load_optional_json(preview_export.get("manifestPath"))
            if preview_manifest_path is not None and preview_manifest is None:
                mismatches.append(f"previewExport.manifestPath: declared but file not found at {preview_manifest_path}")
            if preview_manifest is not None:
                assert_contains_keys(mismatches, "previewManifest", preview_manifest, ["pages"])

        if validation_report is not None:
            embedded_render_report = validation_report.get("renderValidation")
            if isinstance(embedded_render_report, dict):
                assert_equal(
                    mismatches,
                    "renderValidation.status_consistency",
                    embedded_render_report.get("status"),
                    render_report.get("status"),
                )
                assert_equal(
                    mismatches,
                    "renderValidation.previewDir_consistency",
                    embedded_render_report.get("previewDir"),
                    render_report.get("previewDir"),
                )

        expected_render_validation = expected.get("renderValidation")
        if isinstance(expected_render_validation, dict):
            for key, expected_value in expected_render_validation.items():
                assert_equal(mismatches, f"renderValidation.{key}", render_report.get(key), expected_value)

        expected_preview_export = expected.get("previewExport")
        if isinstance(expected_preview_export, dict) and isinstance(preview_export, dict):
            for key, expected_value in expected_preview_export.items():
                assert_equal(mismatches, f"previewExport.{key}", preview_export.get(key), expected_value)

        expected_preview_image_count = expected.get("previewImageCount")
        if expected_preview_image_count is not None and isinstance(preview_export, dict):
            images = preview_export.get("images")
            actual_count = len(images) if isinstance(images, list) else None
            assert_equal(mismatches, "previewImageCount", actual_count, int(expected_preview_image_count))

        expected_preview_manifest = expected.get("previewManifest")
        if isinstance(expected_preview_manifest, dict) and preview_manifest is not None:
            for key, expected_value in expected_preview_manifest.items():
                actual_value = preview_manifest.get(key) if isinstance(preview_manifest, dict) else None
                assert_equal(mismatches, f"previewManifest.{key}", actual_value, expected_value)

        preview_pages = preview_manifest.get("pages") if isinstance(preview_manifest, dict) else None
        if preview_pages is not None and not isinstance(preview_pages, list):
            mismatches.append(f"previewManifest.pages: expected list, got {type(preview_pages).__name__}")
            preview_pages = None

        expected_preview_manifest_min_pages = expected.get("previewManifestMinPages")
        if expected_preview_manifest_min_pages is not None and isinstance(preview_pages, list):
            actual_count = len(preview_pages)
            if actual_count < int(expected_preview_manifest_min_pages):
                mismatches.append(
                    f"previewManifestMinPages: expected >= {int(expected_preview_manifest_min_pages)!r}, got {actual_count!r}"
                )

        expected_preview_manifest_page_count = expected.get("previewManifestPageCount")
        if expected_preview_manifest_page_count is not None and isinstance(preview_pages, list):
            assert_equal(mismatches, "previewManifestPageCount", len(preview_pages), int(expected_preview_manifest_page_count))

        expected_preview_page_labels = expected.get("previewManifestPageLabelsIncludes")
        if isinstance(expected_preview_page_labels, list) and isinstance(preview_pages, list):
            actual_labels = {item.get("label") for item in preview_pages if isinstance(item, dict)}
            for label in expected_preview_page_labels:
                if label not in actual_labels:
                    mismatches.append(f"previewManifestPageLabels missing expected label {label!r}")

        expected_preview_statuses = expected.get("previewManifestStatusesInclude")
        if isinstance(expected_preview_statuses, list) and isinstance(preview_pages, list):
            actual_statuses = {item.get("status") for item in preview_pages if isinstance(item, dict)}
            for status in expected_preview_statuses:
                if status not in actual_statuses:
                    mismatches.append(f"previewManifestStatuses missing expected status {status!r}")

    requires_preview_dir = expected.get("requiresRenderPreviewDir")
    if requires_preview_dir is not None:
        has_preview_dir = isinstance(render_preview_dir, str) and bool(render_preview_dir.strip())
        assert_equal(mismatches, "requiresRenderPreviewDir", has_preview_dir, bool(requires_preview_dir))

    requires_preview_manifest = expected.get("requiresPreviewManifest")
    if requires_preview_manifest is not None:
        has_preview_manifest = preview_manifest_path is not None and preview_manifest is not None
        assert_equal(mismatches, "requiresPreviewManifest", has_preview_manifest, bool(requires_preview_manifest))

    return {
        "validationReport": str(validation_path) if validation_path is not None else None,
        "renderValidationReport": str(render_path) if render_path is not None else None,
        "previewManifest": str(preview_manifest_path) if preview_manifest_path is not None else None,
    }


def check_strategy_and_ir(
    manifest: Dict[str, object],
    expected: Dict[str, object],
    mismatches: List[str],
) -> Dict[str, object]:
    artifacts = manifest.get("artifacts") or {}
    if not isinstance(artifacts, dict):
        artifacts = {}

    strategy_path, strategy_report = load_optional_json(artifacts.get("strategySelection"))
    thesis_ir_path, thesis_ir = load_optional_json(artifacts.get("thesisIr"))
    preflight_path, preflight_report = load_optional_json(artifacts.get("preflightReport"))
    citation_plan_path, citation_plan = load_optional_json(artifacts.get("citationConversionPlan"))
    hybrid_attachment_report_path, hybrid_attachment_report = load_optional_json(artifacts.get("hybridAttachmentReport"))
    hybrid_rebuild_report_path, hybrid_rebuild_report = load_optional_json(artifacts.get("hybridRebuildReport"))

    requires_strategy = expected.get("requiresStrategySelection")
    if requires_strategy is not None:
        assert_equal(
            mismatches,
            "requiresStrategySelection",
            strategy_path is not None and strategy_report is not None,
            bool(requires_strategy),
        )

    if strategy_report is not None:
        if "chosenStrategy" in expected:
            assert_equal(mismatches, "chosenStrategy", strategy_report.get("chosenStrategy"), expected.get("chosenStrategy"))
        if "strategyExecutionMode" in expected:
            assert_equal(
                mismatches,
                "strategyExecutionMode",
                strategy_report.get("executionMode"),
                expected.get("strategyExecutionMode"),
            )
        strategy_reasons = strategy_report.get("reasons") or []
        if "strategyReasonsInclude" in expected and isinstance(strategy_reasons, list):
            for item in expected.get("strategyReasonsInclude") or []:
                if item not in strategy_reasons:
                    mismatches.append(f"strategyReasons missing expected item {item!r}")

    requires_thesis_ir = expected.get("requiresThesisIr")
    if requires_thesis_ir is not None:
        assert_equal(
            mismatches,
            "requiresThesisIr",
            thesis_ir_path is not None and thesis_ir is not None,
            bool(requires_thesis_ir),
        )

    if thesis_ir is not None:
        if "thesisIrTopLevelKeysInclude" in expected and isinstance(expected.get("thesisIrTopLevelKeysInclude"), list):
            for key in expected.get("thesisIrTopLevelKeysInclude") or []:
                if key not in thesis_ir:
                    mismatches.append(f"thesisIr missing expected top-level key {key!r}")
        if "thesisIrStrategyCandidate" in expected:
            assert_equal(
                mismatches,
                "thesisIrStrategyCandidate",
                thesis_ir.get("strategyCandidate"),
                expected.get("thesisIrStrategyCandidate"),
            )
        confidence = thesis_ir.get("confidence") or {}
        if "minOverallStructureConfidence" in expected and isinstance(confidence, dict):
            actual = float(confidence.get("overall") or 0.0)
            minimum = float(expected.get("minOverallStructureConfidence") or 0.0)
            if actual < minimum:
                mismatches.append(f"minOverallStructureConfidence: expected >= {minimum!r}, got {actual!r}")
        references = thesis_ir.get("references") or {}
        normalization = references.get("normalization") or {}
        if "referenceNormalization" in expected and isinstance(expected.get("referenceNormalization"), dict):
            for key, expected_value in (expected.get("referenceNormalization") or {}).items():
                assert_equal(
                    mismatches,
                    f"referenceNormalization.{key}",
                    normalization.get(key) if isinstance(normalization, dict) else None,
                    expected_value,
                )
        acknowledgements = thesis_ir.get("acknowledgements") or {}
        appendix = thesis_ir.get("appendix") or {}
        caption_blocks = thesis_ir.get("captionBlocks") or []
        attachable_asset_candidates = thesis_ir.get("attachableAssetCandidates") or []
        attachable_asset_candidate_summary = thesis_ir.get("attachableAssetCandidateSummary") or {}
        asset_anchor_ambiguities = thesis_ir.get("assetAnchorAmbiguities") or []
        if "thesisIrAcknowledgementsMinBlocks" in expected:
            actual = len(acknowledgements.get("blocks") or []) if isinstance(acknowledgements, dict) else 0
            minimum = int(expected.get("thesisIrAcknowledgementsMinBlocks") or 0)
            if actual < minimum:
                mismatches.append(f"thesisIrAcknowledgementsMinBlocks: expected >= {minimum!r}, got {actual!r}")
        if "thesisIrAppendixMinSections" in expected:
            actual = len(appendix.get("sections") or []) if isinstance(appendix, dict) else 0
            minimum = int(expected.get("thesisIrAppendixMinSections") or 0)
            if actual < minimum:
                mismatches.append(f"thesisIrAppendixMinSections: expected >= {minimum!r}, got {actual!r}")
        if "thesisIrCaptionBlockCount" in expected:
            actual = len(caption_blocks) if isinstance(caption_blocks, list) else 0
            assert_equal(mismatches, "thesisIrCaptionBlockCount", actual, int(expected.get("thesisIrCaptionBlockCount") or 0))
        if "thesisIrAttachableAssetCandidateCount" in expected:
            actual = len(attachable_asset_candidates) if isinstance(attachable_asset_candidates, list) else 0
            assert_equal(
                mismatches,
                "thesisIrAttachableAssetCandidateCount",
                actual,
                int(expected.get("thesisIrAttachableAssetCandidateCount") or 0),
            )
        if "thesisIrAttachableAssetActionCounts" in expected and isinstance(expected.get("thesisIrAttachableAssetActionCounts"), dict):
            actual_counts = attachable_asset_candidate_summary.get("actionCounts") if isinstance(attachable_asset_candidate_summary, dict) else None
            assert_dict_contains_values(
                mismatches,
                "thesisIrAttachableAssetActionCounts",
                actual_counts,
                expected.get("thesisIrAttachableAssetActionCounts") or {},
            )
        if "assetAnchorAmbiguitiesMinCount" in expected:
            actual = len(asset_anchor_ambiguities) if isinstance(asset_anchor_ambiguities, list) else 0
            minimum = int(expected.get("assetAnchorAmbiguitiesMinCount") or 0)
            if actual < minimum:
                mismatches.append(f"assetAnchorAmbiguitiesMinCount: expected >= {minimum!r}, got {actual!r}")
        if "assetAnchorAmbiguityKindsInclude" in expected:
            actual_kinds = {item.get("kind") for item in asset_anchor_ambiguities if isinstance(item, dict)}
            for kind in expected.get("assetAnchorAmbiguityKindsInclude") or []:
                if kind not in actual_kinds:
                    mismatches.append(f"assetAnchorAmbiguityKinds missing expected item {kind!r}")
        if "thesisIrAmbiguityKindsInclude" in expected:
            actual_kinds = {item.get("kind") for item in (thesis_ir.get("ambiguities") or []) if isinstance(item, dict)}
            for kind in expected.get("thesisIrAmbiguityKindsInclude") or []:
                if kind not in actual_kinds:
                    mismatches.append(f"thesisIr ambiguity kinds missing expected item {kind!r}")

    if preflight_report is not None:
        if "preflightBlockedAutoRepairsIncludes" in expected:
            blocked = list(preflight_report.get("blockedAutoRepairs") or [])
            for item in expected.get("preflightBlockedAutoRepairsIncludes") or []:
                if item not in blocked:
                    mismatches.append(f"preflightBlockedAutoRepairs missing expected item {item!r}")

    requires_citation_plan = expected.get("requiresCitationConversionPlan")
    if requires_citation_plan is not None:
        assert_equal(
            mismatches,
            "requiresCitationConversionPlan",
            citation_plan_path is not None and citation_plan is not None,
            bool(requires_citation_plan),
        )
    if citation_plan is not None and isinstance(expected.get("citationConversionPlan"), dict):
        for key, expected_value in (expected.get("citationConversionPlan") or {}).items():
            assert_equal(mismatches, f"citationConversionPlan.{key}", citation_plan.get(key), expected_value)
    if citation_plan is not None and isinstance(expected.get("citationConversionPlanCandidatesByKind"), dict):
        assert_dict_contains_values(
            mismatches,
            "citationConversionPlanCandidatesByKind",
            citation_plan.get("candidatesByKind"),
            expected.get("citationConversionPlanCandidatesByKind") or {},
        )
    if citation_plan is not None and isinstance(expected.get("citationConversionPlanMatchSummary"), dict):
        assert_dict_contains_values(
            mismatches,
            "citationConversionPlanMatchSummary",
            citation_plan.get("matchSummary"),
            expected.get("citationConversionPlanMatchSummary") or {},
        )
    if citation_plan is not None and "citationConversionPlanUnmatchedLabelsIncludes" in expected:
        actual_labels = set(citation_plan.get("unmatchedLabels") or [])
        for label in expected.get("citationConversionPlanUnmatchedLabelsIncludes") or []:
            if label not in actual_labels:
                mismatches.append(f"citationConversionPlan.unmatchedLabels missing expected item {label!r}")
    if citation_plan is not None and "citationConversionPlanManualReviewReasonsInclude" in expected:
        actual_reasons = set(citation_plan.get("manualReviewReasons") or [])
        for reason in expected.get("citationConversionPlanManualReviewReasonsInclude") or []:
            if reason not in actual_reasons:
                mismatches.append(f"citationConversionPlan.manualReviewReasons missing expected item {reason!r}")
    if citation_plan is not None and "citationConversionPlanDuplicateReferenceLabelsIncludes" in expected:
        actual_labels = set(citation_plan.get("duplicateReferenceLabels") or [])
        for label in expected.get("citationConversionPlanDuplicateReferenceLabelsIncludes") or []:
            if label not in actual_labels:
                mismatches.append(f"citationConversionPlan.duplicateReferenceLabels missing expected item {label!r}")
    requires_hybrid_attachment_report = expected.get("requiresHybridAttachmentReport")
    if requires_hybrid_attachment_report is not None:
        assert_equal(
            mismatches,
            "requiresHybridAttachmentReport",
            hybrid_attachment_report_path is not None and hybrid_attachment_report is not None,
            bool(requires_hybrid_attachment_report),
        )
    if hybrid_attachment_report is not None and isinstance(expected.get("hybridAttachmentReport"), dict):
        for key, expected_value in (expected.get("hybridAttachmentReport") or {}).items():
            assert_equal(mismatches, f"hybridAttachmentReport.{key}", hybrid_attachment_report.get(key), expected_value)
    if hybrid_attachment_report is not None and isinstance(expected.get("hybridAttachmentReportActionCounts"), dict):
        assert_dict_contains_values(
            mismatches,
            "hybridAttachmentReportActionCounts",
            hybrid_attachment_report,
            expected.get("hybridAttachmentReportActionCounts") or {},
        )
    if hybrid_rebuild_report is not None and isinstance(expected.get("hybridRebuildReport"), dict):
        for key, expected_value in (expected.get("hybridRebuildReport") or {}).items():
            assert_equal(mismatches, f"hybridRebuildReport.{key}", hybrid_rebuild_report.get(key), expected_value)
    if hybrid_rebuild_report is not None and isinstance(expected.get("hybridRebuildReportActionCounts"), dict):
        assert_dict_contains_values(
            mismatches,
            "hybridRebuildReportActionCounts",
            hybrid_rebuild_report,
            expected.get("hybridRebuildReportActionCounts") or {},
        )

    return {
        "strategySelection": str(strategy_path) if strategy_path is not None else None,
        "thesisIr": str(thesis_ir_path) if thesis_ir_path is not None else None,
        "preflightReport": str(preflight_path) if preflight_path is not None else None,
        "citationConversionPlan": str(citation_plan_path) if citation_plan_path is not None else None,
        "hybridAttachmentReport": str(hybrid_attachment_report_path) if hybrid_attachment_report_path is not None else None,
        "hybridRebuildReport": str(hybrid_rebuild_report_path) if hybrid_rebuild_report_path is not None else None,
    }


def resolve_fixture_input(fixture_dir: Path, expected_spec: Dict[str, object]) -> Path:
    input_value = expected_spec.get("input")
    if not isinstance(input_value, str) or not input_value.strip():
        raise SystemExit(f"Fixture {fixture_dir.name} expected.json must define a non-empty input path")
    return (fixture_dir / input_value).resolve()


def run_fixture(fixture_dir: Path, profile: str, output_dir: Path, expected_spec: Dict[str, object]) -> Tuple[Dict[str, object], List[str]]:
    fixture_output_dir = output_dir / fixture_dir.name
    fixture_output_dir.mkdir(parents=True, exist_ok=True)
    input_path = resolve_fixture_input(fixture_dir, expected_spec)
    mode = str(expected_spec.get("mode") or "audit-only")
    fixture_profile = str(expected_spec.get("profile") or profile)
    compliance = str(expected_spec.get("compliance") or "default")
    command = [
        PYTHON,
        "scripts/thesis_format.py",
        str(input_path),
        "--profile",
        fixture_profile,
        "--mode",
        mode,
        "--compliance",
        compliance,
        "--output-dir",
        str(fixture_output_dir),
    ]
    completed = subprocess.run(command, cwd=str(ROOT), check=False, capture_output=True, text=True)
    manifest_path = fixture_output_dir / "dispatch_manifest.json"
    manifest = read_json(manifest_path) if manifest_path.exists() else {}
    mismatches: List[str] = []
    expected = expected_spec.get("expected") or {}
    if not isinstance(expected, dict):
        expected = {}

    if "sourceKind" in expected:
        assert_equal(mismatches, "sourceKind", manifest.get("sourceKind"), expected.get("sourceKind"))
    if "dispatcherStatus" in expected:
        assert_equal(mismatches, "dispatcherStatus", manifest.get("status"), expected.get("dispatcherStatus"))
    if "effectiveMode" in expected:
        assert_equal(mismatches, "effectiveMode", manifest.get("effectiveMode"), expected.get("effectiveMode"))

    preflight = manifest.get("preflightDecision") or {}
    if "documentRiskClass" in expected:
        assert_equal(mismatches, "documentRiskClass", preflight.get("documentRiskClass"), expected.get("documentRiskClass"))
    if "recommendedMode" in expected:
        assert_equal(mismatches, "recommendedMode", preflight.get("recommendedMode"), expected.get("recommendedMode"))
    if "minConfirmationRequests" in expected:
        actual_count = int(preflight.get("confirmationRequestCount") or 0)
        if actual_count < int(expected.get("minConfirmationRequests") or 0):
            mismatches.append(
                f"confirmationRequestCount: expected >= {expected.get('minConfirmationRequests')!r}, got {actual_count!r}"
            )
    if "blockedAutoRepairsIncludes" in expected:
        actual_blocked = list(preflight.get("blockedAutoRepairs") or [])
        for item in expected.get("blockedAutoRepairsIncludes") or []:
            if item not in actual_blocked:
                mismatches.append(f"blockedAutoRepairs missing expected item {item!r}")

    result_artifact_paths = check_result_artifacts(manifest, expected, mismatches)
    report_paths = check_validation_reports(manifest, expected, mismatches)
    strategy_paths = check_strategy_and_ir(manifest, expected, mismatches)

    result = {
        "name": fixture_dir.name,
        "input": str(input_path),
        "mode": mode,
        "profile": fixture_profile,
        "compliance": compliance,
        "outputDir": str(fixture_output_dir),
        "command": command,
        "exitCode": completed.returncode,
        "stdout": completed.stdout.strip(),
        "stderr": completed.stderr.strip(),
        "manifest": str(manifest_path) if manifest_path.exists() else None,
        "repairPlan": result_artifact_paths["repairPlan"],
        "repairExecution": result_artifact_paths["repairExecution"],
        "repairedDocx": result_artifact_paths["repairedDocx"],
        "repairedPdf": result_artifact_paths["repairedPdf"],
        "strategySelection": strategy_paths["strategySelection"],
        "thesisIr": strategy_paths["thesisIr"],
        "preflightReport": strategy_paths["preflightReport"],
        "citationConversionPlan": strategy_paths["citationConversionPlan"],
        "hybridAttachmentReport": strategy_paths["hybridAttachmentReport"],
        "hybridRebuildReport": strategy_paths["hybridRebuildReport"],
        "validationReport": report_paths["validationReport"],
        "renderValidationReport": report_paths["renderValidationReport"],
        "previewManifest": report_paths["previewManifest"],
        "passed": completed.returncode == 0 and not mismatches,
        "mismatches": mismatches,
    }
    return result, mismatches


def main() -> None:
    args = parse_args()
    fixtures_dir = (ROOT / args.fixtures_dir).resolve()
    output_dir = (ROOT / args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    fixtures = discover_fixtures(fixtures_dir)
    results: List[Dict[str, object]] = []
    passed_count = 0
    executed_count = 0
    skipped_count = 0
    for fixture in fixtures:
        expected_path = Path(str(fixture["expectedSpec"]))
        if not expected_path.exists():
            fixture["status"] = "missing_expected_spec"
            results.append(fixture)
            continue
        expected_spec = read_json(expected_path)
        fixture_tags = normalize_tags(expected_spec.get("tags"))
        if args.include_tag and not set(args.include_tag).issubset(set(fixture_tags)):
            continue
        should_run, skip_reason, detected_commands = should_run_fixture(expected_spec, args.include_tag)
        if not should_run:
            skipped_count += 1
            results.append(
                {
                    "name": fixture["name"],
                    "path": fixture["path"],
                    "status": "skipped",
                    "skipReason": skip_reason,
                    "tags": fixture_tags,
                    "requiredCommands": normalize_required_commands(expected_spec.get("requiresCommands")),
                    "detectedCommands": detected_commands,
                }
            )
            continue
        result, _mismatches = run_fixture(Path(str(fixture["path"])), args.profile, output_dir, expected_spec)
        result["tags"] = fixture_tags
        result["requiredCommands"] = normalize_required_commands(expected_spec.get("requiresCommands"))
        executed_count += 1
        if result["passed"]:
            passed_count += 1
        results.append(result)

    payload = {
        "profile": args.profile,
        "fixturesDir": str(fixtures_dir),
        "fixtureCount": len(fixtures),
        "executedCount": executed_count,
        "skippedCount": skipped_count,
        "passedCount": passed_count,
        "failedCount": executed_count - passed_count,
        "fixtures": results,
        "notes": [
            "Fixtures with expected.json are executed through thesis_format.py and checked against dispatch_manifest.json.",
            "Repair-mode fixtures can now assert repairPlan, repairExecution, repairedDocx, and repairedPdf artifact presence.",
            "The runner now also validates the presence and minimal schema of validation_report.json, render_validation_report.json, and preview_manifest.json when those artifacts are declared.",
            "Fixtures may be skipped when include-tag filters do not match or requiredCommands are unavailable in the current environment.",
            "Current assertions still stop short of full DOCX content diffs or visual golden comparison.",
        ],
    }
    (output_dir / "fixture_inventory.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "fixtureCount": len(fixtures),
                "executedCount": executed_count,
                "skippedCount": skipped_count,
                "passedCount": passed_count,
                "inventory": str(output_dir / "fixture_inventory.json"),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
