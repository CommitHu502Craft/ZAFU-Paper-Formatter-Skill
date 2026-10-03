#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import copy
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import yaml
from formatter_core.checks import CheckPolicy
from formatter_core.workspace import Workspace
from formatter_core.workspace import file_digest
from formatter_core.reports import write_summary
from formatter_core.reference_format import check_reference_format
from formatter_backends.latex import run_latex
from formatter_backends.word import common_flags, repair_commands


ROOT = Path(__file__).resolve().parents[1]
PYTHON = sys.executable
DEFAULT_TEMPLATE = ROOT / "浙江农林大学毕业论文模板参考.docx"
DEFAULT_FRONT_TEMPLATE = ROOT / "assets" / "zafu_front_matter_template.docx"
PROFILES_ROOT = ROOT / "profiles"
PRODUCT_MODE = "thesis_finalization"


@dataclass(frozen=True)
class ProfileConfig:
    name: str
    profile_dir: Path
    rules_yaml: Path
    template_docx: Optional[Path]
    front_matter_template: Optional[Path]
    style_map_yaml: Optional[Path]
    front_matter_policy_yaml: Optional[Path]
    validators_yaml: Optional[Path]
    metadata: Dict[str, object]


def read_yaml(path: Path) -> Dict[str, object]:
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def resolve_profile_resource(base_dir: Path, value: Optional[str]) -> Optional[Path]:
    if not value:
        return None
    candidate = Path(value)
    if candidate.is_absolute():
        return candidate
    profile_relative = (base_dir / candidate).resolve()
    if profile_relative.exists():
        return profile_relative
    root_relative = (ROOT / candidate).resolve()
    if root_relative.exists():
        return root_relative
    return profile_relative


def resolve_rules_yaml(profile_dir: Path, payload: Dict[str, object], profile_name: str) -> Path:
    extends = payload.get("extends")
    if isinstance(extends, str) and extends.strip():
        resolved = resolve_profile_resource(profile_dir, extends.strip())
        if resolved is not None:
            return resolved
    local_rules = profile_dir / "rules.yaml"
    if local_rules.exists():
        return local_rules.resolve()
    if profile_name == "zafu_2022":
        return (ROOT / "references" / "zafu_2022_rules.yaml").resolve()
    raise SystemExit(f"Profile {profile_name} has no usable rules.yaml")


def infer_template_docx(profile_dir: Path, profile_name: str) -> Optional[Path]:
    for candidate_name in ("template.docx", "profile_template.docx"):
        candidate = profile_dir / candidate_name
        if candidate.exists():
            return candidate.resolve()
    if profile_name == "zafu_2022" and DEFAULT_TEMPLATE.exists():
        return DEFAULT_TEMPLATE.resolve()
    return None


def infer_front_matter_template(profile_dir: Path, profile_name: str) -> Optional[Path]:
    for candidate_name in ("front_matter_template.docx", "template_front_matter.docx"):
        candidate = profile_dir / candidate_name
        if candidate.exists():
            return candidate.resolve()
    if profile_name == "zafu_2022" and DEFAULT_FRONT_TEMPLATE.exists():
        return DEFAULT_FRONT_TEMPLATE.resolve()
    return None


def load_profile(profile_dir: Path) -> ProfileConfig:
    profile_name = profile_dir.name
    rules_payload = read_yaml(profile_dir / "rules.yaml") if (profile_dir / "rules.yaml").exists() else {}
    style_map_yaml = (profile_dir / "style_map.yaml").resolve() if (profile_dir / "style_map.yaml").exists() else None
    front_matter_policy_yaml = (
        (profile_dir / "front_matter_policy.yaml").resolve() if (profile_dir / "front_matter_policy.yaml").exists() else None
    )
    validators_yaml = (profile_dir / "validators.yaml").resolve() if (profile_dir / "validators.yaml").exists() else None
    metadata: Dict[str, object] = {
        "rules": rules_payload,
        "styleMap": read_yaml(style_map_yaml) if style_map_yaml else {},
        "frontMatterPolicy": read_yaml(front_matter_policy_yaml) if front_matter_policy_yaml else {},
        "validators": read_yaml(validators_yaml) if validators_yaml else {},
    }
    return ProfileConfig(
        name=profile_name,
        profile_dir=profile_dir.resolve(),
        rules_yaml=resolve_rules_yaml(profile_dir, rules_payload, profile_name),
        template_docx=infer_template_docx(profile_dir, profile_name),
        front_matter_template=infer_front_matter_template(profile_dir, profile_name),
        style_map_yaml=style_map_yaml,
        front_matter_policy_yaml=front_matter_policy_yaml,
        validators_yaml=validators_yaml,
        metadata=metadata,
    )


def discover_profiles() -> Dict[str, ProfileConfig]:
    profiles: Dict[str, ProfileConfig] = {}
    if not PROFILES_ROOT.exists():
        return profiles
    for child in sorted(PROFILES_ROOT.iterdir()):
        if not child.is_dir():
            continue
        if not (child / "profile.md").exists() and not (child / "rules.yaml").exists():
            continue
        profiles[child.name] = load_profile(child)
    return profiles


PROFILES = discover_profiles()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Finalize a thesis into a school-compliant graduation paper deliverable.")
    parser.add_argument("input", help="Source thesis file (.docx, .md, .txt)")
    parser.add_argument("--profile", default="zafu_2022", choices=sorted(PROFILES), help="Formatting profile")
    parser.add_argument(
        "--mode",
        default="conservative-repair",
        choices=["audit-only", "conservative-repair", "rebuild"],
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--output-dir", help="Output base; default: thesis-output beside input, with an isolated subdirectory per run")
    parser.add_argument("--project-name", help="Project name; defaults to input filename, never an inferred thesis title")
    parser.add_argument("--backend", choices=["word", "latex", "both"], default="word", help="Independent output backend")
    parser.add_argument("--no-compile", action="store_true", help="Generate LaTeX source without invoking a TeX engine")
    parser.add_argument("--latex-engine", choices=["auto", "xelatex", "tectonic"], default="auto", help="auto prefers XeLaTeX, then Tectonic on PATH; never installs a runtime")
    parser.add_argument("--latex-metadata", help="Explicit JSON metadata for external CLS front matter; LaTeX only")
    parser.add_argument("--latex-bibliography", help="Explicit BibTeX file; LaTeX-only GB/T 7714-2015 author-year citations")
    parser.add_argument("--compile-timeout", type=int, default=120, help="Timeout in seconds per TeX compilation pass")
    parser.add_argument("--check", choices=["structural", "layout", "visual"], default="structural", help="Default: no PDF rendering or page images")
    parser.add_argument("--export-pdf", action="store_true", help="Export Word PDF independently of visual review")
    parser.add_argument("--allow-structural-rebuild", action="store_true", help="Explicitly allow DOCX whole-document rebuild and template front-matter replacement")
    parser.add_argument("--header-mode", choices=["fixed", "thesis-title", "preserve"], help="Word header text policy; default from profile")
    parser.add_argument("--header-text", help="Fixed Word header text; implies --header-mode fixed")
    parser.add_argument("--latin-font", help="Explicit Word Latin font override (e.g. Times New Roman); preserves the cover and does not affect LaTeX")
    parser.add_argument(
        "--compliance",
        default="default",
        choices=["default", "strict-school"],
        help="Compliance profile overlay; strict-school prefers template-conformant repair defaults over source-preserving heuristics.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Write the dispatch manifest only")
    parser.add_argument(
        "--semantic-overrides",
        help="Optional agent-provided semantic override JSON applied when building ThesisIR",
    )
    parser.add_argument(
        "--visual-refinement-plan",
        help="Optional visual_refinement_plan.json; when given, a single whitelisted visual refinement pass runs after the first repair and the result is re-rendered",
    )
    args = parser.parse_args()
    if args.backend == "latex" and (args.header_mode or args.header_text):
        parser.error("Header options are Word-only; the external CLS manages LaTeX headers")
    if args.header_text is not None and args.header_mode not in {None, "fixed"}:
        parser.error("--header-text requires fixed header mode")
    if args.backend == "word" and (args.latex_metadata or args.latex_bibliography):
        parser.error("LaTeX metadata/bibliography options require --backend latex or both; Word does not manage citations")
    for option in (args.latex_metadata, args.latex_bibliography):
        if option and not Path(option).is_file():
            parser.error(f"LaTeX input file does not exist: {option}")
    if args.visual_refinement_plan and args.check != "visual":
        parser.error("--visual-refinement-plan requires --check visual")
    if args.visual_refinement_plan and args.backend == "latex":
        parser.error("Word visual refinement is not supported by the LaTeX-only backend")
    if args.compile_timeout < 1 or args.compile_timeout > 600:
        parser.error("--compile-timeout must be between 1 and 600")
    if args.backend != "word" and Path(args.input).suffix.lower() == ".docx":
        parser.error("LaTeX MVP supports Markdown/TXT only; DOCX native objects require the Word backend")
    if args.backend != "word" and args.mode == "audit-only":
        parser.error("--mode audit-only is a Word audit operation; do not combine it with LaTeX generation")
    if args.visual_refinement_plan and args.mode == "audit-only":
        parser.error("Visual refinement requires a generated Word candidate, not audit-only mode")
    if args.visual_refinement_plan and not Path(args.visual_refinement_plan).is_file():
        parser.error("Visual refinement plan does not exist")
    return args


def run_step(command: List[str], workdir: Path) -> None:
    completed = subprocess.run(command, cwd=str(workdir), check=False)
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)


def source_kind(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return "docx"
    if suffix in {".md", ".markdown", ".txt"}:
        return "text"
    raise SystemExit(f"Unsupported input type: {path.suffix}")


def add_common_flags(command: List[str], profile: ProfileConfig, *, include_template: bool = True, include_validators: bool = True) -> List[str]:
    return common_flags(command, profile, include_template=include_template, include_validators=include_validators)


def requested_effective_mode(input_path: Path, mode: str) -> str:
    kind = source_kind(input_path)
    effective_mode = mode
    if kind == "text" and mode == "conservative-repair":
        effective_mode = "rebuild"
    return effective_mode


def build_manifest(input_path: Path, profile: ProfileConfig, mode: str, output_dir: Path, compliance: str) -> Dict[str, object]:
    kind = source_kind(input_path)
    effective_mode = requested_effective_mode(input_path, mode)
    return {
        "input": str(input_path),
        "sourceKind": kind,
        "productMode": PRODUCT_MODE,
        "profile": profile.name,
        "requestedExpertMode": mode,
        "effectiveExpertMode": effective_mode,
        "requestedMode": mode,
        "effectiveMode": effective_mode,
        "complianceMode": compliance,
        "outputDir": str(output_dir),
        "defaults": {
            "rulesYaml": str(profile.rules_yaml),
            "templateDocx": str(profile.template_docx) if profile.template_docx else None,
            "frontMatterTemplate": str(profile.front_matter_template) if profile.front_matter_template else None,
        },
        "profileResources": {
            "profileDir": str(profile.profile_dir),
            "styleMapYaml": str(profile.style_map_yaml) if profile.style_map_yaml else None,
            "frontMatterPolicyYaml": str(profile.front_matter_policy_yaml) if profile.front_matter_policy_yaml else None,
            "validatorsYaml": str(profile.validators_yaml) if profile.validators_yaml else None,
        },
        "profileDefaults": {
            "recommendedMode": (
                ((profile.metadata.get("rules") or {}).get("defaults") or {}).get("recommended_mode")
                if isinstance(profile.metadata.get("rules"), dict)
                else None
            ),
            "blockedAutoRepairs": (
                ((profile.metadata.get("rules") or {}).get("defaults") or {}).get("blocked_auto_repairs")
                if isinstance(profile.metadata.get("rules"), dict)
                else None
            ),
            "styleRoles": (
                (profile.metadata.get("styleMap") or {}).get("style_roles")
                if isinstance(profile.metadata.get("styleMap"), dict)
                else None
            ),
            "frontMatterPolicy": (
                (profile.metadata.get("frontMatterPolicy") or {}).get("policy")
                if isinstance(profile.metadata.get("frontMatterPolicy"), dict)
                else None
            ),
            "validators": (
                (profile.metadata.get("validators") or {}).get("validators")
                if isinstance(profile.metadata.get("validators"), dict)
                else None
            ),
        },
        "notes": [
            "The public product behavior is thesis finalization; expert repair modes remain internal compatibility controls.",
            "The dispatcher preserves existing script boundaries and does not bypass deterministic OOXML writers.",
            "Current repository scripts already support preflight, audit, planning, repair, and validation.",
        ],
    }


def preflight_output_path(kind: str, output_dir: Path) -> Path:
    return output_dir / ("preflight_report.json" if kind == "docx" else "source_preflight_report.json")


def evidence_output_path(output_dir: Path) -> Path:
    return output_dir / "source_evidence.json"


def numbering_output_path(output_dir: Path) -> Path:
    return output_dir / "numbering_recovery.json"


def thesis_ir_output_path(output_dir: Path) -> Path:
    return output_dir / "thesis_ir.json"


def thesis_ir_validation_path(output_dir: Path) -> Path:
    return output_dir / "thesis_ir_validation.json"


def strategy_output_path(output_dir: Path) -> Path:
    return output_dir / "strategy_selection.json"


def citation_conversion_output_path(output_dir: Path) -> Path:
    return output_dir / "citation_conversion_plan.json"


def hybrid_report_path(output_dir: Path) -> Path:
    return output_dir / "hybrid_rebuild_report.json"


def hybrid_source_docx_path(output_dir: Path) -> Path:
    return output_dir / "hybrid_source.docx"


def hybrid_attached_docx_path(output_dir: Path) -> Path:
    return output_dir / "hybrid_source_attached.docx"


def hybrid_attachment_report_path(output_dir: Path) -> Path:
    return output_dir / "hybrid_attachment_report.json"


def evidence_command(input_path: Path, profile: ProfileConfig, output_dir: Path) -> List[str]:
    command = [PYTHON, "scripts/extract_source_evidence.py", str(input_path), "--output", str(evidence_output_path(output_dir))]
    if source_kind(input_path) == "docx" and profile.template_docx:
        command.extend(["--template-docx", str(profile.template_docx)])
    command.extend(["--rules-yaml", str(profile.rules_yaml)])
    return command


def apply_compliance_overlay(base_rules: Dict[str, object], compliance: str) -> Dict[str, object]:
    rules = copy.deepcopy(base_rules)
    defaults = dict(rules.get("defaults") or {})
    references = dict(rules.get("references") or {})
    tables = dict(rules.get("tables") or {})
    equations = dict(rules.get("equations") or {})

    defaults["compliance_mode"] = compliance
    references.setdefault("strip_numeric_labels_when_unnumbered", True)
    references.setdefault("auto_split_collapsed_entries", True)
    references.setdefault("auto_format_entries", True)
    references.setdefault("auto_convert_body_citations", False)
    equations.setdefault("display_equation_single_spacing", True)

    if compliance == "strict-school":
        defaults["strict_school_compliance"] = True
        tables["header_rows"] = 1
        tables["max_header_rows"] = 1
        tables["repeat_group_header_rules"] = False
        group_header_rule = dict(tables.get("group_header_rule") or {})
        group_header_rule["enabled"] = False
        tables["group_header_rule"] = group_header_rule
        references["auto_strip_numeric_labels_when_bibliography_numbering_none"] = True
        references["strict_layout_fix_only"] = True
    else:
        defaults["strict_school_compliance"] = False

    rules["defaults"] = defaults
    rules["references"] = references
    rules["tables"] = tables
    rules["equations"] = equations
    return rules


def materialize_effective_rules_yaml(profile: ProfileConfig, compliance: str, output_dir: Path) -> Path:
    base_rules = read_yaml(profile.rules_yaml)
    effective_rules = apply_compliance_overlay(base_rules, compliance)
    target = output_dir / "effective_rules.yaml"
    target.write_text(yaml.safe_dump(effective_rules, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return target


def numbering_command(output_dir: Path) -> List[str]:
    return [PYTHON, "scripts/recover_numbering.py", str(evidence_output_path(output_dir)), "--output", str(numbering_output_path(output_dir))]


def thesis_ir_command(output_dir: Path, semantic_overrides: Optional[str] = None) -> List[str]:
    command = [PYTHON, "scripts/thesis_ir.py", "--evidence-json", str(evidence_output_path(output_dir)), "--output", str(thesis_ir_output_path(output_dir))]
    if semantic_overrides:
        command.extend(["--semantic-overrides", str(Path(semantic_overrides).resolve())])
    return command


def thesis_ir_validation_command(output_dir: Path) -> List[str]:
    return [
        PYTHON,
        "scripts/validate_thesis_ir.py",
        str(thesis_ir_output_path(output_dir)),
        "--output",
        str(thesis_ir_validation_path(output_dir)),
    ]


def strategy_command(output_dir: Path, kind: str) -> List[str]:
    return [
        PYTHON,
        "scripts/select_strategy.py",
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
        "--preflight-json",
        str(preflight_output_path(kind, output_dir)),
        "--output",
        str(strategy_output_path(output_dir)),
    ]


def citation_conversion_command(profile: ProfileConfig, output_dir: Path) -> List[str]:
    return [
        PYTHON,
        "scripts/citation_conversion_plan.py",
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
        "--rules-yaml",
        str(profile.rules_yaml),
        "--output",
        str(citation_conversion_output_path(output_dir)),
    ]


def preflight_command(input_path: Path, profile: ProfileConfig, output_dir: Path) -> List[str]:
    kind = source_kind(input_path)
    command = [
        PYTHON,
        "scripts/preflight_semantic_normalization.py",
        str(input_path),
        "--output",
        str(preflight_output_path(kind, output_dir)),
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
    ]
    return add_common_flags(command, profile, include_template=(kind == "docx"))


def read_json(path: Path) -> Dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def render_output_dir(output_dir: Path) -> Path:
    return output_dir / "render_validation"


def render_report_path(output_dir: Path) -> Path:
    return render_output_dir(output_dir) / "render_validation_report.json"


def expected_pdf_path(docx_path: Path, output_dir: Path) -> Path:
    return render_output_dir(output_dir) / f"{docx_path.stem}.pdf"


def update_manifest_from_preflight(
    manifest: Dict[str, object],
    preflight_report: Dict[str, object],
    *,
    effective_mode: str,
    status: str,
    stop_reason: Optional[str] = None,
) -> None:
    manifest["effectiveMode"] = effective_mode
    manifest["effectiveExpertMode"] = effective_mode
    manifest["status"] = status
    if stop_reason:
        manifest["stopReason"] = stop_reason
    manifest["preflightDecision"] = {
        "documentRiskClass": preflight_report.get("documentRiskClass"),
        "recommendedMode": preflight_report.get("recommendedMode"),
        "requiresUserConfirmation": bool(preflight_report.get("confirmationRequests")),
        "confirmationRequestCount": len(preflight_report.get("confirmationRequests") or []),
        "blockedAutoRepairs": preflight_report.get("blockedAutoRepairs") or [],
        "riskReasons": preflight_report.get("riskReasons") or [],
    }


def update_manifest_with_outputs(manifest: Dict[str, object], input_path: Path, output_dir: Path) -> None:
    artifacts: Dict[str, object] = {"dispatchManifest": str(Path(str(manifest["runDir"])) / "manifest.json")}
    preflight = preflight_output_path(str(manifest["sourceKind"]), output_dir)
    if preflight.exists():
        artifacts["preflightReport"] = str(preflight)
    effective_rules = output_dir / "effective_rules.yaml"
    if effective_rules.exists():
        artifacts["effectiveRulesYaml"] = str(effective_rules)
    audit_report = output_dir / "audit_report.json"
    if audit_report.exists():
        artifacts["auditReport"] = str(audit_report)
    repair_plan = output_dir / "repair_plan.json"
    if repair_plan.exists():
        artifacts["repairPlan"] = str(repair_plan)
    repair_execution = output_dir / "repair_execution.json"
    if repair_execution.exists():
        artifacts["repairExecution"] = str(repair_execution)
    repaired_docx = output_dir / "repaired.docx"
    if repaired_docx.exists():
        artifacts["repairedDocx"] = str(repaired_docx)
    validation_report = output_dir / "validation_report.json"
    if validation_report.exists():
        artifacts["validationReport"] = str(validation_report)
    review_summary = output_dir / "review_summary.json"
    if review_summary.exists():
        artifacts["reviewSummary"] = str(review_summary)
    evidence_report = evidence_output_path(output_dir)
    if evidence_report.exists():
        artifacts["sourceEvidence"] = str(evidence_report)
    numbering_report = numbering_output_path(output_dir)
    if numbering_report.exists():
        artifacts["numberingRecovery"] = str(numbering_report)
    thesis_ir_report = thesis_ir_output_path(output_dir)
    if thesis_ir_report.exists():
        artifacts["thesisIr"] = str(thesis_ir_report)
    thesis_ir_validation = thesis_ir_validation_path(output_dir)
    if thesis_ir_validation.exists():
        artifacts["thesisIrValidation"] = str(thesis_ir_validation)
    strategy_report = strategy_output_path(output_dir)
    if strategy_report.exists():
        artifacts["strategySelection"] = str(strategy_report)
    citation_report = citation_conversion_output_path(output_dir)
    if citation_report.exists():
        artifacts["citationConversionPlan"] = str(citation_report)
    hybrid_report = hybrid_report_path(output_dir)
    if hybrid_report.exists():
        artifacts["hybridRebuildReport"] = str(hybrid_report)
    hybrid_source_docx = hybrid_source_docx_path(output_dir)
    if hybrid_source_docx.exists():
        artifacts["hybridSourceDocx"] = str(hybrid_source_docx)
    hybrid_attached_docx = hybrid_attached_docx_path(output_dir)
    if hybrid_attached_docx.exists():
        artifacts["hybridAttachedDocx"] = str(hybrid_attached_docx)
    hybrid_attachment_report = hybrid_attachment_report_path(output_dir)
    if hybrid_attachment_report.exists():
        artifacts["hybridAttachmentReport"] = str(hybrid_attachment_report)

    render_report = render_report_path(output_dir)
    if render_report.exists():
        artifacts["renderValidationReport"] = str(render_report)
        try:
            render_report_payload = read_json(render_report)
        except json.JSONDecodeError:
            render_report_payload = {}
        pdf_path = render_report_payload.get("pdfPath")
        if isinstance(pdf_path, str) and pdf_path:
            artifacts["repairedPdf"] = pdf_path
        else:
            expected_pdf = expected_pdf_path(repaired_docx if repaired_docx.exists() else input_path, output_dir)
            if expected_pdf.exists():
                artifacts["repairedPdf"] = str(expected_pdf)
        preview_dir = render_report_payload.get("previewDir")
        if isinstance(preview_dir, str) and preview_dir and Path(preview_dir).is_dir():
            artifacts["renderPreviewDir"] = preview_dir

    manifest["artifacts"] = artifacts


def issue_preview(items: object, *, limit: int = 5) -> List[str]:
    if not isinstance(items, list):
        return []
    lines: List[str] = []
    for item in items[:limit]:
        if not isinstance(item, dict):
            continue
        issue_type = str(item.get("type") or item.get("kind") or "issue")
        details = item.get("details")
        if isinstance(details, list) and details:
            first = details[0]
            if isinstance(first, dict):
                text = first.get("text") or first.get("message") or first.get("kind")
                if text:
                    lines.append(f"{issue_type}: {text}")
                    continue
        message = item.get("message") or item.get("expected") or item.get("reason")
        if message:
            lines.append(f"{issue_type}: {message}")
        else:
            lines.append(issue_type)
    return lines


def build_review_summary(output_dir: Path, manifest: Dict[str, object]) -> Optional[Dict[str, object]]:
    validation_path = output_dir / "validation_report.json"
    if not validation_path.exists():
        return None
    validation = read_json(validation_path)
    preflight = read_json(preflight_output_path(str(manifest["sourceKind"]), output_dir))
    repair_execution_path = output_dir / "repair_execution.json"
    repair_execution = read_json(repair_execution_path) if repair_execution_path.exists() else {}
    citation_plan_path = output_dir / "citation_conversion_plan.json"
    citation_plan = read_json(citation_plan_path) if citation_plan_path.exists() else {}
    quality_gate = validation.get("qualityGate") or {}
    execution_log = repair_execution.get("executionLog") or []
    auto_fixed_actions = [
        item.get("action")
        for item in execution_log
        if isinstance(item, dict) and item.get("action")
    ]
    blocked_by_policy = list((manifest.get("preflightDecision") or {}).get("blockedAutoRepairs") or [])
    citation_plan_action = citation_plan.get("recommendedAction")
    summary = {
        "productMode": PRODUCT_MODE,
        "source": manifest.get("input"),
        "profile": manifest.get("profile"),
        "outputDocx": str((output_dir / "repaired.docx").resolve()) if (output_dir / "repaired.docx").exists() else None,
        "qualityGate": quality_gate,
        "readyForDelivery": bool(quality_gate.get("passed")),
        "hardFailurePreview": issue_preview(quality_gate.get("hardFailures")),
        "warningPreview": issue_preview(quality_gate.get("warnings")),
        "riskSummary": {
            "documentRiskClass": (manifest.get("preflightDecision") or {}).get("documentRiskClass"),
            "recommendedMode": (manifest.get("preflightDecision") or {}).get("recommendedMode"),
            "riskReasons": (manifest.get("preflightDecision") or {}).get("riskReasons") or [],
        },
        "remediationSummary": {
            "detected": {
                "warningCount": quality_gate.get("warningCount"),
                "hardFailureCount": quality_gate.get("hardFailureCount"),
                "citationConversionPlanAction": citation_plan_action,
            },
            "autoFixedActions": sorted({str(item) for item in auto_fixed_actions if item}),
            "blockedByPolicy": blocked_by_policy,
        },
        "nextAction": (
            "deliverable_ready"
            if quality_gate.get("passed")
            else "manual_review_required"
        ),
        "notes": [
            "The formatter always attempts to produce a thesis deliverable and separates hard failures from warnings.",
            "Warnings indicate remaining quality gaps; hard failures indicate the output is not ready as a final thesis deliverable.",
        ],
        "frontMatterIssues": ((preflight.get("frontMatter") or {}).get("issues") or [])[:20],
    }
    return summary


def profile_policy_flag(profile: ProfileConfig, key: str, default: bool = False) -> bool:
    payload = profile.metadata.get("frontMatterPolicy")
    if not isinstance(payload, dict):
        return default
    policy = payload.get("policy")
    if not isinstance(policy, dict):
        return default
    value = policy.get(key)
    return bool(default if value is None else value)


def decide_mode(requested_mode: str, preflight_report: Dict[str, object], source_kind_value: str, profile: ProfileConfig) -> Tuple[str, List[str]]:
    """Deliver-first mode decision.

    The formatter always tries to produce a repaired deliverable. Risk class C
    downgrades to audit-only ONLY when the source itself is not reliably
    extractable (parse failure / body loss signals); style pollution alone is
    handled by strategy selection (preserve_first / hybrid_rebuild) instead of
    refusing to format. Users can always force audit-only explicitly.
    """
    effective_mode = requested_mode
    reasons: List[str] = []
    recommended_mode = preflight_report.get("recommendedMode")
    document_risk_class = preflight_report.get("documentRiskClass")

    if source_kind_value == "text" and requested_mode == "conservative-repair":
        effective_mode = "rebuild"
        reasons.append("text_source_conservative_repair_upgraded_to_rebuild")

    if requested_mode == "audit-only":
        reasons.append("user_requested_audit_only")
        return "audit-only", reasons

    hard_stop_signals = {
        "document_xml_unparseable",
        "docx_package_corrupt",
        "body_text_unextractable",
        "body_text_mostly_lost",
    }
    risk_reasons = {str(item.get("kind") or item) if isinstance(item, dict) else str(item) for item in (preflight_report.get("riskReasons") or [])}
    if risk_reasons & hard_stop_signals:
        effective_mode = "audit-only"
        reasons.append("source_not_reliably_extractable_forces_audit_only")
        return effective_mode, reasons

    if document_risk_class == "C":
        reasons.append("risk_class_c_continues_with_deliver_first_policy")
    if recommended_mode == "audit-only":
        reasons.append("preflight_recommended_audit_only_downgraded_to_warning")
    return effective_mode, reasons


def reconcile_mode_with_strategy(effective_mode: str, strategy_report: Dict[str, object]) -> Tuple[str, List[str]]:
    if effective_mode == "audit-only":
        return effective_mode, ["explicit_audit_only_is_non_escalatable"]
    strategy_mode = strategy_report.get("executionMode")
    if not isinstance(strategy_mode, str) or not strategy_mode:
        return effective_mode, []
    if strategy_mode == effective_mode:
        return effective_mode, []
    return strategy_mode, [f"strategy_selected_{strategy_report.get('chosenStrategy')}"]


def write_hybrid_stub_report(output_dir: Path, input_path: Path, strategy_report: Dict[str, object]) -> None:
    thesis_ir = read_json(thesis_ir_output_path(output_dir)) if thesis_ir_output_path(output_dir).exists() else {}
    asset_anchor_ambiguities = thesis_ir.get("assetAnchorAmbiguities") or []
    attachable_asset_candidates = thesis_ir.get("attachableAssetCandidates") or []
    reattach_candidates = [item for item in attachable_asset_candidates if item.get("recommendedAction") == "reattach_candidate"]
    manual_review_candidates = [item for item in attachable_asset_candidates if item.get("recommendedAction") == "manual_review"]
    payload = {
        "source": str(input_path),
        "chosenStrategy": strategy_report.get("chosenStrategy"),
        "executionMode": strategy_report.get("executionMode"),
        "status": "phase2_text_rebuild_backend",
        "currentBackend": "thesis_ir_text_rebuild_then_conservative_repair",
        "intermediateDocx": str(hybrid_source_docx_path(output_dir)),
        "assetAugmentedDocx": str(hybrid_attached_docx_path(output_dir)),
        "attachableAssetCandidateCount": len(attachable_asset_candidates),
        "reattachCandidateCount": len(reattach_candidates),
        "manualReviewAssetCount": len(manual_review_candidates),
        "assetAnchorAmbiguityCount": len(asset_anchor_ambiguities),
        "notes": [
            "Hybrid rebuild has been selected explicitly by the strategy layer.",
            "Phase 2 rebuilds a clean intermediate DOCX from ThesisIR before invoking the existing conservative DOCX repair backend.",
            "Preserved asset reattachment is still pending; asset-anchor ambiguities remain explicit review inputs instead of being silently guessed.",
        ],
    }
    hybrid_report_path(output_dir).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def refresh_hybrid_execution_report(output_dir: Path, input_path: Path, strategy_report: Dict[str, object]) -> None:
    thesis_ir = read_json(thesis_ir_output_path(output_dir)) if thesis_ir_output_path(output_dir).exists() else {}
    asset_anchor_ambiguities = thesis_ir.get("assetAnchorAmbiguities") or []
    attachable_asset_candidates = thesis_ir.get("attachableAssetCandidates") or []
    candidate_summary = thesis_ir.get("attachableAssetCandidateSummary") or {}
    attachment_report = read_json(hybrid_attachment_report_path(output_dir)) if hybrid_attachment_report_path(output_dir).exists() else {}
    payload = {
        "source": str(input_path),
        "chosenStrategy": strategy_report.get("chosenStrategy"),
        "executionMode": strategy_report.get("executionMode"),
        "status": "phase2_executed",
        "currentBackend": "thesis_ir_text_rebuild_then_conservative_repair",
        "intermediateDocx": str(hybrid_source_docx_path(output_dir)),
        "assetAugmentedDocx": str(hybrid_attached_docx_path(output_dir)),
        "attachableAssetCandidateCount": len(attachable_asset_candidates),
        "candidateActionCounts": (candidate_summary.get("actionCounts") if isinstance(candidate_summary, dict) else {}),
        "attachedAssetCount": attachment_report.get("attachedCount"),
        "manualReviewAssetCount": attachment_report.get("manualReviewCount"),
        "skippedByPolicyCount": attachment_report.get("skippedByPolicyCount"),
        "attachmentSkippedCount": attachment_report.get("skippedCount"),
        "assetAnchorAmbiguityCount": len(asset_anchor_ambiguities),
        "attachmentReport": str(hybrid_attachment_report_path(output_dir)) if hybrid_attachment_report_path(output_dir).exists() else None,
        "notes": [
            "Hybrid rebuild executed through ThesisIR intermediate DOCX generation, asset candidate application, and conservative repair validation.",
            "Attached, manual-review, and policy-skipped assets are reported explicitly.",
            "Low-confidence or unresolved anchors remain deferred instead of being guessed.",
        ],
    }
    hybrid_report_path(output_dir).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def docx_commands(input_path: Path, profile: ProfileConfig, mode: str, output_dir: Path, check_policy: Optional[CheckPolicy] = None, *, thesis_ir: Optional[Path] = None, allow_structural_rebuild: bool = False) -> List[List[str]]:
    return repair_commands(PYTHON, input_path, profile, mode, output_dir, check_policy or CheckPolicy(), thesis_ir=thesis_ir, allow_structural_rebuild=allow_structural_rebuild)


def text_commands(input_path: Path, profile: ProfileConfig, output_dir: Path, check_policy: Optional[CheckPolicy] = None) -> List[List[str]]:
    build = [
        PYTHON,
        "scripts/build_docx_from_markdown.py",
        str(input_path),
        "--output-dir",
        str(output_dir),
        "--keep-source-docx",
        "--build-only",
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
    ]
    if profile.front_matter_template:
        build.extend(["--template-docx", str(profile.front_matter_template)])
    build.extend(["--rules-yaml", str(profile.rules_yaml)])
    commands = [build]
    source_docx = output_dir / "markdown_source.docx"
    commands.extend(docx_commands(source_docx, profile, "conservative-repair", output_dir, check_policy, allow_structural_rebuild=True))
    commands[-1].extend(["--thesis-ir-json", str(thesis_ir_output_path(output_dir))])
    return commands


def hybrid_docx_commands(input_path: Path, profile: ProfileConfig, output_dir: Path, check_policy: Optional[CheckPolicy] = None) -> List[List[str]]:
    source_docx = hybrid_source_docx_path(output_dir)
    attached_docx = hybrid_attached_docx_path(output_dir)
    build = [
        PYTHON,
        "scripts/build_docx_from_thesis_ir.py",
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
        "--rules-yaml",
        str(profile.rules_yaml),
        "--output-docx",
        str(source_docx),
    ]
    attach = [
        PYTHON,
        "scripts/apply_hybrid_asset_candidates.py",
        "--source-docx",
        str(input_path),
        "--intermediate-docx",
        str(source_docx),
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
        "--output-docx",
        str(attached_docx),
        "--report-json",
        str(hybrid_attachment_report_path(output_dir)),
    ]
    commands = [build, attach]
    commands.extend(docx_commands(attached_docx, profile, "conservative-repair", output_dir, check_policy, allow_structural_rebuild=True))
    return commands


W_XML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_XML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def count_docx_assets(docx_path: Path) -> Optional[Dict[str, int]]:
    """Count preservable assets for before/after comparison."""
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = zf.namelist()
            root = ET.fromstring(zf.read("word/document.xml"))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return None
    text_chars = sum(len(node.text or "") for node in root.iter(f"{{{W_XML_NS}}}t"))
    return {
        "mediaFiles": sum(1 for name in names if name.startswith("word/media/")),
        "tables": sum(1 for _ in root.iter(f"{{{W_XML_NS}}}tbl")),
        "drawings": sum(1 for _ in root.iter(f"{{{W_XML_NS}}}drawing")),
        "equations": sum(1 for _ in root.iter(f"{{{M_XML_NS}}}oMath")),
        "textChars": text_chars,
    }


def build_asset_preservation_report(input_path: Path, output_dir: Path) -> Optional[Dict[str, object]]:
    repaired = output_dir / "repaired.docx"
    if input_path.suffix.lower() != ".docx" or not repaired.exists():
        return None
    before = count_docx_assets(input_path)
    after = count_docx_assets(repaired)
    if before is None or after is None:
        return None
    warnings: List[str] = []
    for key, label in (("mediaFiles", "图片等媒体文件"), ("tables", "表格"), ("equations", "公式"), ("drawings", "图形对象")):
        b, a = before.get(key, 0), after.get(key, 0)
        if b and a < b:
            warnings.append(f"{label}数量减少: {b} -> {a}")
    if before.get("textChars", 0) and after.get("textChars", 0) < before["textChars"] * 0.6:
        warnings.append(f"正文字符量大幅下降: {before['textChars']} -> {after['textChars']}")
    attachment_report_path = hybrid_attachment_report_path(output_dir)
    manual_review_assets = 0
    if warnings and attachment_report_path.exists():
        attachment_report = read_json(attachment_report_path)
        manual_review_assets = int(attachment_report.get("manualReviewCount") or 0)
        if manual_review_assets:
            warnings.append(
                f"其中 {manual_review_assets} 个资产因锚点不确定进入人工检查清单(未静默丢弃,见 debug/hybrid_attachment_report.json)"
            )
    return {
        "before": before,
        "after": after,
        "warnings": warnings,
        "manualReviewAssets": manual_review_assets,
        "passed": not warnings,
    }


DEBUG_ARTIFACT_NAMES = [
    "source_evidence.json",
    "numbering_recovery.json",
    "thesis_ir.json",
    "thesis_ir_validation.json",
    "strategy_selection.json",
    "citation_conversion_plan.json",
    "preflight_report.json",
    "source_preflight_report.json",
    "audit_report.json",
    "repair_plan.json",
    "repair_execution.json",
    "validation_report.json",
    "review_summary.json",
    "dispatch_manifest.json",
    "effective_rules.yaml",
    "hybrid_rebuild_report.json",
    "hybrid_attachment_report.json",
    "visual_refinement_execution.json",
    "template_baseline.json",
    "asset_preservation.json",
    "template_regions.json",
]


def run_visual_refinement_pass(output_dir: Path, plan_path: Path) -> Optional[Dict[str, object]]:
    """Refine once; the dispatcher revalidates the final candidate afterward."""
    repaired = output_dir / "repaired.docx"
    if not repaired.exists() or not plan_path.exists():
        return None
    pass1 = output_dir / "repaired_pass1.docx"
    shutil.copy2(repaired, pass1)
    execution_report = output_dir / "visual_refinement_execution.json"
    refine_cmd = [
        PYTHON,
        "scripts/apply_visual_refinements.py",
        str(pass1),
        str(repaired),
        "--plan-json",
        str(plan_path),
        "--report-json",
        str(execution_report),
    ]
    completed = subprocess.run(refine_cmd, cwd=str(ROOT), check=False)
    if completed.returncode != 0:
        # Refinement must never destroy the deliverable; restore pass1.
        shutil.copy2(pass1, repaired)
        return {"status": "failed", "restoredPass1": True}
    payload = read_json(execution_report) if execution_report.exists() else {}
    payload["status"] = "ok"
    payload["pass1Docx"] = str(pass1)
    return payload


def finalize_outputs(output_dir: Path, input_path: Path, workspace: Workspace, manifest: Dict[str, object]) -> Dict[str, object]:
    deliverables: Dict[str, object] = {}
    relocations: Dict[str, str] = {}
    repaired = output_dir / "repaired.docx"
    gate = manifest.get("qualityGate") or {}
    publish_word = gate.get("passed") is not False
    if repaired.exists() and publish_word:
        deliverables["docx"] = str(workspace.publish(repaired, "word", "docx"))
        relocations[str(repaired)] = deliverables["docx"]
        manifest.setdefault("artifacts", {})["repairedDocx"] = deliverables["docx"]

    render_report = read_json(render_report_path(output_dir)) if render_report_path(output_dir).exists() else {}
    pdf_source = render_report.get("pdfPath")
    if isinstance(pdf_source, str) and Path(pdf_source).exists() and publish_word:
        deliverables["pdf"] = str(workspace.publish(Path(pdf_source), "word", "pdf"))
        relocations[pdf_source] = deliverables["pdf"]
        manifest.setdefault("artifacts", {})["repairedPdf"] = deliverables["pdf"]

    if render_report.get("contactSheets"):
        deliverables["contactSheets"] = render_report["contactSheets"]

    critical_src = render_output_dir(output_dir) / "critical_pages"
    if critical_src.exists():
        deliverables["criticalPagesDir"] = str(critical_src)

    deliverables["finalDir"] = str(workspace.deliverables)
    deliverables["reportsDir"] = str(workspace.reports)
    for name in DEBUG_ARTIFACT_NAMES:
        if name in {"source_evidence.json", "numbering_recovery.json", "thesis_ir.json", "dispatch_manifest.json", "effective_rules.yaml"}:
            continue
        source = output_dir / name
        if source.is_file():
            destination = workspace.move(source, workspace.reports / name)
            relocations[str(source)] = str(destination)
            for key, value in (manifest.get("artifacts") or {}).items():
                if value == str(source):
                    manifest["artifacts"][key] = str(destination)
    render_source = render_report_path(output_dir)
    if render_source.is_file():
        render_destination = workspace.move(render_source, workspace.reports / "render_validation.json")
        relocations[str(render_source)] = str(render_destination)
        manifest.setdefault("artifacts", {})["renderValidationReport"] = str(render_destination)
    workspace.rewrite_references(relocations)
    manifest.update(workspace.remap(manifest, relocations))
    write_manifest(output_dir, manifest, manifest.get("commands") or [])
    if manifest.get("backend") != "latex":
        report_html = workspace.reports / "summary.html"
        run_step([PYTHON, "scripts/generate_review_report.py", "--output-dir", str(output_dir), "--report-html", str(report_html)], ROOT)
        deliverables["reviewReport"] = str(report_html)
    return deliverables


def finish_run(args: argparse.Namespace, workspace: Workspace, manifest: Dict[str, object], commands: List[List[str]], latex_result: Optional[Dict[str, object]] = None) -> None:
    update_manifest_with_outputs(manifest, workspace.source, workspace.work)
    manifest["commands"] = commands
    deliverables = finalize_outputs(workspace.work, workspace.source, workspace, manifest)
    if args.backend != "latex":
        manifest.setdefault("backendResults", {})["word"] = {
            "status": "ok" if deliverables.get("docx") else "audit_only" if manifest.get("effectiveMode") == "audit-only" else "blocked",
            "qualityGate": manifest.get("qualityGate"), "visualReviewed": False,
        }
    if latex_result is not None:
        deliverables.update(latex_result.get("deliverables") or {})
        compilation = latex_result.get("compilation") or {}
        manifest.setdefault("backendResults", {})["latex"] = {
            "conversionPassed": latex_result["passed"], "compilation": compilation,
            "profileCompliance": latex_result["profileCompliance"], "errors": latex_result["errors"],
            "warnings": latex_result["warnings"], "visualReviewed": False,
        }
        manifest.setdefault("artifacts", {})["latexReport"] = latex_result["reportPath"]
        word_gate = manifest.get("qualityGate") or {}
        compile_failed = compilation.get("status") in {"failed", "blocked"}
        manifest["qualityGate"] = {
            "passed": word_gate.get("passed", True) and latex_result["passed"] and not compile_failed,
            "hardFailureCount": int(word_gate.get("hardFailureCount") or 0) + len(latex_result["errors"]) + int(compile_failed),
            "warningCount": int(word_gate.get("warningCount") or 0) + len(latex_result["warnings"]) + len(compilation.get("warnings") or []) + int(compilation.get("status") in {"skipped", "unavailable"}),
        }
        if not latex_result["passed"] or compilation.get("status") in {"failed", "blocked"}:
            manifest["status"] = "completed_with_hard_failures"
        elif compilation.get("status") != "ok" and manifest.get("status") == "completed":
            manifest["status"] = "completed_source_only" if args.backend == "latex" else "completed_with_warnings"
    manifest["deliverables"] = deliverables
    if args.backend != "latex" and thesis_ir_output_path(workspace.work).is_file():
        reference_report = check_reference_format(read_json(thesis_ir_output_path(workspace.work)))
        reference_path = workspace.reports / "reference-format.json"
        reference_path.write_text(json.dumps(reference_report, ensure_ascii=False, indent=2), encoding="utf-8")
        manifest.setdefault("artifacts", {})["referenceFormatReport"] = str(reference_path)
        manifest["backendResults"]["word"]["referenceFormatting"] = {"entryCount": reference_report["entryCount"], "manualCheckCount": len(reference_report["issues"]), "citationManagement": False}
    summary_path = write_summary(workspace, manifest, deliverables)
    manifest.setdefault("artifacts", {})["summary"] = summary_path
    write_manifest(workspace.work, manifest, commands)
    print(json.dumps({"status": manifest["status"], "outputDir": str(workspace.root), "manifest": str(workspace.manifest_path), "finalDir": str(workspace.deliverables)}, ensure_ascii=False))
    print_cli_summary(manifest, deliverables, workspace.work)


def print_cli_summary(manifest: Dict[str, object], deliverables: Dict[str, object], output_dir: Path) -> None:
    quality_gate = manifest.get("qualityGate") or {}
    lines = ["", "=" * 46, "论文排版完成"]
    if deliverables.get("docx"):
        lines.append(f"  Word 文件: {deliverables['docx']}")
    else:
        lines.append("  Word 文件: 未交付(见 reports 报告或后端选择)")
    if deliverables.get("pdf"):
        lines.append(f"  PDF 预览: {deliverables['pdf']}")
    else:
        lines.append("  Word PDF: 未交付(默认不导出,或转换不可用/校验未通过)")
    if deliverables.get("latexSource"):
        lines.append(f"  LaTeX 工程: {deliverables['latexSource']}")
    if deliverables.get("latexPdf"):
        lines.append(f"  LaTeX PDF: {deliverables['latexPdf']}")
    elif manifest.get("backend") in {"latex", "both"}:
        compilation = ((manifest.get("backendResults") or {}).get("latex") or {}).get("compilation") or {}
        lines.append(f"  LaTeX 编译: {compilation.get('status')} ({compilation.get('reason') or '见报告'})")
    lines.append("  视觉审阅: 未执行;结构校验不代表分页和视觉效果已经通过")
    if deliverables.get("reviewReport"):
        lines.append(f"  检查报告: {deliverables['reviewReport']}")
    sheets = deliverables.get("contactSheets") or []
    if sheets:
        lines.append(f"  全文缩略图: {sheets[0]}")
    risk = (manifest.get("preflightDecision") or {}).get("documentRiskClass")
    strategy = (manifest.get("strategySelection") or {}).get("chosenStrategy")
    lines.append(f"  风险等级: {risk or '—'}    处理策略: {strategy or '—'}")
    asset_report = manifest.get("assetPreservation")
    if isinstance(asset_report, dict):
        if asset_report.get("passed"):
            after = asset_report.get("after") or {}
            lines.append(
                f"  资产保留: 图片/媒体 {after.get('mediaFiles', 0)}、表格 {after.get('tables', 0)}、公式 {after.get('equations', 0)},与源文档一致"
            )
        else:
            for warning in (asset_report.get("warnings") or [])[:3]:
                lines.append(f"  资产警告: {warning}")
    if isinstance(quality_gate, dict) and quality_gate:
        hard = quality_gate.get("hardFailureCount") or 0
        warn = quality_gate.get("warningCount") or 0
        if hard:
            lines.append(f"  质量门: {hard} 个硬性问题、{warn} 个警告(结果仍已生成,请查看报告)")
        else:
            lines.append(f"  质量门: 通过({warn} 个警告,不影响交付)")
    if manifest.get("backend") != "latex":
        lines.append("  提示: 在 Word 中打开后请更新目录域(Ctrl+A → F9),并核对封面个人信息。")
    lines.append("=" * 46)
    print("\n".join(lines))


def write_manifest(output_dir: Path, manifest: Dict[str, object], commands: List[List[str]]) -> None:
    payload = dict(manifest)
    payload["commands"] = commands
    workspace = Workspace(Path(str(manifest["runDir"])), str(manifest["projectName"]), str(manifest["profile"]), Path(str(manifest["input"])))
    payload.setdefault("artifacts", {})["dispatchManifest"] = str(workspace.manifest_path)
    workspace.write_manifest(payload)


def execute(args: argparse.Namespace, workspace: Workspace, manifest: Dict[str, object]) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except (ValueError, OSError):
            pass
    check_policy = CheckPolicy(args.check, args.export_pdf)
    input_path = workspace.source

    profile = PROFILES[args.profile]
    output_dir = workspace.work

    effective_rules_yaml = materialize_effective_rules_yaml(profile, args.compliance, output_dir)
    if args.header_mode or args.header_text is not None or args.latin_font or input_path.suffix.lower() != ".docx":
        effective_rules = read_yaml(effective_rules_yaml)
        page = effective_rules.setdefault("page", {})
        if args.header_mode or args.header_text is not None:
            page["header_text_mode"] = "thesis_title_after_toc" if args.header_mode == "thesis-title" else args.header_mode or "fixed"
        if input_path.suffix.lower() != ".docx":
            page["fresh_document"] = True
            page["create_missing_body_headers"] = True
        if args.latin_font:
            effective_rules["latin_font_override"] = args.latin_font
            page["header_font_ascii"] = args.latin_font
            for spec in (effective_rules.get("styles") or {}).values():
                if isinstance(spec, dict):
                    spec["ascii_font"] = args.latin_font
                    spec["hansi_font"] = args.latin_font
        if args.header_text is not None:
            page["header_text"] = args.header_text
        if page.get("header_text_mode") == "fixed" and not str(page.get("header_text") or "").strip():
            raise ValueError("Fixed header mode requires non-empty header text")
        effective_rules_yaml.write_text(yaml.safe_dump(effective_rules, allow_unicode=True, sort_keys=False), encoding="utf-8")
    profile = ProfileConfig(
        name=profile.name,
        profile_dir=profile.profile_dir,
        rules_yaml=effective_rules_yaml,
        template_docx=profile.template_docx,
        front_matter_template=profile.front_matter_template,
        style_map_yaml=profile.style_map_yaml,
        front_matter_policy_yaml=profile.front_matter_policy_yaml,
        validators_yaml=profile.validators_yaml,
        metadata=profile.metadata,
    )
    manifest.update(build_manifest(input_path, profile, args.mode, output_dir, args.compliance))
    manifest["runDir"] = str(workspace.root)
    manifest["projectName"] = workspace.project
    manifest["backend"] = args.backend
    manifest["latexOptions"] = {"noCompile": args.no_compile, "compileTimeout": args.compile_timeout, "engine": args.latex_engine}
    manifest["checkPolicy"] = check_policy.summary()
    kind = manifest["sourceKind"]
    evidence_cmd = evidence_command(input_path, profile, output_dir)
    numbering_cmd = numbering_command(output_dir)
    ir_cmd = thesis_ir_command(output_dir, args.semantic_overrides)
    ir_validation_cmd = thesis_ir_validation_command(output_dir)
    preflight_cmd = preflight_command(input_path, profile, output_dir)
    pipeline_prefix_commands = [evidence_cmd, numbering_cmd, ir_cmd, ir_validation_cmd, preflight_cmd]
    manifest["wordCitationManagement"] = "disabled; bibliography presentation and non-mutating checks only"
    manifest["wordInputPolicy"] = {
        "mode": "preserve-input" if kind == "docx" else "fresh-from-source",
        "coverSource": str(profile.front_matter_template) if kind != "docx" else "input (unless explicitly authorized rebuild)",
        "coverSourceSha256": file_digest(profile.front_matter_template) if kind != "docx" and profile.front_matter_template else None,
        "latinFontOverride": args.latin_font,
    }

    if args.dry_run:
        manifest["status"] = "dry-run"
        manifest["notes"] = list(manifest.get("notes") or []) + [
            "Dry-run does not execute preflight, so dispatcher decisions are limited to source-type defaults.",
        ]
        commands = list(pipeline_prefix_commands)
        update_manifest_with_outputs(manifest, input_path, output_dir)
        write_manifest(output_dir, manifest, commands)
        print(json.dumps({"manifest": str(workspace.manifest_path), "commandCount": len(commands)}, ensure_ascii=False))
        return

    run_step(evidence_cmd, ROOT)
    run_step(numbering_cmd, ROOT)
    run_step(ir_cmd, ROOT)
    run_step(ir_validation_cmd, ROOT)
    run_step(preflight_cmd, ROOT)
    preflight_report = read_json(preflight_output_path(kind, output_dir))
    effective_mode, decision_reasons = decide_mode(args.mode, preflight_report, str(kind), profile)
    run_step(strategy_command(output_dir, str(kind)), ROOT)
    strategy_report = read_json(strategy_output_path(output_dir))
    manifest["structuralRebuildAllowed"] = args.allow_structural_rebuild
    if kind == "docx" and not args.allow_structural_rebuild:
        manifest["deferredStructuralChanges"] = ["template_front_matter_replacement"]
        if strategy_report.get("chosenStrategy") in {"hybrid_rebuild", "text_rebuild"}:
            manifest["deferredStructuralChanges"].append(str(strategy_report["chosenStrategy"]))
            strategy_report["chosenStrategy"] = "preserve_first"
            strategy_report["executionMode"] = "conservative-repair"
            strategy_report.setdefault("reasons", []).append("structural_rebuild_requires_explicit_permission")
        strategy_output_path(output_dir).write_text(json.dumps(strategy_report, ensure_ascii=False, indent=2), encoding="utf-8")
    effective_mode, strategy_reasons = reconcile_mode_with_strategy(effective_mode, strategy_report)
    manifest["effectiveMode"] = effective_mode
    manifest["effectiveExpertMode"] = effective_mode
    manifest["decisionReasons"] = decision_reasons + strategy_reasons
    manifest["strategySelection"] = strategy_report

    if args.backend == "latex":
        manifest["status"] = "completed"
        manifest["effectiveMode"] = "rebuild"
        result = run_latex(read_json(thesis_ir_output_path(output_dir)), input_path, read_yaml(profile.rules_yaml), profile.profile_dir, workspace, check_policy, no_compile=args.no_compile, timeout=args.compile_timeout, engine_name=args.latex_engine, metadata_path=Path(args.latex_metadata) if args.latex_metadata else None, bibliography_path=Path(args.latex_bibliography) if args.latex_bibliography else None)
        finish_run(args, workspace, manifest, pipeline_prefix_commands + [strategy_command(output_dir, str(kind))], result)
        return

    if kind == "docx":
        chosen_strategy = str(strategy_report.get("chosenStrategy") or "")
        if chosen_strategy == "hybrid_rebuild" and effective_mode != "audit-only":
            write_hybrid_stub_report(output_dir, input_path, strategy_report)
            manifest["decisionReasons"] = list(manifest.get("decisionReasons") or []) + [
                "hybrid_rebuild_uses_thesis_ir_intermediate_docx"
            ]
            commands = hybrid_docx_commands(input_path, profile, output_dir, check_policy)
        else:
            commands = docx_commands(input_path, profile, effective_mode, output_dir, check_policy, thesis_ir=thesis_ir_output_path(output_dir), allow_structural_rebuild=args.allow_structural_rebuild)
        if effective_mode != "audit-only" and preflight_report.get("confirmationRequests"):
            # Deliver-first: confirmation requests become manual-review warnings
            # in the report instead of blocking the whole run.
            manifest["decisionReasons"] = list(manifest.get("decisionReasons") or []) + [
                "confirmation_requests_downgraded_to_manual_review_warnings"
            ]
            manifest["pendingConfirmationRequests"] = list(preflight_report.get("confirmationRequests") or [])[:40]
    else:
        if effective_mode == "audit-only":
            update_manifest_from_preflight(
                manifest,
                preflight_report,
                effective_mode=effective_mode,
                status="stopped_after_preflight",
                stop_reason="text_source_preflight_recommended_audit_only",
            )
            update_manifest_with_outputs(manifest, input_path, output_dir)
            write_manifest(output_dir, manifest, pipeline_prefix_commands)
            print(json.dumps({"status": "stopped_after_preflight", "outputDir": str(output_dir), "mode": effective_mode}, ensure_ascii=False))
            return
        if preflight_report.get("confirmationRequests"):
            manifest["decisionReasons"] = list(manifest.get("decisionReasons") or []) + [
                "confirmation_requests_downgraded_to_manual_review_warnings"
            ]
            manifest["pendingConfirmationRequests"] = list(preflight_report.get("confirmationRequests") or [])[:40]
        commands = text_commands(input_path, profile, output_dir, check_policy)

    update_manifest_from_preflight(manifest, preflight_report, effective_mode=effective_mode, status="running")
    update_manifest_with_outputs(manifest, input_path, output_dir)
    write_manifest(output_dir, manifest, pipeline_prefix_commands + commands)

    for command in commands:
        run_step(command, ROOT)

    repair_plan_path = output_dir / "repair_plan.json"
    if repair_plan_path.is_file():
        manifest["semanticCorrections"] = read_json(repair_plan_path).get("semanticCorrectionReport")
    execution_path = output_dir / "repair_execution.json"
    if execution_path.is_file():
        manifest["formattingReview"] = [
            item for item in read_json(execution_path).get("executionLog") or []
            if item.get("action") in {"header_manual_review", "semantic_correction_needs_review"}
        ]

    if kind == "docx" and effective_mode != "audit-only" and str(strategy_report.get("chosenStrategy") or "") == "hybrid_rebuild":
        refresh_hybrid_execution_report(output_dir, input_path, strategy_report)

    if args.visual_refinement_plan:
        refinement_result = run_visual_refinement_pass(output_dir, Path(args.visual_refinement_plan).resolve())
        if refinement_result is not None:
            manifest["visualRefinement"] = refinement_result
            run_step(commands[-1], ROOT)
            commands.append(list(commands[-1]))

    review_summary = build_review_summary(output_dir, manifest)
    if review_summary is not None:
        (output_dir / "review_summary.json").write_text(json.dumps(review_summary, ensure_ascii=False, indent=2), encoding="utf-8")
        quality_gate = review_summary.get("qualityGate") or {}
        if quality_gate.get("passed"):
            manifest["status"] = "completed"
        else:
            manifest["status"] = "completed_with_hard_failures"
        manifest["qualityGate"] = quality_gate
    else:
        manifest["status"] = "completed"
    update_manifest_with_outputs(manifest, input_path, output_dir)

    asset_report = build_asset_preservation_report(input_path, output_dir)
    if asset_report is not None:
        (output_dir / "asset_preservation.json").write_text(json.dumps(asset_report, ensure_ascii=False, indent=2), encoding="utf-8")
        manifest["assetPreservation"] = asset_report

    if profile.template_docx:
        subprocess.run(
            [
                PYTHON,
                "scripts/analyze_template_regions.py",
                str(profile.template_docx),
                "--output",
                str(output_dir / "template_regions.json"),
            ],
            cwd=str(ROOT),
            check=False,
        )

    latex_result = None
    if args.backend == "both":
        latex_result = run_latex(read_json(thesis_ir_output_path(output_dir)), input_path, read_yaml(profile.rules_yaml), profile.profile_dir, workspace, check_policy, no_compile=args.no_compile, timeout=args.compile_timeout, engine_name=args.latex_engine, metadata_path=Path(args.latex_metadata) if args.latex_metadata else None, bibliography_path=Path(args.latex_bibliography) if args.latex_bibliography else None)
    finish_run(args, workspace, manifest, pipeline_prefix_commands + [strategy_command(output_dir, str(kind))] + commands, latex_result)


def main() -> None:
    args = parse_args()
    source = Path(args.input).resolve()
    if not source.is_file():
        raise SystemExit(f"Input file not found: {source}")
    workspace = Workspace.create(source, args.profile, Path(args.output_dir) if args.output_dir else None, args.project_name)
    manifest: Dict[str, object] = {
        "input": str(source), "profile": args.profile, "runDir": str(workspace.root),
        "projectName": workspace.project, "status": "running", "backend": args.backend,
        "inputSha256": file_digest(source),
    }
    workspace.write_manifest(manifest)
    try:
        execute(args, workspace, manifest)
    except (Exception, SystemExit) as error:
        manifest["status"] = "failed"
        manifest["error"] = {"type": type(error).__name__, "message": str(error)}
        workspace.write_manifest(manifest)
        print(json.dumps({"status": "failed", "manifest": str(workspace.manifest_path)}, ensure_ascii=False), file=sys.stderr)
        raise
    if manifest.get("status") == "completed_with_hard_failures":
        raise SystemExit(2)


if __name__ == "__main__":
    main()
