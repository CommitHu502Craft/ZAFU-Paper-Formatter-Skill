#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import yaml


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
    parser.add_argument("--output-dir", default="output", help="Directory for reports and generated files")
    parser.add_argument(
        "--compliance",
        default="default",
        choices=["default", "strict-school"],
        help="Compliance profile overlay; strict-school prefers template-conformant repair defaults over source-preserving heuristics.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Write the dispatch manifest only")
    return parser.parse_args()


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
    if include_template and profile.template_docx:
        command.extend(["--template-docx", str(profile.template_docx)])
    command.extend(["--rules-yaml", str(profile.rules_yaml)])
    if profile.style_map_yaml:
        command.extend(["--style-map-yaml", str(profile.style_map_yaml)])
    if profile.front_matter_policy_yaml:
        command.extend(["--front-matter-policy-yaml", str(profile.front_matter_policy_yaml)])
    if include_validators and profile.validators_yaml:
        command.extend(["--validators-yaml", str(profile.validators_yaml)])
    return command


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


def thesis_ir_command(output_dir: Path) -> List[str]:
    return [PYTHON, "scripts/thesis_ir.py", "--evidence-json", str(evidence_output_path(output_dir)), "--output", str(thesis_ir_output_path(output_dir))]


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
    artifacts: Dict[str, object] = {
        "preflightReport": str(preflight_output_path(str(manifest["sourceKind"]), output_dir)),
        "dispatchManifest": str(output_dir / "dispatch_manifest.json"),
    }
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
        if isinstance(preview_dir, str) and preview_dir:
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
    effective_mode = requested_mode
    reasons: List[str] = []
    recommended_mode = preflight_report.get("recommendedMode")
    document_risk_class = preflight_report.get("documentRiskClass")
    allow_risk_class_c_repair = (
        source_kind_value == "docx"
        and requested_mode != "audit-only"
        and profile_policy_flag(profile, "allow_conservative_repair_for_risk_class_c", False)
    )

    if source_kind_value == "text" and requested_mode == "conservative-repair":
        effective_mode = "rebuild"
        reasons.append("text_source_conservative_repair_upgraded_to_rebuild")
    if document_risk_class == "C" and effective_mode != "audit-only":
        if allow_risk_class_c_repair:
            reasons.append("risk_class_c_kept_repair_mode_due_to_profile_policy")
        else:
            effective_mode = "audit-only"
            reasons.append("risk_class_c_forces_audit_only")
    elif recommended_mode == "audit-only" and effective_mode != "audit-only":
        if allow_risk_class_c_repair and document_risk_class == "C":
            reasons.append("preflight_recommended_audit_only_overridden_by_profile_policy")
        else:
            effective_mode = "audit-only"
            reasons.append("preflight_recommended_audit_only")
    elif source_kind_value == "docx" and recommended_mode in {"conservative-repair", "rebuild"} and requested_mode == "audit-only":
        reasons.append("user_requested_audit_only")
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


def docx_commands(input_path: Path, profile: ProfileConfig, mode: str, output_dir: Path) -> List[List[str]]:
    inspect = [PYTHON, "scripts/inspect_docx.py", str(input_path), "--output", str(output_dir / "audit_report.json")]
    plan = [PYTHON, "scripts/plan_docx_repairs.py", str(input_path), "--output", str(output_dir / "repair_plan.json")]
    validate_audit = [
        PYTHON,
        "scripts/validate_docx.py",
        str(input_path),
        "--output",
        str(output_dir / "validation_report.json"),
        "--render-output-dir",
        str(render_output_dir(output_dir)),
    ]

    commands = [add_common_flags(inspect, profile)]
    if mode == "audit-only":
        commands.append(add_common_flags(validate_audit, profile))
        return commands

    repaired = output_dir / "repaired.docx"
    apply_cmd = [
        PYTHON,
        "scripts/apply_ooxml_fixes.py",
        str(input_path),
        str(repaired),
        "--plan-json",
        str(output_dir / "repair_plan.json"),
        "--report-json",
        str(output_dir / "repair_execution.json"),
    ]
    validate_repaired = [
        PYTHON,
        "scripts/validate_docx.py",
        str(repaired),
        "--before-docx",
        str(input_path),
        "--plan-json",
        str(output_dir / "repair_plan.json"),
        "--output",
        str(output_dir / "validation_report.json"),
        "--render-output-dir",
        str(render_output_dir(output_dir)),
    ]
    commands.extend(
        [
            add_common_flags(plan, profile, include_validators=False),
            add_common_flags(apply_cmd, profile),
            add_common_flags(validate_repaired, profile),
        ]
    )
    return commands


def text_commands(input_path: Path, profile: ProfileConfig, output_dir: Path) -> List[List[str]]:
    build = [
        PYTHON,
        "scripts/build_docx_from_markdown.py",
        str(input_path),
        "--output-dir",
        str(output_dir),
        "--keep-source-docx",
        "--thesis-ir-json",
        str(thesis_ir_output_path(output_dir)),
    ]
    if profile.front_matter_template:
        build.extend(["--template-docx", str(profile.front_matter_template)])
    build.extend(["--rules-yaml", str(profile.rules_yaml)])
    commands = [build]
    source_docx = output_dir / "markdown_source.docx"
    commands.extend(docx_commands(source_docx, profile, "conservative-repair", output_dir))
    return commands


def hybrid_docx_commands(input_path: Path, profile: ProfileConfig, output_dir: Path) -> List[List[str]]:
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
    commands.extend(docx_commands(attached_docx, profile, "conservative-repair", output_dir))
    return commands


def write_manifest(output_dir: Path, manifest: Dict[str, object], commands: List[List[str]]) -> None:
    payload = dict(manifest)
    payload["commands"] = commands
    (output_dir / "dispatch_manifest.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {input_path}")

    profile = PROFILES[args.profile]
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    effective_rules_yaml = materialize_effective_rules_yaml(profile, args.compliance, output_dir)
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
    manifest = build_manifest(input_path, profile, args.mode, output_dir, args.compliance)
    kind = manifest["sourceKind"]
    evidence_cmd = evidence_command(input_path, profile, output_dir)
    numbering_cmd = numbering_command(output_dir)
    ir_cmd = thesis_ir_command(output_dir)
    ir_validation_cmd = thesis_ir_validation_command(output_dir)
    preflight_cmd = preflight_command(input_path, profile, output_dir)
    citation_cmd = citation_conversion_command(profile, output_dir)
    pipeline_prefix_commands = [evidence_cmd, numbering_cmd, ir_cmd, ir_validation_cmd, citation_cmd, preflight_cmd]

    if args.dry_run:
        manifest["status"] = "dry-run"
        manifest["notes"] = list(manifest.get("notes") or []) + [
            "Dry-run does not execute preflight, so dispatcher decisions are limited to source-type defaults.",
        ]
        commands = list(pipeline_prefix_commands)
        update_manifest_with_outputs(manifest, input_path, output_dir)
        write_manifest(output_dir, manifest, commands)
        print(json.dumps({"manifest": str(output_dir / "dispatch_manifest.json"), "commandCount": len(commands)}, ensure_ascii=False))
        return

    run_step(evidence_cmd, ROOT)
    run_step(numbering_cmd, ROOT)
    run_step(ir_cmd, ROOT)
    run_step(ir_validation_cmd, ROOT)
    run_step(citation_cmd, ROOT)
    run_step(preflight_cmd, ROOT)
    preflight_report = read_json(preflight_output_path(kind, output_dir))
    effective_mode, decision_reasons = decide_mode(args.mode, preflight_report, str(kind), profile)
    run_step(strategy_command(output_dir, str(kind)), ROOT)
    strategy_report = read_json(strategy_output_path(output_dir))
    effective_mode, strategy_reasons = reconcile_mode_with_strategy(effective_mode, strategy_report)
    manifest["effectiveMode"] = effective_mode
    manifest["effectiveExpertMode"] = effective_mode
    manifest["decisionReasons"] = decision_reasons + strategy_reasons
    manifest["strategySelection"] = strategy_report

    if kind == "docx":
        chosen_strategy = str(strategy_report.get("chosenStrategy") or "")
        if chosen_strategy == "hybrid_rebuild":
            write_hybrid_stub_report(output_dir, input_path, strategy_report)
            manifest["decisionReasons"] = list(manifest.get("decisionReasons") or []) + [
                "hybrid_rebuild_uses_thesis_ir_intermediate_docx"
            ]
            commands = hybrid_docx_commands(input_path, profile, output_dir)
        else:
            commands = docx_commands(input_path, profile, effective_mode, output_dir)
        if effective_mode != "audit-only" and preflight_report.get("confirmationRequests"):
            if profile_policy_flag(profile, "continue_safe_subset_repair_when_confirmation_requests_present", False):
                manifest["decisionReasons"] = list(manifest.get("decisionReasons") or []) + [
                    "docx_safe_subset_repair_continues_despite_confirmation_requests"
                ]
            else:
                update_manifest_from_preflight(
                    manifest,
                    preflight_report,
                    effective_mode=effective_mode,
                    status="stopped_for_confirmation",
                    stop_reason="preflight_confirmation_requests_block_automatic_repair",
                )
                update_manifest_with_outputs(manifest, input_path, output_dir)
                write_manifest(output_dir, manifest, pipeline_prefix_commands)
                print(
                    json.dumps(
                        {
                            "status": "stopped_for_confirmation",
                            "outputDir": str(output_dir),
                            "mode": effective_mode,
                            "confirmationRequestCount": len(preflight_report.get("confirmationRequests") or []),
                        },
                        ensure_ascii=False,
                    )
                )
                return
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
            update_manifest_from_preflight(
                manifest,
                preflight_report,
                effective_mode=effective_mode,
                status="stopped_for_confirmation",
                stop_reason="text_source_confirmation_requests_block_build",
            )
            update_manifest_with_outputs(manifest, input_path, output_dir)
            write_manifest(output_dir, manifest, pipeline_prefix_commands)
            print(
                json.dumps(
                    {
                        "status": "stopped_for_confirmation",
                        "outputDir": str(output_dir),
                        "mode": effective_mode,
                        "confirmationRequestCount": len(preflight_report.get("confirmationRequests") or []),
                    },
                    ensure_ascii=False,
                )
            )
            return
        commands = text_commands(input_path, profile, output_dir)

    update_manifest_from_preflight(manifest, preflight_report, effective_mode=effective_mode, status="running")
    update_manifest_with_outputs(manifest, input_path, output_dir)
    write_manifest(output_dir, manifest, pipeline_prefix_commands + commands)

    for command in commands:
        run_step(command, ROOT)

    if kind == "docx" and str(strategy_report.get("chosenStrategy") or "") == "hybrid_rebuild":
        refresh_hybrid_execution_report(output_dir, input_path, strategy_report)

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
    write_manifest(output_dir, manifest, pipeline_prefix_commands + commands)
    print(
        json.dumps(
            {
                "status": "ok",
                "outputDir": str(output_dir),
                "productMode": PRODUCT_MODE,
                "expertMode": effective_mode,
                "qualityGatePassed": ((manifest.get("qualityGate") or {}).get("passed") if manifest.get("qualityGate") else None),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
