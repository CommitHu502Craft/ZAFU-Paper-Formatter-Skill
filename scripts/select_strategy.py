#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List


def select_strategy(thesis_ir: Dict[str, Any], preflight: Dict[str, Any]) -> Dict[str, Any]:
    source_type = str(thesis_ir.get("sourceType") or "")
    preflight_recommended = str(preflight.get("recommendedMode") or "")
    candidate = str(thesis_ir.get("strategyCandidate") or "audit_only")
    conflicts = thesis_ir.get("numberingAnalysis", {}).get("conflicts") or []
    ambiguities = thesis_ir.get("ambiguities") or []
    preserveable_assets = thesis_ir.get("preserveableAssets") or {}
    reference_normalization = (thesis_ir.get("references") or {}).get("normalization") or {}
    structure_confidence = float((thesis_ir.get("confidence") or {}).get("overall") or 0.0)
    document_risk_class = str(preflight.get("documentRiskClass") or "")
    reasons: List[str] = []

    if source_type in {"txt", "md", "markdown"}:
        chosen = "text_rebuild"
        reasons.append("text_source_uses_rebuild_route")
    else:
        chosen = candidate

    if chosen == "preserve_first" and conflicts and (preserveable_assets.get("images") or preserveable_assets.get("tables")):
        chosen = "hybrid_rebuild"
        reasons.append("docx_has_assets_and_numbering_conflicts")
    elif chosen != "audit_only" and structure_confidence < 0.4:
        chosen = "audit_only"
        reasons.append("low_structure_confidence")

    if document_risk_class == "C":
        reasons.append("risk_class_c_requires_mode_gate_outside_strategy")
    if preflight_recommended == "audit-only":
        reasons.append("preflight_recommended_audit_only_requires_mode_gate_outside_strategy")
    if reference_normalization.get("safeListLabelStripCandidate"):
        reasons.append("reference_list_numeric_labels_can_be_stripped_in_rebuild")
    if reference_normalization.get("bodyNumericCitationCount"):
        reasons.append("body_numeric_citations_still_require_manual_mode_conversion")

    execution_mode_map = {
        "preserve_first": "conservative-repair",
        "hybrid_rebuild": "conservative-repair",
        "text_rebuild": "rebuild",
        "audit_only": "audit-only",
    }
    return {
        "sourceType": source_type,
        "strategyCandidate": candidate,
        "chosenStrategy": chosen,
        "executionMode": execution_mode_map[chosen],
        "documentRiskClass": document_risk_class,
        "structureConfidence": structure_confidence,
        "ambiguityCount": len(ambiguities),
        "numberingConflictCount": len(conflicts),
        "referenceNormalization": reference_normalization,
        "reasons": reasons,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Select preserve/rebuild strategy from ThesisIR and preflight output.")
    parser.add_argument("--thesis-ir-json", required=True, help="Path to thesis_ir.json")
    parser.add_argument("--preflight-json", required=True, help="Path to preflight report JSON")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    thesis_ir = json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8"))
    preflight = json.loads(Path(args.preflight_json).read_text(encoding="utf-8"))
    payload = select_strategy(thesis_ir, preflight)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
