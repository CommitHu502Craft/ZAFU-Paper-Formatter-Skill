from __future__ import annotations

from typing import Any


def apply_ir_corrections(plan: dict[str, Any], audit: dict[str, Any], ir: dict[str, Any]) -> None:
    targets = plan["profileStyleTargets"]
    role_styles = {
        "body": "body_text",
        "figure_caption": "caption",
        "table_caption": "caption",
        "cn_abstract_heading": "abstract_heading",
        "en_abstract_heading": "abstract_heading",
        "references_heading": "heading1",
        "acknowledgements_heading": "heading1",
        "appendix_heading": "heading1",
        "acknowledgements_body": "body_text",
        "appendix_body": "body_text",
        **{f"heading_{level}": f"heading{min(level, 3)}" for level in range(1, 7)},
    }
    paragraphs = {item["index"]: item for item in audit.get("paragraphs") or []}
    corrections = []
    skipped = list((ir.get("semanticOverrides") or {}).get("rejected") or [])
    for block in ir.get("semanticBlocks") or []:
        if not block.get("overridden"):
            continue
        index = block.get("docxParagraphIndex")
        role = block.get("role")
        target = targets.get(role_styles.get(role, ""))
        reason = None
        if ir.get("sourceType") != "docx" or index not in paragraphs:
            reason = "source_paragraph_not_found"
        elif not target:
            reason = "role_not_supported_by_in_place_repair"
        elif any(
            region.get("startParagraph", 0) <= index < region.get("endParagraphExclusive", 0)
            for region in plan.get("preserveRegions") or []
        ):
            reason = "preserved_front_matter"
        if reason:
            skipped.append({"blockId": block.get("id"), "reason": reason})
            continue
        plan["styleMapping"][str(index)] = target
        plan["paragraphRoles"][str(index)] = {"role": role, "confidence": 0.99, "signals": ["explicit_ir_correction"]}
        plan["paragraphActions"] = [item for item in plan.get("paragraphActions") or [] if item.get("paragraphIndex") != index]
        plan["numberingActions"] = [item for item in plan.get("numberingActions") or [] if item.get("paragraphIndex") != index]
        for action in plan.get("actions") or []:
            if action.get("paragraphIndex") == index:
                action["supersededBySemanticCorrection"] = True
        corrections.append({"blockId": block.get("id"), "paragraphIndex": index, "role": role, "styleId": target})
    plan["semanticCorrections"] = corrections
    plan["semanticCorrectionReport"] = {"applied": corrections, "skipped": skipped}
    plan.setdefault("manualReview", []).extend({"issue": "semantic_correction_skipped", **item} for item in skipped)
