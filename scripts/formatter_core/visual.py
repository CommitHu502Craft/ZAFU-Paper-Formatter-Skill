"""Candidate-bound visual evidence, unique anchors and a single batch repair."""
from __future__ import annotations
import json
import re
import shutil
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from formatter_core.workspace import file_digest
from formatter_core.session import run_in_process

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def squash(value):
    return re.sub(r"\s+", "", value or "")

def anchor_findings(docx, findings, text_boxes, ir=None):
    with zipfile.ZipFile(docx) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    paragraph_nodes = list(root.iter(W + "p"))
    paragraphs = ["".join(t.text or "" for t in p.iter(W + "t")) for p in paragraph_nodes]
    normalized = [squash(text) for text in paragraphs]
    blocks = (ir or {}).get("semanticBlocks") or []
    for finding in findings:
        finding["pageNumber"] = finding.get("page")
        finding["issueType"] = finding.get("type")
        finding.setdefault("confidence", 0.65)
        finding.setdefault("classification", "advisory")
        text = finding.get("headingText") or finding.get("captionText") or finding.get("text") or ""
        token = squash(text)
        matches = [i for i, candidate in enumerate(normalized)
                   if (len(token) >= 3 and token == candidate)
                   or (len(token) >= 4 and token in candidate)]
        finding["mappingStatus"] = "unique" if len(matches) == 1 else "ambiguous" if matches else "unmapped"
        finding["candidateParagraphIndexes"] = matches
        if len(matches) == 1:
            index = matches[0]
            finding["docxParagraphIndex"] = index
            finding["objectId"] = f"paragraph-{index}"
            finding["textPrefix"] = paragraphs[index][:40]
            style = paragraph_nodes[index].find(W + "pPr/" + W + "pStyle")
            if finding.get("type") == "heading_orphan_at_page_bottom_geometry" and style is not None and "heading" in style.get(W + "val", "").lower():
                finding["confidence"] = 0.92
            candidates = [b for b in blocks if squash(b.get("text")) == normalized[index]]
            finding["blockId"] = candidates[0].get("id") if len(candidates) == 1 else None
        page = finding.get("page")
        fragments = text_boxes[page - 1].get("fragments", []) if isinstance(page, int) and 0 < page <= len(text_boxes) else []
        regions = [f for f in fragments if token and (token in squash(f.get("text")) or squash(f.get("text")) in token)]
        if len(regions) == 1:
            fragment = regions[0]
            finding["boundingBox"] = fragment.get("boundingBox")
        finding["checkSuggestion"] = "核对该页高清图中的标题与正文连接、题注归属和版心；仅在唯一锚点且证据充分时修改。"
    return findings

def structural_layout_findings(docx, page_texts):
    """Cross-page evidence from unique text anchors; uncertain associations stay advisory."""
    from docx_ooxml import NS
    with zipfile.ZipFile(docx) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    pages = [squash(text) for text in page_texts]
    findings = []
    for table_index, table in enumerate(root.iter(W + "tbl")):
        anchors = [squash("".join(t.text or "" for t in p.iter(W + "t"))) for p in table.iter(W + "p")]
        located = {}
        for token in set(anchors):
            if len(token) < 6:
                continue
            matches = [i + 1 for i, text in enumerate(pages) if token in text]
            if len(matches) == 1:
                located[matches[0]] = located.get(matches[0], 0) + 1
        if len(located) > 1 and min(located.values()) >= 2 and table.find(".//w:tblHeader", NS) is None:
            findings.append({"type": "table_spans_pages_without_repeat_header", "page": min(located), "pages": sorted(located), "tableIndex": table_index, "objectId": f"table-{table_index}", "confidence": 0.9, "classification": "advisory", "severity": "warning"})
    for index, paragraph in enumerate(root.iter(W + "p")):
        page_breaks = [b for b in paragraph.iter(W + "br") if b.get(W + "type") == "page"]
        if len(page_breaks) > 1:
            text = "".join(t.text or "" for t in paragraph.iter(W + "t"))
            matches = [i + 1 for i, page in enumerate(pages) if len(squash(text)) >= 4 and squash(text) in page]
            findings.append({"type": "duplicate_explicit_page_breaks", "page": matches[0] if len(matches) == 1 else None, "text": text, "docxParagraphIndex": index, "confidence": 1.0, "classification": "deterministic_structure", "severity": "warning"})
    return findings

def batch_plan(findings, candidate_hash):
    actions, manual, seen = [], [], set()
    for item in sorted(findings, key=lambda f: (f.get("severity") != "error", f.get("pageNumber") or 0)):
        index = item.get("docxParagraphIndex")
        kind = item.get("issueType") or item.get("type")
        if kind == "heading_orphan_at_page_bottom_geometry" and item.get("mappingStatus") == "unique" and item.get("confidence", 0) >= 0.9:
            key = ("set_keep_with_next", index)
            if key not in seen:
                actions.append({"type": key[0], "target": {"paragraphIndex": index, "textPrefix": item["textPrefix"]}, "confidence": item["confidence"], "reason": kind})
                seen.add(key)
        else:
            manual.append(item)
    return {"candidateSha256": candidate_hash, "actions": actions, "unmappedFindings": manual, "maxRepairRounds": 1}

def merge_actions(plan):
    """Conflicting values on the same property are deferred, never order-dependent."""
    groups, deferred = {}, []
    opposites = {"clear_page_break_before": "pageBreak", "set_page_break_before": "pageBreak"}
    for action in plan.get("actions") or []:
        key = (json.dumps(action.get("target") or {}, sort_keys=True), opposites.get(action.get("type"), action.get("type")))
        groups.setdefault(key, []).append(action)
    accepted = []
    for group in groups.values():
        variants = {json.dumps({k: v for k, v in a.items() if k not in {"reason", "confidence"}}, sort_keys=True) for a in group}
        if len(variants) > 1:
            deferred.extend(group)
        else:
            accepted.append(group[0])
    return {**plan, "actions": accepted, "conflictingActions": deferred}

def complete_visual_workflow(args, workspace, manifest, validation_command, api):
    path = workspace.work / "render_validation/render_validation_report.json"
    report = json.loads(path.read_text(encoding="utf-8"))
    candidate = workspace.work / "repaired.docx"
    source_hash = file_digest(candidate)
    findings = report.get("visualFindings") or []
    plan = merge_actions(batch_plan(findings, source_hash))
    if args.visual_refinement_plan:
        supplied = json.loads(Path(args.visual_refinement_plan).read_text(encoding="utf-8"))
        if supplied.get("candidateSha256") != source_hash:
            raise ValueError("Visual repair plan is stale or missing candidateSha256")
        plan = merge_actions(supplied)
    plan_path = workspace.reports / "visual_repair_plan.json"
    plan_path.write_text(json.dumps(plan, ensure_ascii=True, indent=2), encoding="utf-8")
    rounds = 0
    if plan["actions"] and report.get("status") == "ok":
        from apply_visual_refinements import apply_plan
        from formatter_core.integrity import verify
        backup = workspace.work / "repaired_pass1.docx"
        shutil.copy2(candidate, backup)
        result = apply_plan(backup, candidate, plan)
        rounds = 1
        integrity = verify(backup, candidate)
        if not integrity["passed"]:
            shutil.copy2(backup, candidate)
            result["restoredPass1"] = True
        run_in_process(validation_command, api.ROOT)
        final = json.loads((workspace.work / "validation_report.json").read_text(encoding="utf-8"))
        if not final["qualityGate"]["passed"]:
            # Bounded recovery: retain the safe first candidate and its own render.
            shutil.copy2(backup, candidate)
            result["restoredPass1"] = True
            run_in_process(validation_command, api.ROOT)
        manifest["visualRefinement"] = result
        report = json.loads(path.read_text(encoding="utf-8"))
    rendered = report.get("status") == "ok" and bool(report.get("pdfPath")) and bool(report.get("pageCount"))
    preview = bool(report.get("previewGenerated"))
    manifest["visualAcceptance"] = {
        "rendered": rendered, "previewGenerated": preview, "visualReviewed": False,
        "visualReviewStatus": "pending_agent_review" if preview else "pending_manual_review" if rendered else "not_requested" if not api.CheckPolicy(args.check, args.export_pdf).render else "render_unavailable",
        "unresolvedVisualIssues": report.get("visualFindings") or [], "repairRounds": rounds,
        "candidateSha256": file_digest(candidate), "renderBackend": report.get("backend"),
        "requestedBackend": args.render_backend,
    }
    manifest["deliveryStatus"] = "candidate_pending_visual_review" if args.quality == "final" else "draft"
