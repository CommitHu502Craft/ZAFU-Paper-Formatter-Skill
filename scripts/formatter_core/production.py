"""Formatting-first orchestration, reusing the existing builders and validators."""
from __future__ import annotations
import json
import copy
from pathlib import Path
import time
import zipfile
import xml.etree.ElementTree as ET

from docx_ooxml import NS, qn, parse_document, build_repair_plan, load_rules
from formatter_core.session import run_in_process

def write(path, value):
    path.write_text(json.dumps(value, ensure_ascii=True, indent=2), encoding="utf-8")

def inspect_risks(source):
    """One package scan, no run-level effective-format auditing."""
    with zipfile.ZipFile(source) as package:
        files = {name: package.read(name) for name in package.namelist()}
    root = ET.fromstring(files["word/document.xml"])
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("DOCX has no body")
    risky = []
    protected_tags = {"ins", "del", "moveFrom", "moveTo", "commentRangeStart", "commentRangeEnd", "commentReference", "anchor", "object", "pict", "oMath", "oMathPara", "sdt", "altChunk", "sectPr"}
    for child_index, child in enumerate(body):
        kinds = sorted({node.tag.rsplit("}", 1)[-1] for node in child.iter()} & protected_tags)
        if kinds and child.tag != qn("sectPr"):
            risky.append({"bodyChildIndex": child_index, "kinds": kinds, "reason": "preserve_native_region"})
    return files, root, risky

def normalize_docx(source, destination, audit, plan, rules, *, scan=None):
    """Batch styles; leave risky body children and all native package parts intact."""
    if Path(source).resolve() == Path(destination).resolve():
        raise ValueError("Normalization must not overwrite the original input")
    from apply_ooxml_fixes import (
        style_spec_from_rules, ensure_style, ensure_builtin_toc_styles,
        apply_style_to_paragraph, apply_run_defaults, update_sections,
        normalize_tables, split_first_run_with_prefix, apply_spec_to_run, normalize_oversized_inline_images,
        clear_paragraph_heading_semantics,
    )
    from formatter_core.word_headers import apply_header_policy, initialize_fresh_footers
    files, root, risky = scan or inspect_risks(source)
    body = root.find("w:body", NS)
    protected = {item["bodyChildIndex"] for item in risky}
    # Revision/comment ranges can span multiple paragraphs. Protect the whole range,
    # including otherwise ordinary paragraphs between the markers.
    open_ranges = set()
    for index, child in enumerate(body):
        for node in child.iter():
            local = node.tag.rsplit("}", 1)[-1]
            identity = node.get(qn("id"), "")
            if local in {"commentRangeStart", "moveFromRangeStart", "moveToRangeStart"}:
                open_ranges.add((local.replace("Start", ""), identity))
            if open_ranges:
                protected.add(index)
            if local in {"commentRangeEnd", "moveFromRangeEnd", "moveToRangeEnd"}:
                open_ranges.discard((local.replace("End", ""), identity))
    styles = ET.fromstring(files["word/styles.xml"])
    specs = style_spec_from_rules(rules)
    protected_styles = set()
    # Styles referenced by preserved content (and their ancestors) must remain
    # unchanged; create a distinct target style for editable paragraphs instead.
    prefix = int(plan.get("editableStartParagraph") or 0)
    pi = 0
    for ci, child in enumerate(body):
        if ci in protected or child.tag == qn("p") and pi < prefix:
            protected_styles.update(n.get(qn("val")) for n in child.iter(qn("pStyle")))
            protected_styles.update(n.get(qn("val")) for n in child.iter(qn("rStyle")))
        if child.tag == qn("p"):
            pi += 1
    style_nodes = {n.get(qn("styleId")): n for n in styles.findall("w:style", NS)}
    pending = list(protected_styles)
    while pending:
        node = style_nodes.get(pending.pop())
        if node is not None:
            for related in node.findall("w:basedOn", NS) + node.findall("w:link", NS):
                value = related.get(qn("val"))
                if value not in protected_styles:
                    protected_styles.add(value)
                    pending.append(value)
    for key, original_spec in list(specs.items()):
        spec = copy.deepcopy(original_spec)
        if spec["styleId"] in protected_styles:
            spec["styleId"] += "_normalized"
        specs[key] = spec
        ensure_style(styles, spec)
    if not protected_styles.intersection({"TOC1", "TOC2", "TOC3", "TOCHeading"}):
        ensure_builtin_toc_styles(styles, rules)
    prefix = int(plan.get("editableStartParagraph") or 0)
    paragraph_index = 0
    log = []
    for child_index, child in enumerate(body):
        if child.tag == qn("p"):
            index = paragraph_index
            paragraph_index += 1
            if child_index in protected or index < prefix:
                continue
            spec = specs.get((plan.get("styleMapping") or {}).get(str(index)))
            if spec:
                if any(item.get("paragraphIndex") == index and not item.get("role", "").startswith("heading_") for item in plan.get("semanticCorrections") or []):
                    clear_paragraph_heading_semantics(child)
                ppr = child.find("w:pPr", NS)
                if ppr is not None:
                    for tag in ("spacing", "ind", "jc"):
                        node = ppr.find("w:" + tag, NS)
                        if node is not None:
                            ppr.remove(node)
                apply_style_to_paragraph(child, spec)
                # All explicit emphasis, vertical alignment and native objects survive.
                apply_run_defaults(child, spec, preserve_emphasis=True)
                for fonts in child.iter(qn("rFonts")):
                    for attr in list(fonts.attrib):
                        if attr.rsplit("}", 1)[-1].lower().endswith("theme"):
                            fonts.attrib.pop(attr)
                text = "".join(node.text or "" for node in child.iter(qn("t"))).strip()
                for labels, label_style in ((["摘要:", "摘要："], "zafu_abstract_label_cn"), (["关键词:", "关键词："], "zafu_keywords_label_cn"), (["Abstract:", "ABSTRACT:"], "zafu_abstract_label_en"), (["Key words:", "Keywords:"], "zafu_keywords_label_en")):
                    for label in labels:
                        if label_style in specs and text.startswith(label):
                            label_run, remainder = split_first_run_with_prefix(child, label)
                            if label_run is not None:
                                apply_spec_to_run(label_run, specs[label_style], preserve_emphasis=False)
                            if remainder is not None and not "".join(t.text or "" for t in remainder.iter(qn("t"))).strip():
                                # The label's trailing space is not academic emphasis.
                                rpr = remainder.find("w:rPr", NS)
                                if rpr is not None:
                                    for bold in list(rpr.findall("w:b", NS)):
                                        rpr.remove(bold)
                log.append({"action": "apply_style_mapping", "paragraphIndex": index, "styleId": spec.get("styleId")})
            # Inline image sizing is deterministic; floating or native object regions
            # are excluded above. A temporary body avoids touching other objects.
            isolated = ET.Element(qn("document"))
            isolated_body = ET.SubElement(isolated, qn("body"))
            isolated_body.append(child)
            normalize_oversized_inline_images(isolated, rules, 0)
        elif child.tag == qn("tbl") and child_index not in protected and paragraph_index >= prefix:
            isolated = ET.Element(qn("document"))
            ET.SubElement(isolated, qn("body")).append(child)
            normalize_tables(isolated, 0, rules)
    # Preserve complex section layouts; normalize page geometry only for one-section sources.
    if len(audit.get("sections") or []) <= 1 and prefix == 0:
        update_sections(root, rules)
    page = rules.get("page") or {}
    mode = page.get("header_text_mode", "preserve")
    log.extend(apply_header_policy(root, files, str(plan.get("headerText") or ""), mode, __import__("apply_ooxml_fixes").replace_header_text, create_missing=page.get("create_missing_body_headers", True)))
    log.extend(initialize_fresh_footers(root, files, str(rules.get("latin_font_override") or "Times New Roman"), missing_only=True))
    files["word/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
    files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as output:
        for name, data in files.items():
            output.writestr(name, data)
    return {"executionLog": log, "preservedRegions": risky, "protectedBodyChildren": sorted(protected), "allowStructuralRebuild": False}

def run(args, workspace, manifest, profile, policy, api):
    """Normal route. Audit-only and explicitly authorized rebuild stay on legacy APIs."""
    from extract_source_evidence import extract_docx_evidence, extract_text_evidence
    from thesis_ir import build_thesis_ir, apply_semantic_overrides
    from validate_thesis_ir import validate
    from formatter_core.word_semantics import apply_ir_corrections
    from build_docx_from_markdown import build_source_docx_from_markdown
    work, source = workspace.work, workspace.source
    commands = []
    rules = load_rules(str(profile.rules_yaml))
    audit = None
    scan = None
    if source.suffix.lower() == ".docx":
        scan = inspect_risks(source)
        manifest["processingRoute"] = "conservative-repair" if args.conservative or scan[2] else "smart-normalize"
        audit = parse_document(str(source), template_docx=str(profile.template_docx) if profile.template_docx else None, rules_path=str(profile.rules_yaml), detail="semantic", _package_files=scan[0], _document_root=scan[1])
        evidence = extract_docx_evidence(source, None, str(profile.rules_yaml), audit=audit)
    else:
        manifest["processingRoute"] = "fast-build"
        evidence = extract_text_evidence(source)
    ir = build_thesis_ir(evidence)
    if audit:
        from formatter_core.sections import inspect_requirements
        requirements = inspect_requirements(audit, rules, ir)
        write(work / "document_requirements.json", requirements)
        manifest["documentRequirements"] = requirements
        if requirements["missing"]:
            print("排版前发现缺项：" + "、".join(item["label"] for item in requirements["missing"]))
    manifest["semanticExtractionCount"] = 1
    if args.semantic_overrides:
        ir = apply_semantic_overrides(ir, json.loads(Path(args.semantic_overrides).read_text(encoding="utf-8")))
    ir_report = validate(ir)
    write(work / "thesis_ir.json", ir)
    write(work / "thesis_ir_validation.json", ir_report)
    if not ir_report.get("passed"):
        raise ValueError(f"ThesisIR validation failed: {ir_report}")
    manifest["effectiveMode"] = "conservative-repair" if audit else "rebuild"
    manifest["effectiveExpertMode"] = manifest["effectiveMode"]
    manifest["structuralRebuildAllowed"] = args.allow_structural_rebuild
    manifest["strategySelection"] = {"chosenStrategy": "preserve_first" if audit else "text_rebuild"}
    manifest["preflightDecision"] = {"documentRiskClass": "C" if scan and scan[2] else "A"}
    if args.backend != "latex":
        candidate = work / "repaired.docx"
        if audit:
            plan = build_repair_plan(audit, style_map_path=str(profile.style_map_yaml) if profile.style_map_yaml else None, front_matter_policy_path=str(profile.front_matter_policy_yaml) if profile.front_matter_policy_yaml else None)
            apply_ir_corrections(plan, audit, ir)
            plan["contentPolicy"] = "exact"
            write(work / "repair_plan.json", plan)
            execution = normalize_docx(source, candidate, audit, plan, rules, scan=scan)
            write(work / "repair_execution.json", execution)
            manifest["preservedRegions"] = execution["preservedRegions"]
            manifest["semanticCorrections"] = plan.get("semanticCorrectionReport")
        else:
            built = work / "markdown_source.docx"
            build_source_docx_from_markdown(str(source), str(built), rules, thesis_ir=ir)
            # Deterministic template composition/style pass, without inspection,
            # preflight or a repair-planner CLI on the freshly built document.
            built_audit = parse_document(str(built), template_docx=str(profile.template_docx) if profile.template_docx else None, rules_path=str(profile.rules_yaml), detail="semantic")
            plan = build_repair_plan(built_audit, style_map_path=str(profile.style_map_yaml) if profile.style_map_yaml else None, front_matter_policy_path=str(profile.front_matter_policy_yaml) if profile.front_matter_policy_yaml else None, allow_structural_rebuild=True)
            write(work / "repair_plan.json", plan)
            command = api.add_common_flags([api.PYTHON, "scripts/apply_ooxml_fixes.py", str(built), str(candidate), "--plan-json", str(work / "repair_plan.json"), "--report-json", str(work / "repair_execution.json"), "--allow-structural-rebuild", "--defer-validation"], profile)
            run_in_process(command, api.ROOT)
            commands.append(command)
        command = api.add_common_flags([api.PYTHON, "scripts/validate_docx.py", str(candidate), "--output", str(work / "validation_report.json"), "--render-output-dir", str(work / "render_validation"), "--plan-json", str(work / "repair_plan.json"), "--thesis-ir-json", str(work / "thesis_ir.json"), "--render-backend", args.render_backend, *policy.flags()], profile)
        if audit:
            command.extend(["--before-docx", str(source)])
        run_in_process(command, api.ROOT)
        commands.append(command)
        from formatter_core.visual import complete_visual_workflow
        complete_visual_workflow(args, workspace, manifest, command, api)
        report = api.read_json(work / "validation_report.json")
        from formatter_core.sections import inspect_requirements
        final_audit = parse_document(str(candidate), rules_path=str(profile.rules_yaml), detail="semantic")
        remaining = inspect_requirements(final_audit, rules)
        manifest["documentRequirementsAfter"] = remaining
        write(work / "document_requirements_after.json", remaining)
        if remaining["missing"]:
            print("当前文档仍缺少：" + "、".join(item["label"] for item in remaining["missing"]))
        if not audit:
            from formatter_core.integrity import verify_generated, verify_source_assets
            integrity = verify_generated(work / "markdown_source.docx", candidate)
            source_assets = verify_source_assets(ir, work / "markdown_source.docx")
            integrity["issues"].extend(source_assets["issues"])
            integrity["passed"] = not integrity["issues"]
            report["contentIntegrity"] = integrity
            if not integrity["passed"]:
                report["issues"].extend(integrity["issues"])
                from validate_docx import summarize_quality_gate
                report["qualityGate"] = summarize_quality_gate(report["issues"])
            write(work / "validation_report.json", report)
        manifest["qualityGate"] = report["qualityGate"]
        manifest["contentIntegrity"] = report.get("contentIntegrity")
        manifest["status"] = "completed" if report["qualityGate"]["passed"] else "completed_with_hard_failures"
        manifest["assetPreservation"] = api.build_asset_preservation_report(source, work)
    else:
        manifest["status"] = "completed"
    latex_result = None
    if args.backend in {"latex", "both"}:
        latex_result = api.run_latex(ir, source, rules, profile.profile_dir, workspace, policy, no_compile=args.no_compile, timeout=args.compile_timeout, engine_name=args.latex_engine, metadata_path=Path(args.latex_metadata) if args.latex_metadata else None, bibliography_path=Path(args.latex_bibliography) if args.latex_bibliography else None)
    api.finish_run(args, workspace, manifest, commands, latex_result)
