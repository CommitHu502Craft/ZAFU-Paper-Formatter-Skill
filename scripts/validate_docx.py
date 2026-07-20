#!/usr/bin/env python3
import argparse
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from docx_ooxml import diff_audits, load_rules, parse_document, write_json
from render_validate_docx import run_render_validation
from validate_visual_contracts import validate_visual_contracts

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CORE_INTERNAL_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments": "word/comments.xml",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes": "word/footnotes.xml",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes": "word/endnotes.xml",
}

REQUIRED_DOCX_PARTS = [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/styles.xml",
]

DEFAULT_REQUIRED_STYLE_IDS = {"zafu_body", "zafu_heading1", "zafu_heading2", "zafu_heading3"}
HARD_FAIL_ISSUE_TYPES = {
    "docx_package_issues",
    "missing_header_footer_parts",
    "missing_styles_part",
    "core_relationship_part_issues",
    "image_relationship_issues",
    "missing_required_styles",
    "missing_front_matter_markers",
    "abstract_label_content_format",
    "keywords_label_content_format",
    "excessive_section_count",
    "excessive_page_breaks",
}


def load_profile_yaml(path):
    if not path:
        return {}
    return load_rules(path)


def collect_enabled_validators(validators_config, layer):
    validators = (validators_config.get("validators") or {}) if isinstance(validators_config, dict) else {}
    items = validators.get(layer) or []
    return {str(item) for item in items if isinstance(item, str)}


def rule_enabled(enabled, validator_name):
    return not enabled or validator_name in enabled


def resolve_required_style_ids(style_map_config):
    style_roles = (style_map_config.get("style_roles") or {}) if isinstance(style_map_config, dict) else {}
    mapping = {
        "body_text": "zafu_body",
        "heading1": "zafu_heading1",
        "heading2": "zafu_heading2",
        "heading3": "zafu_heading3",
    }
    required = set()
    for role_key, fallback in mapping.items():
        value = style_roles.get(role_key)
        if isinstance(value, str) and value.strip() and value != "template_derived":
            required.add(value.strip())
        else:
            required.add(fallback)
    return required or set(DEFAULT_REQUIRED_STYLE_IDS)


def summarize_quality_gate(issues):
    hard_failures = []
    warnings = []
    for issue in issues:
        issue_type = str((issue or {}).get("type") or "")
        target = hard_failures if issue_type in HARD_FAIL_ISSUE_TYPES else warnings
        target.append(issue)
    return {
        "passed": not hard_failures,
        "hardFailureCount": len(hard_failures),
        "warningCount": len(warnings),
        "hardFailures": hard_failures,
        "warnings": warnings,
    }


def normalize_marker_text(value):
    return "".join(str(value or "").split())


def check_required_front_matter_markers(audit, front_matter_policy):
    if not isinstance(front_matter_policy, dict):
        return []
    policy = front_matter_policy.get("policy") or {}
    markers = policy.get("required_front_matter_markers") or []
    normalized_markers = [normalize_marker_text(item) for item in markers if isinstance(item, str) and item.strip()]
    if not normalized_markers:
        return []
    normalized_paragraphs = [normalize_marker_text((paragraph or {}).get("text")) for paragraph in (audit.get("paragraphs") or [])[:60]]
    joined = "\n".join(normalized_paragraphs)
    missing = [marker for marker in normalized_markers if marker not in joined]
    return missing


def section_geometry_check(audit, rules, template_audit=None):
    page = rules.get("page") or {}
    issues = []
    template_sections = (template_audit or {}).get("sections") or []
    for section_index, section in enumerate(audit.get("sections") or []):
        section_page = section.get("page") or {}
        template_page = template_sections[section_index].get("page") if section_index < len(template_sections) else None
        for field, rule_key in [("top", "margin_top_cm"), ("bottom", "margin_bottom_cm"), ("left", "margin_left_cm"), ("right", "margin_right_cm"), ("header", "header_cm"), ("footer", "footer_cm")]:
            if section_index == 0 and template_page:
                expected = ((template_page or {}).get(field) or {}).get("cm")
            else:
                expected = page.get(rule_key)
                if expected is None and template_page:
                    expected = ((template_page or {}).get(field) or {}).get("cm")
            actual = ((section_page.get(field) or {}).get("cm")) if section_page.get(field) else None
            if expected is not None and actual is not None and abs(float(expected) - float(actual)) > 0.02:
                issues.append(
                    {
                        "section": section_index,
                        "field": field,
                        "expectedCm": expected,
                        "actualCm": actual,
                    }
                )
    return issues


def header_bottom_border_check(docx_path):
    issues = []
    with zipfile.ZipFile(docx_path) as zf:
        rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        rels = {
            rel.attrib["Id"]: rel.attrib.get("Target")
            for rel in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
            if rel.attrib.get("Id")
        }
        document_root = ET.fromstring(zf.read("word/document.xml"))
        body = document_root.find("w:body", NS)
        required_headers = set()
        if body is not None:
            for child in list(body):
                sect = None
                if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p":
                    ppr = child.find("w:pPr", NS)
                    sect = ppr.find("w:sectPr", NS) if ppr is not None else None
                elif child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr":
                    sect = child
                if sect is None:
                    continue
                for header_ref in sect.findall("w:headerReference", NS):
                    ref_type = header_ref.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
                    if ref_type != "default":
                        continue
                    rel_id = header_ref.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = rels.get(rel_id or "")
                    if target:
                        required_headers.add(f"word/{target}" if not target.startswith("word/") else target)
        for name in zf.namelist():
            if not (name.startswith("word/header") and name.endswith(".xml")):
                continue
            if required_headers and name not in required_headers:
                continue
            root = ET.fromstring(zf.read(name))
            paragraph = root.find("w:p", NS)
            if paragraph is None:
                continue
            text = "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()
            if not text:
                continue
            ppr = paragraph.find("w:pPr", NS)
            pbdr = ppr.find("w:pBdr", NS) if ppr is not None else None
            bottom = pbdr.find("w:bottom", NS) if pbdr is not None else None
            actual = bottom.attrib if bottom is not None else None
            if bottom is None or bottom.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") not in {"single", "thick", "double"}:
                issues.append({"part": name, "expected": "bottom border line", "actual": actual})
    return issues


def docx_package_check(docx_path):
    issues = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = set(zf.namelist())
    except zipfile.BadZipFile:
        return [{"type": "bad_zip_package", "path": str(Path(docx_path).resolve())}]

    for part in REQUIRED_DOCX_PARTS:
        if part not in names:
            issues.append({"type": "missing_required_part", "part": part})
    return issues


def numbering_part_check(audit):
    numbering = audit.get("numbering") or {}
    issues = []
    heading_tree = (audit.get("numberingAnalysis") or {}).get("headingTree") or []
    uses_word_numbering = any(item.get("numId") is not None for item in heading_tree)
    if heading_tree and uses_word_numbering and not (numbering.get("abstractNums") or numbering.get("nums")):
        issues.append({"type": "missing_numbering_definitions_for_detected_headings"})
    return issues


def image_relationship_integrity_check(docx_path):
    issues = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = set(zf.namelist())
            if "word/document.xml" not in names:
                return [{"type": "missing_required_part", "part": "word/document.xml"}]

            def rels_path_for_part(part_name):
                part_path = Path(part_name)
                return str(part_path.parent / "_rels" / f"{part_path.name}.rels").replace("\\", "/")

            def normalize_target(part_name, target):
                normalized = str(Path(part_name).parent / Path(target or "")).replace("\\", "/")
                normalized_parts = []
                for token in normalized.split("/"):
                    if token in {"", "."}:
                        continue
                    if token == "..":
                        if normalized_parts:
                            normalized_parts.pop()
                    else:
                        normalized_parts.append(token)
                return "/".join(normalized_parts)

            xml_parts = sorted(
                name for name in names
                if name.startswith("word/")
                and name.endswith(".xml")
                and "/_rels/" not in name
            )
            for part_name in xml_parts:
                root = ET.fromstring(zf.read(part_name))
                rels = {}
                rels_path = rels_path_for_part(part_name)
                if rels_path in names:
                    rels_root = ET.fromstring(zf.read(rels_path))
                    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
                        rel_id = rel.attrib.get("Id")
                        if rel_id:
                            rels[rel_id] = {
                                "type": rel.attrib.get("Type"),
                                "target": rel.attrib.get("Target"),
                            }

                for blip in root.findall(".//a:blip", NS):
                    rel_id = blip.attrib.get(f"{{{NS['r']}}}embed")
                    if not rel_id:
                        continue
                    rel = rels.get(rel_id)
                    if rel is None:
                        issues.append(
                            {
                                "type": "missing_image_relationship",
                                "part": part_name,
                                "relationshipId": rel_id,
                            }
                        )
                        continue
                    if rel.get("type") != IMAGE_REL_TYPE:
                        issues.append(
                            {
                                "type": "non_image_relationship_used_by_image_embed",
                                "part": part_name,
                                "relationshipId": rel_id,
                                "relationshipType": rel.get("type"),
                                "target": rel.get("target"),
                            }
                        )
                        continue
                    part_target = normalize_target(part_name, rel.get("target"))
                    if part_target not in names:
                        issues.append(
                            {
                                "type": "missing_image_part",
                                "part": part_name,
                                "relationshipId": rel_id,
                                "target": rel.get("target"),
                                "resolvedPart": part_target,
                            }
                        )
    except zipfile.BadZipFile:
        return [{"type": "bad_zip_package", "path": str(Path(docx_path).resolve())}]
    return issues


def core_relationship_part_integrity_check(docx_path):
    issues = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            names = set(zf.namelist())
            rels_name = "word/_rels/document.xml.rels"
            if rels_name not in names:
                return [{"type": "missing_required_part", "part": rels_name}]
            rels_root = ET.fromstring(zf.read(rels_name))
            for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
                rel_type = rel.attrib.get("Type", "")
                expected_part = CORE_INTERNAL_REL_TYPES.get(rel_type)
                if not expected_part:
                    continue
                if rel.attrib.get("TargetMode") == "External":
                    issues.append(
                        {
                            "type": "unexpected_external_core_relationship",
                            "relationshipId": rel.attrib.get("Id"),
                            "relationshipType": rel_type,
                            "target": rel.attrib.get("Target"),
                        }
                    )
                    continue
                if expected_part not in names:
                    issues.append(
                        {
                            "type": "missing_core_relationship_part",
                            "relationshipId": rel.attrib.get("Id"),
                            "relationshipType": rel_type,
                            "expectedPart": expected_part,
                            "target": rel.attrib.get("Target"),
                        }
                    )
    except zipfile.BadZipFile:
        return [{"type": "bad_zip_package", "path": str(Path(docx_path).resolve())}]
    return issues


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate a repaired DOCX and optionally compare it against the original.")
    parser.add_argument("docx", help="DOCX file to validate")
    parser.add_argument("--before-docx", help="Optional original DOCX for diff reporting")
    parser.add_argument("--template-docx", help="Optional template DOCX for baseline-aware audit")
    parser.add_argument("--rules-yaml", required=True, help="Rules YAML")
    parser.add_argument("--style-map-yaml", help="Optional profile style-role mapping YAML")
    parser.add_argument("--front-matter-policy-yaml", help="Optional profile front-matter policy YAML")
    parser.add_argument("--validators-yaml", help="Optional profile validator selection YAML")
    parser.add_argument("--plan-json", help="Optional plan JSON to include in the report")
    parser.add_argument("--output", "-o", help="Write JSON output to this path")
    parser.add_argument("--render-output-dir", help="Optional directory for render validation artifacts")
    args = parser.parse_args()

    rules = load_rules(args.rules_yaml)
    style_map_config = load_profile_yaml(args.style_map_yaml)
    front_matter_policy = load_profile_yaml(args.front_matter_policy_yaml)
    validators_config = load_profile_yaml(args.validators_yaml)
    structural_enabled = collect_enabled_validators(validators_config, "structural")
    rule_enabled_set = collect_enabled_validators(validators_config, "rule")
    render_enabled = collect_enabled_validators(validators_config, "render")
    audit = parse_document(args.docx, template_docx=args.template_docx, rules_path=args.rules_yaml)
    plan = json.loads(Path(args.plan_json).read_text(encoding="utf-8")) if args.plan_json else None
    template_audit = parse_document(args.template_docx, rules_path=args.rules_yaml) if args.template_docx else None

    geometry_issues = section_geometry_check(audit, rules, template_audit=template_audit)
    header_border_issues = header_bottom_border_check(args.docx)
    header_text = (plan or {}).get("headerText") or None
    if not header_text:
        for paragraph in audit.get("paragraphs") or []:
            if (paragraph.get("role") or {}).get("role") == "title_cn":
                candidate = (paragraph.get("text") or "").strip()
                if candidate:
                    header_text = candidate
                    break
    if not header_text:
        header_text = ((audit.get("preservationHints") or {}).get("preferredHeaderText")) or None
    header_ok = True
    if header_text:
        header_texts = [part.get("text") for name, part in (audit.get("headerFooterParts") or {}).items() if "header" in name]
        header_ok = any(text == header_text for text in header_texts) if header_texts else False

    style_ids = set(audit.get("styles") or {})
    required_style_ids = resolve_required_style_ids(style_map_config)
    missing_styles = sorted(required_style_ids - style_ids)

    numbering_analysis = audit.get("numberingAnalysis") or {}
    preserve_regions = (plan or {}).get("preserveRegions") or []

    def in_preserved_region(paragraph_index: int) -> bool:
        for region in preserve_regions:
            start = region.get("startParagraph", 0)
            end = region.get("endParagraphExclusive")
            if end is not None and start <= paragraph_index < end:
                return True
        return False

    filtered_font_risks = [
        item for item in ((audit.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or [])
        if not in_preserved_region(item.get("paragraphIndex", -1))
    ]
    package_issues = docx_package_check(args.docx)
    numbering_part_issues = numbering_part_check(audit)
    image_relationship_issues = image_relationship_integrity_check(args.docx)
    core_relationship_part_issues = core_relationship_part_integrity_check(args.docx)

    structural_validation = {
        "checks": {
            "docxPackageIntegrity": not package_issues if rule_enabled(structural_enabled, "package_integrity") else None,
            "requiredPartsPresent": (
                not any(item.get("type") == "missing_required_part" for item in package_issues)
                if rule_enabled(structural_enabled, "package_integrity")
                else None
            ),
            "headerFooterPartsPresent": (
                bool(audit.get("headerFooterParts") or {})
                if rule_enabled(structural_enabled, "header_footer_part_presence")
                else None
            ),
            "stylesPresent": bool(audit.get("styles") or {}),
            "numberingPartValid": not numbering_part_issues if rule_enabled(structural_enabled, "relationship_integrity") else None,
            "coreRelationshipPartsValid": not core_relationship_part_issues if rule_enabled(structural_enabled, "relationship_integrity") else None,
            "imageRelationshipsValid": not image_relationship_issues if rule_enabled(structural_enabled, "relationship_integrity") else None,
        },
        "issues": [],
    }
    if package_issues and rule_enabled(structural_enabled, "package_integrity"):
        structural_validation["issues"].append({"type": "docx_package_issues", "details": package_issues})
    if not (audit.get("headerFooterParts") or {}) and rule_enabled(structural_enabled, "header_footer_part_presence"):
        structural_validation["issues"].append({"type": "missing_header_footer_parts"})
    if not (audit.get("styles") or {}):
        structural_validation["issues"].append({"type": "missing_styles_part"})
    if numbering_part_issues and rule_enabled(structural_enabled, "relationship_integrity"):
        structural_validation["issues"].append({"type": "numbering_part_issues", "details": numbering_part_issues})
    if core_relationship_part_issues and rule_enabled(structural_enabled, "relationship_integrity"):
        structural_validation["issues"].append({"type": "core_relationship_part_issues", "details": core_relationship_part_issues[:100]})
    if image_relationship_issues and rule_enabled(structural_enabled, "relationship_integrity"):
        structural_validation["issues"].append({"type": "image_relationship_issues", "details": image_relationship_issues[:100]})

    missing_front_matter_markers = check_required_front_matter_markers(audit, front_matter_policy)
    rule_validation = {
        "checks": {
            "sectionGeometry": not geometry_issues if rule_enabled(rule_enabled_set, "page_geometry") else None,
            "headerText": header_ok if rule_enabled(rule_enabled_set, "header_text") else None,
            "headerBottomBorder": not header_border_issues if rule_enabled(rule_enabled_set, "header_bottom_border") else None,
            "requiredStyles": not missing_styles if rule_enabled(rule_enabled_set, "body_style") else None,
            "headingSequence": not bool(numbering_analysis.get("anomalies")) if rule_enabled(rule_enabled_set, "heading_hierarchy") else None,
            "fontSlotRisks": len(filtered_font_risks) if rule_enabled(rule_enabled_set, "body_style") else None,
            "frontMatterStructure": not bool((audit.get("frontMatterAnalysis") or {}).get("issues")) if rule_enabled(rule_enabled_set, "front_matter_structure") else None,
            "requiredFrontMatterMarkers": not bool(missing_front_matter_markers) if rule_enabled(rule_enabled_set, "front_matter_structure") else None,
            "captionLayout": not bool((audit.get("captionLayout") or {}).get("issues")) if rule_enabled(rule_enabled_set, "caption_layout") else None,
            "crossReferences": not bool((audit.get("crossReferenceAnalysis") or {}).get("unresolved")) if rule_enabled(rule_enabled_set, "cross_references") else None,
        },
        "issues": [],
    }
    report_output_path = Path(args.output).resolve() if args.output else None
    default_render_output_dir = (
        report_output_path.parent / "render_validation"
        if report_output_path is not None
        else (Path(args.docx).resolve().parent / "render_validation")
    )
    render_output_dir = Path(args.render_output_dir).resolve() if args.render_output_dir else default_render_output_dir
    if render_enabled and "pdf_conversion_if_available" not in render_enabled:
        render_validation = {
            "sourceDocx": str(Path(args.docx).resolve()),
            "outputDir": str(render_output_dir),
            "available": False,
            "backend": None,
            "backendPath": None,
            "status": "skipped",
            "reason": "render validator disabled by profile validators.yaml",
            "pdfPath": None,
            "pageCount": None,
            "previewDir": str(render_output_dir / "before_after_page_preview"),
            "suggestedReviewPages": {
                "cover": None,
                "toc": None,
                "abstract": None,
                "firstBodyPage": None,
                "references": None,
            },
        }
    else:
        render_validation = run_render_validation(args.docx, str(render_output_dir))
    (render_output_dir / "render_validation_report.json").write_text(json.dumps(render_validation, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        visual_validation = validate_visual_contracts(args.docx)
    except (KeyError, ValueError, OSError, zipfile.BadZipFile) as exc:
        # Polluted packages can be sufficiently readable for the OOXML audit but
        # still fail python-docx relationship traversal (for example a target
        # literally named "NULL"). Audit mode must report that limitation rather
        # than crash and leave the dispatcher in a false running state.
        visual_validation = {
            "checks": {},
            "issues": [
                {
                    "type": "visual_contract_validation_unavailable",
                    "details": [{"error": type(exc).__name__, "message": str(exc)}],
                }
            ],
        }

    report = {
        "generatedAt": audit["generatedAt"],
        "source": audit["source"],
        "checks": {},
        "issues": [],
        "auditSummary": audit.get("summary"),
        "numberingAnalysis": numbering_analysis,
        "manualReview": (plan or {}).get("manualReview") or [],
        "profileValidationConfig": {
            "styleMapYaml": str(Path(args.style_map_yaml).resolve()) if args.style_map_yaml else None,
            "frontMatterPolicyYaml": str(Path(args.front_matter_policy_yaml).resolve()) if args.front_matter_policy_yaml else None,
            "validatorsYaml": str(Path(args.validators_yaml).resolve()) if args.validators_yaml else None,
            "requiredStyleIds": sorted(required_style_ids),
            "enabledStructuralValidators": sorted(structural_enabled) if structural_enabled else [],
            "enabledRuleValidators": sorted(rule_enabled_set) if rule_enabled_set else [],
            "enabledRenderValidators": sorted(render_enabled) if render_enabled else [],
            "frontMatterPolicy": (front_matter_policy.get("policy") or {}) if isinstance(front_matter_policy, dict) else {},
        },
        "structuralValidation": structural_validation,
        "ruleValidation": rule_validation,
        "renderValidation": render_validation,
        "visualValidation": visual_validation,
    }
    if geometry_issues and rule_enabled(rule_enabled_set, "page_geometry"):
        rule_validation["issues"].append({"type": "section_geometry_mismatch", "details": geometry_issues})
    if not header_ok and rule_enabled(rule_enabled_set, "header_text"):
        rule_validation["issues"].append({"type": "header_text_mismatch", "expected": header_text})
    if header_border_issues and rule_enabled(rule_enabled_set, "header_bottom_border"):
        rule_validation["issues"].append({"type": "header_bottom_border_missing", "details": header_border_issues})
    if missing_styles and rule_enabled(rule_enabled_set, "body_style"):
        rule_validation["issues"].append({"type": "missing_required_styles", "styleIds": missing_styles})
    if numbering_analysis.get("anomalies") and rule_enabled(rule_enabled_set, "heading_hierarchy"):
        rule_validation["issues"].append({"type": "numbering_anomalies", "details": numbering_analysis["anomalies"][:100]})
    if filtered_font_risks and rule_enabled(rule_enabled_set, "body_style"):
        rule_validation["issues"].append({"type": "font_slot_risks", "details": filtered_font_risks[:100]})
    if missing_front_matter_markers and rule_enabled(rule_enabled_set, "front_matter_structure"):
        rule_validation["issues"].append({"type": "missing_front_matter_markers", "details": missing_front_matter_markers})
    if (audit.get("frontMatterAnalysis") or {}).get("issues") and rule_enabled(rule_enabled_set, "front_matter_structure"):
        rule_validation["issues"].append({"type": "front_matter_structure", "details": (audit.get("frontMatterAnalysis") or {}).get("issues")[:50]})
    if (audit.get("captionLayout") or {}).get("issues") and rule_enabled(rule_enabled_set, "caption_layout"):
        rule_validation["issues"].append({"type": "caption_layout", "details": (audit.get("captionLayout") or {}).get("issues")[:50]})
    if (audit.get("crossReferenceAnalysis") or {}).get("unresolved") and rule_enabled(rule_enabled_set, "cross_references"):
        rule_validation["issues"].append({"type": "unresolved_cross_references", "details": (audit.get("crossReferenceAnalysis") or {}).get("unresolved")[:50]})

    if args.before_docx:
        before = parse_document(args.before_docx, template_docx=args.template_docx, rules_path=args.rules_yaml)
        report["diffReport"] = diff_audits(before, audit, plan=plan, execution_log=None)

    report["checks"] = dict(structural_validation["checks"])
    report["checks"].update(rule_validation["checks"])
    report["checks"].update(visual_validation.get("checks") or {})
    report["issues"].extend(structural_validation["issues"])
    report["issues"].extend(rule_validation["issues"])
    report["issues"].extend(visual_validation.get("issues") or [])
    report["qualityGate"] = summarize_quality_gate(report["issues"])
    report["validationSummary"] = {
        "structuralPassed": not structural_validation["issues"],
        "rulePassed": not rule_validation["issues"],
        "visualPassed": not bool(visual_validation.get("issues")),
        "renderPassed": render_validation.get("status") in {"ok", "skipped"},
        "renderStatus": render_validation.get("status"),
        "qualityGatePassed": report["qualityGate"]["passed"],
        "hardFailureCount": report["qualityGate"]["hardFailureCount"],
        "warningCount": report["qualityGate"]["warningCount"],
    }
    report["passed"] = not report["issues"] and render_validation.get("status") in {"ok", "unavailable", "skipped"}
    write_json(report, args.output)


if __name__ == "__main__":
    main()
