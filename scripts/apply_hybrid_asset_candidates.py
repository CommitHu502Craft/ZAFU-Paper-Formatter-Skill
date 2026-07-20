#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import posixpath
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from typing import Any, Dict, List, Optional, Tuple


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {
    "w": W_NS,
    "r": R_NS,
    "a": A_NS,
    "pr": PKG_REL_NS,
    "ct": CONTENT_TYPES_NS,
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def clone_xml(node: ET.Element) -> ET.Element:
    return ET.fromstring(ET.tostring(node, encoding="utf-8"))


def paragraph_text(node: ET.Element) -> str:
    return "".join(text.text or "" for text in node.findall(".//w:t", NS)).strip()


def read_zip_tree(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def body_paragraphs(document_root: ET.Element) -> List[ET.Element]:
    body = document_root.find("w:body", NS)
    if body is None:
        return []
    return [child for child in list(body) if child.tag == qn(W_NS, "p")]


def body_tables(document_root: ET.Element) -> List[ET.Element]:
    body = document_root.find("w:body", NS)
    if body is None:
        return []
    return [child for child in list(body) if child.tag == qn(W_NS, "tbl")]


def find_body(document_root: ET.Element) -> ET.Element:
    body = document_root.find("w:body", NS)
    if body is None:
        raise SystemExit("word/document.xml has no body element")
    return body


def parse_relationships(root: ET.Element) -> List[ET.Element]:
    return list(root.findall("pr:Relationship", NS))


def next_rid(relationships: List[ET.Element]) -> str:
    numeric_ids = []
    for rel in relationships:
        rid = rel.attrib.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            numeric_ids.append(int(rid[3:]))
    return f"rId{(max(numeric_ids) if numeric_ids else 0) + 1}"


def unique_media_target(existing_targets: set[str], original_target: str, index: int) -> str:
    suffix = PurePosixPath(original_target).suffix or ".bin"
    stem = f"hybrid_asset_{index:03d}"
    candidate = posixpath.join("media", f"{stem}{suffix.lower()}")
    serial = 1
    while candidate in existing_targets:
        serial += 1
        candidate = posixpath.join("media", f"{stem}_{serial:02d}{suffix.lower()}")
    return candidate


def content_type_for_suffix(suffix: str) -> Optional[str]:
    mapping = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".bmp": "image/bmp",
        ".tif": "image/tiff",
        ".tiff": "image/tiff",
        ".webp": "image/webp",
        ".emf": "image/x-emf",
        ".wmf": "image/x-wmf",
    }
    return mapping.get(suffix.lower())


def ensure_content_type(content_types_root: ET.Element, target_name: str) -> None:
    suffix = PurePosixPath(target_name).suffix
    content_type = content_type_for_suffix(suffix)
    if not content_type:
        return
    extension = suffix.lstrip(".")
    for node in content_types_root.findall("ct:Default", NS):
        if node.attrib.get("Extension", "").lower() == extension.lower():
            return
    default = ET.Element(qn(CONTENT_TYPES_NS, "Default"))
    default.set("Extension", extension)
    default.set("ContentType", content_type)
    content_types_root.append(default)


def resolve_word_target(rel_target: str) -> str:
    normalized = rel_target.replace("\\", "/")
    if normalized.startswith("/"):
        normalized = normalized.lstrip("/")
    if normalized.startswith("word/"):
        return normalized
    return posixpath.normpath(posixpath.join("word", normalized))


def build_attachment_report(
    source_docx: Path,
    intermediate_docx: Path,
    output_docx: Path,
    attached: List[Dict[str, Any]],
    manual_review: List[Dict[str, Any]],
    skipped_by_policy: List[Dict[str, Any]],
    skipped: List[Dict[str, Any]],
) -> Dict[str, Any]:
    return {
        "sourceDocx": str(source_docx),
        "intermediateDocx": str(intermediate_docx),
        "outputDocx": str(output_docx),
        "attachedAssets": attached,
        "manualReviewAssets": manual_review,
        "skippedByPolicy": skipped_by_policy,
        "skippedAssets": skipped,
        "attachedCount": len(attached),
        "manualReviewCount": len(manual_review),
        "skippedByPolicyCount": len(skipped_by_policy),
        "skippedCount": len(skipped),
        "notes": [
            "Current phase reattaches only high-confidence image and table candidates.",
            "Low-confidence candidates remain deferred to manual review.",
            "Candidates still require a matching caption block in the rebuilt intermediate DOCX.",
        ],
    }


def apply_candidates(
    source_docx: Path,
    intermediate_docx: Path,
    thesis_ir: Dict[str, Any],
    output_docx: Path,
) -> Dict[str, Any]:
    candidates = thesis_ir.get("attachableAssetCandidates") or []
    manual_review = [
        item
        for item in candidates
        if item.get("assetKind") in {"image", "table"}
        and item.get("recommendedAction") == "manual_review"
    ]
    skipped_by_policy = [
        item
        for item in candidates
        if item.get("assetKind") in {"image", "table"}
        and item.get("recommendedAction") == "skip"
    ]
    actionable = [
        item
        for item in candidates
        if item.get("assetKind") in {"image", "table"}
        and item.get("recommendedAction") == "reattach_candidate"
        and float(item.get("confidence") or 0.0) >= 0.7
    ]

    with zipfile.ZipFile(source_docx) as src_zip, zipfile.ZipFile(intermediate_docx) as base_zip:
        source_doc_root = read_zip_tree(src_zip, "word/document.xml")
        source_rels_root = read_zip_tree(src_zip, "word/_rels/document.xml.rels")
        target_doc_root = read_zip_tree(base_zip, "word/document.xml")
        target_rels_root = read_zip_tree(base_zip, "word/_rels/document.xml.rels")
        content_types_root = read_zip_tree(base_zip, "[Content_Types].xml")

        source_paragraphs = body_paragraphs(source_doc_root)
        source_tables = body_tables(source_doc_root)
        target_body = find_body(target_doc_root)
        target_paragraphs = body_paragraphs(target_doc_root)
        target_relationships = parse_relationships(target_rels_root)
        source_relationships = {rel.attrib.get("Id"): rel for rel in parse_relationships(source_rels_root)}
        existing_targets = {rel.attrib.get("Target", "") for rel in target_relationships}

        file_map = {name: base_zip.read(name) for name in base_zip.namelist()}
        attached: List[Dict[str, Any]] = []
        skipped: List[Dict[str, Any]] = []
        inserted_caption_texts: set[str] = set()

        for asset_index, candidate in enumerate(actionable, start=1):
            source_index = candidate.get("sourceIndex")
            caption_text = str(candidate.get("captionText") or "").strip()
            if not caption_text:
                skipped.append({**candidate, "reason": "missing_caption_text"})
                continue
            if caption_text in inserted_caption_texts:
                skipped.append({**candidate, "reason": "caption_already_handled"})
                continue

            target_caption_paragraph = None
            for paragraph in target_paragraphs:
                if paragraph_text(paragraph) == caption_text:
                    target_caption_paragraph = paragraph
                    break
            if target_caption_paragraph is None:
                skipped.append({**candidate, "reason": "caption_not_found_in_intermediate_docx"})
                continue

            copied_targets: List[str] = []
            children = list(target_body)
            insertion_index = children.index(target_caption_paragraph)

            if candidate.get("assetKind") == "image":
                if source_index is None or int(source_index) >= len(source_paragraphs):
                    skipped.append({**candidate, "reason": "source_paragraph_out_of_range"})
                    continue
                source_paragraph = source_paragraphs[int(source_index)]
                embed_nodes = []
                for node in source_paragraph.iter():
                    embed = node.attrib.get(qn(R_NS, "embed")) or node.attrib.get(qn(R_NS, "id"))
                    if embed and embed in source_relationships:
                        embed_nodes.append((node, embed))
                if not embed_nodes:
                    skipped.append({**candidate, "reason": "no_embedded_relationships_found"})
                    continue

                cloned_node = clone_xml(source_paragraph)
                cloned_embed_nodes = []
                for node in cloned_node.iter():
                    embed = node.attrib.get(qn(R_NS, "embed")) or node.attrib.get(qn(R_NS, "id"))
                    if embed and embed in source_relationships:
                        cloned_embed_nodes.append((node, embed))

                for node, old_rid in cloned_embed_nodes:
                    source_rel = source_relationships.get(old_rid)
                    if source_rel is None:
                        continue
                    rel_target = source_rel.attrib.get("Target")
                    if not rel_target:
                        continue
                    source_part_name = resolve_word_target(rel_target)
                    if source_part_name not in src_zip.namelist():
                        continue
                    new_target = unique_media_target(existing_targets, rel_target, asset_index)
                    existing_targets.add(new_target)
                    new_rid = next_rid(target_relationships)
                    rel_node = ET.Element(qn(PKG_REL_NS, "Relationship"))
                    rel_node.set("Id", new_rid)
                    rel_node.set("Type", source_rel.attrib.get("Type", ""))
                    rel_node.set("Target", new_target)
                    target_rels_root.append(rel_node)
                    target_relationships.append(rel_node)
                    file_map[posixpath.join("word", new_target)] = src_zip.read(source_part_name)
                    ensure_content_type(content_types_root, new_target)
                    copied_targets.append(new_target)
                    if node.attrib.get(qn(R_NS, "embed")) == old_rid:
                        node.set(qn(R_NS, "embed"), new_rid)
                    if node.attrib.get(qn(R_NS, "id")) == old_rid:
                        node.set(qn(R_NS, "id"), new_rid)
                    if node.attrib.get(qn(R_NS, "link")) == old_rid:
                        node.set(qn(R_NS, "link"), new_rid)
            else:
                if source_index is None or int(source_index) >= len(source_tables):
                    skipped.append({**candidate, "reason": "source_table_out_of_range"})
                    continue
                cloned_node = clone_xml(source_tables[int(source_index)])

            target_body.insert(insertion_index, cloned_node)
            inserted_caption_texts.add(caption_text)
            attached.append(
                {
                    **candidate,
                    "copiedTargets": copied_targets,
                    "insertionMode": "before_caption",
                }
            )

        file_map["word/document.xml"] = ET.tostring(target_doc_root, encoding="utf-8", xml_declaration=True)
        file_map["word/_rels/document.xml.rels"] = ET.tostring(target_rels_root, encoding="utf-8", xml_declaration=True)
        file_map["[Content_Types].xml"] = ET.tostring(content_types_root, encoding="utf-8", xml_declaration=True)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for name, payload in file_map.items():
            out_zip.writestr(name, payload)

    return build_attachment_report(
        source_docx,
        intermediate_docx,
        output_docx,
        attached,
        manual_review,
        skipped_by_policy,
        skipped,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Apply high-confidence hybrid asset candidates to an intermediate DOCX.")
    parser.add_argument("--source-docx", required=True, help="Original DOCX with preserveable assets")
    parser.add_argument("--intermediate-docx", required=True, help="Text-first rebuilt intermediate DOCX")
    parser.add_argument("--thesis-ir-json", required=True, help="Path to thesis_ir.json")
    parser.add_argument("--output-docx", required=True, help="Output DOCX path")
    parser.add_argument("--report-json", required=True, help="Attachment report output path")
    args = parser.parse_args()

    thesis_ir = json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8"))
    report = apply_candidates(
        Path(args.source_docx),
        Path(args.intermediate_docx),
        thesis_ir,
        Path(args.output_docx),
    )
    Path(args.report_json).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
