#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import zipfile
import xml.etree.ElementTree as ET

from docx_ooxml import NS, parse_document, qn


def child_index_until_paragraph_count(body: ET.Element, paragraph_count: int) -> int:
    seen = 0
    for idx, child in enumerate(list(body)):
        if child.tag == qn("p"):
            if seen >= paragraph_count:
                return idx
            seen += 1
    return len(list(body))


def build_front_matter_template(source_docx: str, output_docx: str) -> None:
    audit = parse_document(source_docx)
    immutable_prefix = ((audit.get("preservationHints") or {}).get("immutablePrefix")) or {}
    prefix_end = int(immutable_prefix.get("endParagraphExclusive") or 0)
    if prefix_end <= 0:
        raise SystemExit("No immutable front-matter prefix detected in source document.")

    with zipfile.ZipFile(source_docx) as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    document_root = ET.fromstring(files["word/document.xml"])
    body = document_root.find("w:body", NS)
    if body is None:
        raise SystemExit("word/document.xml has no w:body")

    paragraphs = body.findall("w:p", NS)
    if prefix_end > len(paragraphs):
        raise SystemExit("Detected prefix exceeds paragraph count.")

    boundary = child_index_until_paragraph_count(body, prefix_end)
    children = list(body)
    for child in children[boundary:]:
        body.remove(child)

    tail_sectpr = None
    last_prefix_para = paragraphs[prefix_end - 1]
    last_ppr = last_prefix_para.find("w:pPr", NS)
    if last_ppr is not None and last_ppr.find("w:sectPr", NS) is not None:
        tail_sectpr = copy.deepcopy(last_ppr.find("w:sectPr", NS))
        last_ppr.remove(last_ppr.find("w:sectPr", NS))

    existing_tail = body.find("w:sectPr", NS)
    if existing_tail is not None:
        body.remove(existing_tail)
    if tail_sectpr is not None:
        body.append(tail_sectpr)

    files["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a reusable front-matter template from a reference DOCX.")
    parser.add_argument("source_docx", help="Reference DOCX containing cover and integrity pages")
    parser.add_argument("output_docx", help="Output front-matter template DOCX")
    args = parser.parse_args()
    build_front_matter_template(args.source_docx, args.output_docx)


if __name__ == "__main__":
    main()
