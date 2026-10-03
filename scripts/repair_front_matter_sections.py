#!/usr/bin/env python3
"""Repair the front-matter section boundary without changing thesis text."""
from __future__ import annotations

import argparse
import json
import zipfile
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS}


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def paragraph_text(node: ET.Element) -> str:
    return "".join(item.text or "" for item in node.findall(".//w:t", NS))


def paragraph_properties(node: ET.Element) -> ET.Element:
    ppr = node.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(qn("pPr"))
        node.insert(0, ppr)
    return ppr


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source_docx")
    parser.add_argument("output_docx")
    args = parser.parse_args()

    source = Path(args.source_docx).resolve()
    output = Path(args.output_docx).resolve()
    with zipfile.ZipFile(source, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    ET.register_namespace("w", W_NS)
    ET.register_namespace("r", R_NS)
    root = ET.fromstring(files["word/document.xml"])
    paragraphs = list(root.iter(qn("p")))

    abstract_index = next(
        (i for i, node in enumerate(paragraphs) if paragraph_text(node).strip() == "Abstract"),
        None,
    )
    body_index = next(
        (
            i
            for i, node in enumerate(paragraphs)
            if paragraph_text(node).strip() == "1 引言" and (abstract_index is None or i > abstract_index)
        ),
        None,
    )
    keywords_index = next(
        (
            i
            for i, node in enumerate(paragraphs)
            if abstract_index is not None
            and i > abstract_index
            and (body_index is None or i < body_index)
            and paragraph_text(node).strip().startswith("Keywords:")
        ),
        None,
    )
    if abstract_index is None or keywords_index is None or body_index is None:
        raise SystemExit(
            json.dumps(
                {
                    "status": "blocked",
                    "reason": "front_matter_anchor_not_found",
                    "abstractIndex": abstract_index,
                    "keywordsIndex": keywords_index,
                    "bodyIndex": body_index,
                },
                ensure_ascii=False,
            )
        )

    abstract_ppr = paragraph_properties(paragraphs[abstract_index])
    section = abstract_ppr.find("w:sectPr", NS)
    if section is None:
        raise SystemExit(json.dumps({"status": "blocked", "reason": "abstract_section_boundary_not_found"}))

    abstract_ppr.remove(section)
    keep_next = abstract_ppr.find("w:keepNext", NS)
    if keep_next is None:
        ET.SubElement(abstract_ppr, qn("keepNext"))

    keywords_ppr = paragraph_properties(paragraphs[keywords_index])
    old_section = keywords_ppr.find("w:sectPr", NS)
    if old_section is not None:
        keywords_ppr.remove(old_section)
    keywords_ppr.append(deepcopy(section))

    files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(
        json.dumps(
            {
                "status": "ok",
                "abstractParagraphIndex": abstract_index,
                "keywordsParagraphIndex": keywords_index,
                "bodyParagraphIndex": body_index,
                "movedSectionBoundary": True,
                "keepNextApplied": True,
                "outputDocx": str(output),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
