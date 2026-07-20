#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document

from build_docx_from_markdown import (
    build_plain_text_docx_from_thesis_ir,
    clear_document,
    force_section_a4,
    set_update_fields_on_open,
)
from docx_ooxml import load_rules


def main() -> None:
    parser = argparse.ArgumentParser(description="Build an intermediate DOCX directly from ThesisIR.")
    parser.add_argument("--thesis-ir-json", required=True, help="Path to thesis_ir.json")
    parser.add_argument("--rules-yaml", required=True, help="Rules YAML")
    parser.add_argument("--output-docx", required=True, help="Output DOCX path")
    args = parser.parse_args()

    thesis_ir = json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8"))
    rules = load_rules(args.rules_yaml)

    document = Document()
    clear_document(document)
    set_update_fields_on_open(document)
    force_section_a4(document.sections[0])

    state = {
        "toc_inserted": False,
        "pending_page_break": False,
        "is_first_block": True,
        "last_block_was_page_break": False,
        "insert_toc_for_markdown": False,
        "chapter_page_breaks": False,
        "reuse_last_empty_paragraph": False,
    }
    build_plain_text_docx_from_thesis_ir(document, thesis_ir, rules, state)
    Path(args.output_docx).parent.mkdir(parents=True, exist_ok=True)
    document.save(args.output_docx)


if __name__ == "__main__":
    main()
