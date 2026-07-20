#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List


ALLOWED_FORMATS = {"docx", "markdown", "text"}
ALLOWED_KINDS = {"heading", "paragraph", "caption", "reference", "image", "table", "equation", "page_break"}


def validate(payload: Dict[str, Any]) -> Dict[str, Any]:
    errors: List[Dict[str, Any]] = []
    warnings: List[Dict[str, Any]] = []
    if payload.get("schema") != "paper-formatter.thesis-ir":
        errors.append({"kind": "schema", "message": "Unexpected or missing ThesisIR schema identifier."})
    if str(payload.get("version") or "") != "2.0":
        errors.append({"kind": "version", "message": "ThesisIR version must be 2.0."})

    source = payload.get("source") or {}
    source_format = source.get("format")
    if source_format not in ALLOWED_FORMATS:
        errors.append({"kind": "source_format", "value": source_format})
    if not source.get("path"):
        errors.append({"kind": "source_path", "message": "Source path is required."})

    blocks = payload.get("semanticBlocks")
    if not isinstance(blocks, list):
        errors.append({"kind": "semantic_blocks", "message": "semanticBlocks must be a list."})
        blocks = []
    ids = set()
    for index, block in enumerate(blocks):
        if not isinstance(block, dict):
            errors.append({"kind": "block_type", "index": index})
            continue
        block_id = block.get("id")
        if not block_id or block_id in ids:
            errors.append({"kind": "block_id", "index": index, "value": block_id})
        ids.add(block_id)
        if block.get("kind") not in ALLOWED_KINDS:
            errors.append({"kind": "block_kind", "index": index, "value": block.get("kind")})
        if not isinstance(block.get("sourceAnchor"), dict):
            errors.append({"kind": "source_anchor", "index": index})
        confidence = block.get("confidence")
        if not isinstance(confidence, (int, float)) or not 0 <= confidence <= 1:
            errors.append({"kind": "confidence", "index": index, "value": confidence})
        if block.get("kind") in {"heading", "paragraph", "caption", "reference"} and not block.get("text"):
            warnings.append({"kind": "empty_text_block", "index": index, "id": block_id})

    if payload.get("semanticBlockCount") != len(blocks):
        errors.append(
            {
                "kind": "semantic_block_count",
                "declared": payload.get("semanticBlockCount"),
                "actual": len(blocks),
            }
        )
    for required in ("frontMatter", "headingTree", "bodyBlocks", "references", "numberingAnalysis"):
        if required not in payload:
            errors.append({"kind": "legacy_compatibility_view", "field": required})

    capabilities = source.get("capabilities") or {}
    if source_format == "docx" and not capabilities.get("wordNativeStructure"):
        errors.append({"kind": "docx_capabilities", "message": "DOCX must retain Word-native evidence capability."})
    if source_format != "docx" and capabilities.get("wordNativeStructure"):
        errors.append({"kind": "text_capabilities", "message": "Text inputs cannot claim Word-native structure."})

    return {
        "schema": payload.get("schema"),
        "version": payload.get("version"),
        "sourceFormat": source_format,
        "semanticBlockCount": len(blocks),
        "passed": not errors,
        "errorCount": len(errors),
        "warningCount": len(warnings),
        "errors": errors,
        "warnings": warnings,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate unified ThesisIR v2.")
    parser.add_argument("thesis_ir_json")
    parser.add_argument("--output", "-o")
    args = parser.parse_args()
    payload = json.loads(Path(args.thesis_ir_json).read_text(encoding="utf-8"))
    report = validate(payload)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)
    if not report["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
