#!/usr/bin/env python3
"""Draft a visual_refinement_plan.json from text/geometry findings alone.

This closes the review loop for agents WITHOUT image-reading ability: the
findings in visual_review_manifest.json (produced from PDF text boxes and
page-thumbnail geometry, no model needed) are mapped to whitelisted
refinement actions with conservative confidence rules.

Only findings that can be anchored to a concrete paragraph via text matching
are converted; everything else stays a manual-review note. The output plan is
reviewable/editable before being applied by apply_visual_refinements.py.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def paragraph_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS))


def _squash(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def load_paragraph_texts(docx_path: Path) -> List[str]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    return [paragraph_text(p) for p in root.iter(f"{{{W_NS}}}p")]


def find_paragraph_by_text(paragraph_texts: List[str], wanted: str) -> Optional[int]:
    """Unique-match lookup; ambiguous or missing text returns None."""
    target = _squash(wanted)[:24]
    if len(target) < 4:
        return None
    matches = [i for i, text in enumerate(paragraph_texts) if _squash(text).startswith(target)]
    return matches[0] if len(matches) == 1 else None


def suggest_actions(
    findings: List[Dict[str, Any]],
    paragraph_texts: List[str],
) -> Dict[str, Any]:
    actions: List[Dict[str, Any]] = []
    unmapped: List[Dict[str, Any]] = []
    for finding in findings:
        kind = str(finding.get("type") or "")
        if kind == "heading_orphan_at_page_bottom":
            heading_text = str(finding.get("headingText") or "")
            index = find_paragraph_by_text(paragraph_texts, heading_text)
            if index is not None:
                actions.append(
                    {
                        "type": "set_keep_with_next",
                        "target": {"paragraphIndex": index, "textPrefix": heading_text[:16]},
                        "reason": f"第 {finding.get('page')} 页末尾的标题与后续正文分离",
                        "confidence": 0.85,
                    }
                )
                continue
        elif kind == "caption_possibly_split_from_asset":
            caption_text = str(finding.get("captionText") or "")
            index = find_paragraph_by_text(paragraph_texts, caption_text)
            if index is not None and index > 0:
                actions.append(
                    {
                        "type": "set_keep_with_next",
                        "target": {"paragraphIndex": index - 1},
                        "reason": f"第 {finding.get('page')} 页题注与图表分离,前一段保持与题注同页",
                        "confidence": 0.7,
                    }
                )
                continue
        unmapped.append(finding)
    return {"actions": actions, "unmappedFindings": unmapped}


def main() -> None:
    parser = argparse.ArgumentParser(description="Suggest a visual refinement plan from render-validation findings (no image reading required).")
    parser.add_argument("docx", help="The rendered DOCX (typically output/repaired.docx)")
    parser.add_argument("--manifest-json", required=True, help="visual_review_manifest.json path")
    parser.add_argument("--output", "-o", required=True, help="Where to write the draft visual_refinement_plan.json")
    args = parser.parse_args()

    manifest = json.loads(Path(args.manifest_json).read_text(encoding="utf-8"))
    findings = manifest.get("visualFindings") or []
    paragraph_texts = load_paragraph_texts(Path(args.docx))
    plan = suggest_actions(findings, paragraph_texts)
    plan["source"] = "suggest_visual_refinements(text_and_geometry_only)"
    Path(args.output).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"suggestedActions": len(plan["actions"]), "unmappedFindings": len(plan["unmappedFindings"])}, ensure_ascii=False))


if __name__ == "__main__":
    main()
