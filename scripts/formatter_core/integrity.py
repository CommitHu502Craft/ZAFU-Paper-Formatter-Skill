"""Minimal sufficient content/asset checks independent of formatting audits."""
from __future__ import annotations
import hashlib
from pathlib import Path
import posixpath
import re
from collections import Counter
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

def snapshot(path):
    with zipfile.ZipFile(path) as package:
        files = {name: package.read(name) for name in package.namelist()}
    root = ET.fromstring(files["word/document.xml"])
    return {
        "text": "".join(node.text or "" for node in root.iter(W + "t")),
        "mathText": "".join(node.text or "" for node in root.iter(M + "t")),
        "objects": Counter(node.tag for node in root.iter() if node.tag.rsplit("}", 1)[-1] in {"tbl", "drawing", "object", "pict", "oMath", "footnoteReference", "endnoteReference", "commentReference", "ins", "del"}),
        "assets": Counter(hashlib.sha256(data).hexdigest() for name, data in files.items() if name.startswith(("word/media/", "word/embeddings/")) or name in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}),
        "files": files,
        "paragraphs": ["".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p")],
    }

def verify(before, after, *, exact_text=True):
    old, new = snapshot(before), snapshot(after)
    issues = []
    if exact_text and (old["text"] != new["text"] or old["mathText"] != new["mathText"]):
        issues.append({"type": "content_integrity_failure", "reason": "text_sequence_changed"})
    for key in ("objects", "assets"):
        missing = old[key] - new[key]
        if missing:
            issues.append({"type": "asset_integrity_failure", "kind": key, "missing": dict(missing)})
    return {"executed": True, "passed": not issues, "issues": issues}

def verify_generated(before, after):
    result = verify(before, after, exact_text=False)
    old, new = snapshot(before), snapshot(after)
    canonical = lambda text: re.sub(r"\s+", "", text)
    target = canonical(new["text"])
    cursor = 0
    missing = []
    for text in old["paragraphs"]:
        # Existing presentation rule removes leading numeric bibliography labels.
        token = canonical(re.sub(r"^\s*\[\d+\]\s*", "", text))
        if token:
            position = target.find(token, cursor)
            if position < 0:
                missing.append(text[:120])
            else:
                cursor = position + len(token)
    if missing or old["mathText"] != new["mathText"]:
        result["issues"].append({"type": "content_integrity_failure", "reason": "generated_source_text_missing", "paragraphs": missing})
    result["passed"] = not result["issues"]
    return result

def verify_source_assets(ir, built):
    state = snapshot(built)
    blocks = ir.get("semanticBlocks") or []
    issues = []
    for kind, tag in (("image", W + "drawing"), ("table", W + "tbl")):
        expected = sum(block.get("kind") == kind for block in blocks)
        actual = state["objects"][tag]
        if actual < expected:
            issues.append({"type": "asset_integrity_failure", "kind": kind, "expected": expected, "actual": actual, "reason": "source_asset_was_not_built"})
    return {"executed": True, "passed": not issues, "issues": issues}

def relationship_issues(path):
    """Check every internal relationship, including embedded object dependencies."""
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        issues = []
        for name in sorted(names):
            if not name.endswith(".rels"):
                continue
            root = ET.fromstring(package.read(name))
            base = posixpath.dirname(posixpath.dirname(name))
            for rel in root:
                if rel.get("TargetMode") == "External":
                    continue
                from urllib.parse import unquote
                target = unquote(rel.get("Target", "")).split("#", 1)[0]
                resolved = posixpath.normpath(posixpath.join(base, target)) if not target.startswith("/") else target.lstrip("/")
                if resolved not in names:
                    issues.append({"part": name, "relationshipId": rel.get("Id"), "target": resolved})
        return issues
