from __future__ import annotations

import json
from typing import Any

from formatter_core.workspace import Workspace, file_digest


def write_summary(workspace: Workspace, manifest: dict[str, Any], deliverables: dict[str, Any]) -> str:
    summary = {
        "status": manifest.get("status"), "backend": manifest.get("backend", "word"),
        "projectName": workspace.project, "profile": workspace.profile,
        "checkPolicy": manifest.get("checkPolicy"), "visualReviewed": False,
        "inputPreserved": file_digest(workspace.source) == manifest.get("inputSha256"),
        "qualityGate": manifest.get("qualityGate"),
        "deferredStructuralChanges": manifest.get("deferredStructuralChanges") or [],
        "semanticCorrections": manifest.get("semanticCorrections"),
        "formattingReview": manifest.get("formattingReview") or [],
        "deliverables": workspace.portable_paths(deliverables),
        "backendResults": workspace.portable_paths(manifest.get("backendResults") or {}),
        "limitations": ["No Agent visual review was performed; structural checks do not certify pagination or appearance."],
    }
    if not summary["inputPreserved"]:
        raise RuntimeError("Original input changed during formatting")
    destination = workspace.reports / "summary.json"
    destination.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(destination)
