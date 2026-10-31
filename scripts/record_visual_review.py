#!/usr/bin/env python3
"""Record actual Agent/manual image inspection against a delivered candidate.

This command does not perform a visual review. Call only AFTER opening the
overview and every exported critical/anomaly page. Evidence must name the
reviewer and include SHA256 for every inspected image and the DOCX candidate.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from formatter_core.workspace import file_digest

def record_review(run_dir, evidence):
    run_dir = Path(run_dir).resolve()
    manifest_path = run_dir / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    def artifact(value):
        candidate = (run_dir / value).resolve()
        if not candidate.is_relative_to(run_dir):
            raise ValueError("Review artifact escapes run workspace")
        return candidate
    docx = artifact(manifest["deliverables"]["docx"])
    if not evidence.get("reviewer") or evidence.get("candidateSha256") != file_digest(docx):
        raise ValueError("Reviewer missing or review refers to a different DOCX candidate")
    render_path = artifact(manifest["artifacts"]["renderValidationReport"])
    render = json.loads(render_path.read_text(encoding="utf-8"))
    if not render.get("candidateSha256") or render["candidateSha256"] != evidence["candidateSha256"]:
        raise ValueError("Render belongs to a different candidate")
    required = [*render.get("contactSheets", []), *[item["imagePath"] for item in render.get("criticalPages", []) if item.get("imagePath")]]
    if not required or any(item.get("status") != "exported" for item in render.get("criticalPages", [])):
        raise ValueError("Complete overview/critical images are required")
    reviewed = {str(artifact(item["path"])): item.get("sha256") for item in evidence.get("reviewedImages", [])}
    for image in required:
        path = artifact(image)
        if reviewed.get(str(path)) != file_digest(path):
            raise ValueError(f"Missing or stale inspected image: {path.name}")
    unresolved = []
    for finding in render.get("visualFindings", []):
        key = f"{finding.get('type')}:{finding.get('page')}"
        disposition = (evidence.get("issueDispositions") or {}).get(key) or {}
        if disposition.get("status") != "dismissed" or not disposition.get("reason"):
            unresolved.append(finding)
    unresolved.extend(evidence.get("additionalIssues") or [])
    acceptance = manifest.setdefault("visualAcceptance", {})
    acceptance.update({"rendered": True, "previewGenerated": True, "visualReviewed": True, "visualReviewStatus": "reviewed_with_issues" if unresolved else "reviewed", "unresolvedVisualIssues": unresolved, "candidateSha256": evidence["candidateSha256"]})
    manifest["visualReviewed"] = True
    manifest["deliveryStatus"] = "candidate_with_visual_issues" if unresolved else "accepted"
    manifest.setdefault("backendResults", {}).setdefault("word", {}).update({"visualReviewed": True, "visualAcceptance": acceptance})
    evidence_path = run_dir / "reports/visual_review_record.json"
    evidence_path.write_text(json.dumps(evidence, ensure_ascii=True, indent=2), encoding="utf-8")
    manifest["artifacts"]["visualReviewRecord"] = evidence_path.relative_to(run_dir).as_posix()
    for destination in (render_path, artifact(manifest["artifacts"]["summary"])):
        report = json.loads(destination.read_text(encoding="utf-8"))
        report.update({"visualReviewed": True, "visualAcceptance": acceptance, "visualReviewStatus": acceptance["visualReviewStatus"], "unresolvedVisualIssues": unresolved, "deliveryStatus": manifest["deliveryStatus"]})
        if "limitations" in report:
            report["limitations"] = ["Review covers the recorded renderer and candidate; other editors may paginate differently."]
        destination.write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    if render.get("visualReviewManifestPath"):
        preview_manifest_path = artifact(render["visualReviewManifestPath"])
        preview_manifest = json.loads(preview_manifest_path.read_text(encoding="utf-8"))
        preview_manifest.update({"visualReviewed": True, "visualReviewStatus": acceptance["visualReviewStatus"], "unresolvedVisualIssues": unresolved, "visualReviewRecord": str(evidence_path)})
        preview_manifest_path.write_text(json.dumps(preview_manifest, ensure_ascii=True, indent=2), encoding="utf-8")
    temp = manifest_path.with_suffix(".json.tmp")
    temp.write_text(json.dumps(manifest, ensure_ascii=True, indent=2), encoding="utf-8")
    temp.replace(manifest_path)
    from generate_review_report import main as build_html
    import sys
    previous = sys.argv
    try:
        sys.argv = ["generate_review_report.py", "--output-dir", str(run_dir / "work"), "--report-html", str(run_dir / "reports/summary.html")]
        build_html()
    finally:
        sys.argv = previous
    return acceptance

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_dir")
    parser.add_argument("--record", required=True)
    args = parser.parse_args()
    print(json.dumps(record_review(args.run_dir, json.loads(Path(args.record).read_text(encoding="utf-8"))), ensure_ascii=True))

if __name__ == "__main__":
    main()
