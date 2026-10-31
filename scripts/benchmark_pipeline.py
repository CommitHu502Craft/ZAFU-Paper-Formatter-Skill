#!/usr/bin/env python3
"""Reproducible cold/warm end-to-end runs; stage times are inclusive."""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import platform

ROOT = Path(__file__).resolve().parents[1]
CASES = {
    "A_markdown": "tests/fixtures/unified_markdown/input.md",
    "B_normal_docx": "profiles/hybrid_fixture/template.docx",
    "C_messy_docx": "tests/samples/messy_student_thesis.docx",
    "D_native_assets": "tests/fixtures/docx_hybrid_mixed_confidence_assets/input.docx",
}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--label", required=True)
    parser.add_argument("--check", choices=["structural", "visual"], default="structural")
    parser.add_argument("--cases", nargs="+", choices=list(CASES), help="Subset of fixtures; default runs all four")
    parser.add_argument("--refresh-results", action="store_true", help="Backfill manifest metadata for an existing measurement; never re-time it")
    args = parser.parse_args()
    base = Path(args.output_dir).resolve() / args.label
    if args.refresh_results:
        path = base / "results.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        for result in data["results"]:
            manifests = sorted((base / result["case"] / "runs").rglob("manifest.json"), key=lambda p: p.stat().st_mtime_ns)
            if len(manifests) != 2:
                raise ValueError("Refresh requires exactly two original run manifests per case")
            manifest_path = manifests[0 if result["temperature"] == "cold" else 1]
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            result.update({"manifest": str(manifest_path), "status": manifest.get("status"), "qualityGate": manifest.get("qualityGate"), "inputSha256": manifest.get("inputSha256")})
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        return
    if base.exists():
        raise ValueError("Benchmark label already exists; use a new label for reproducible cold/warm runs")
    results = []
    for case, relative in CASES.items():
        if args.cases and case not in args.cases:
            continue
        source = ROOT / relative
        if not source.is_file():
            results.append({"case": case, "status": "missing_fixture", "source": relative})
            continue
        for temperature in ("cold", "warm"):
            out = base / case / temperature
            perf = out / "instrumentation"
            out.mkdir(parents=True, exist_ok=True)
            env = dict(os.environ, THESIS_PERF_DIR=str(perf), PYTHONIOENCODING="utf-8")
            env["PYTHONPATH"] = str(ROOT / "tests/perf_support") + os.pathsep + env.get("PYTHONPATH", "")
            start = time.perf_counter()
            with (out / "run.log").open("w", encoding="utf-8") as log:
                run = subprocess.run([sys.executable, str(ROOT / "scripts/thesis_format.py"), str(source), "--check", args.check, "--output-dir", str(base / case / "runs")], cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT)
            elapsed = time.perf_counter() - start
            stages, count = {}, 0
            for item in perf.glob("*.json"):
                data = json.loads(item.read_text(encoding="utf-8"))
                count += data["externalProcesses"]
                for key, value in data["stages"].items():
                    stage = stages.setdefault(key, {"seconds": 0, "calls": 0})
                    for field in stage:
                        stage[field] += value[field]
            manifests = sorted((base / case / "runs").rglob("manifest.json"), key=lambda p: p.stat().st_mtime_ns)
            manifest = json.loads(manifests[-1].read_text(encoding="utf-8")) if manifests else {}
            results.append({"case": case, "temperature": temperature, "source": relative, "inputSha256": manifest.get("inputSha256"), "seconds": elapsed, "exitCode": run.returncode, "externalProcesses": count, "stages": stages, "status": manifest.get("status"), "qualityGate": manifest.get("qualityGate"), "visualRepairRounds": int(bool(manifest.get("visualRefinement"))), "manifest": str(manifests[-1]) if manifests else None})
            print(json.dumps(results[-1], ensure_ascii=True), flush=True)
    (base / "results.json").write_text(json.dumps({"label": args.label, "check": args.check, "python": platform.python_version(), "platform": platform.platform(), "notes": "Inclusive instrumented timings; cold=fresh case cache, warm=repeat process with shared template cache; OS disk cache not flushed. Missing stages were not executed. Times are not additive because parse/extract/validation overlap.", "results": results}, indent=2), encoding="utf-8")

if __name__ == "__main__":
    main()
