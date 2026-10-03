from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
import urllib.request
from pathlib import Path
from typing import Any


REPOSITORY = "Stolorzs/ZafuTemplatePublic"
REVISION = "81c0d127d485aaaf00c5cbccc5d0d29affd446f2"
CLASS_SHA256 = "c731c74248efacf9edb7bf72cb6613dd3f1193a81f52be5655c931032c42111e"
CLASS_URL = f"https://raw.githubusercontent.com/{REPOSITORY}/{REVISION}/thesis/ZafuThesis.cls"
ROOT = Path(__file__).resolve().parents[2]
MAX_CLASS_BYTES = 1024 * 1024
COVER_ASSETS = {
    "zafu.png": "ccd4540334cacc07ba9d61cff6a5f38df0514ba116aa25a1f523202662363792",
    "zafuLogo.jpg": "2929792646f91742eccc6088fe89da655ec3c40a948f2653ccd4773bfe76cb3d",
}


def install_verified_resource(url: str, expected_hash: str, cached: Path, destination: Path) -> None:
    cached.parent.mkdir(parents=True, exist_ok=True)
    if cached.exists():
        data = cached.read_bytes()
    else:
        with urllib.request.urlopen(url, timeout=30) as response:
            data = response.read(5 * MAX_CLASS_BYTES + 1)
    if len(data) > 5 * MAX_CLASS_BYTES or hashlib.sha256(data).hexdigest() != expected_hash:
        raise ValueError(f"External resource integrity check failed: {cached.name}")
    if not cached.exists():
        with tempfile.NamedTemporaryFile(dir=cached.parent, suffix=".tmp", delete=False) as temporary:
            temporary.write(data)
            temporary_path = Path(temporary.name)
        try:
            os.replace(temporary_path, cached)
        finally:
            temporary_path.unlink(missing_ok=True)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(cached, destination)


def install_cover_assets(project: Path) -> list[dict[str, str]]:
    cache = ROOT / ".local/dependencies/zafu-template"
    cache.mkdir(parents=True, exist_ok=True)
    figures = project / "figures"
    figures.mkdir(exist_ok=True)
    assets = []
    for name, expected_hash in COVER_ASSETS.items():
        cached = cache / name
        url = f"https://raw.githubusercontent.com/{REPOSITORY}/{REVISION}/thesis/figures/{name}"
        install_verified_resource(url, expected_hash, cached, figures / name)
        assets.append({"path": f"figures/{name}", "sha256": expected_hash, "sourceUrl": url})
    return assets


def install_external_class(project: Path, cache: Path | None = None) -> dict[str, Any]:
    cache = cache if cache is not None else ROOT / ".local/dependencies/zafu-template"
    cache.mkdir(parents=True, exist_ok=True)
    cached_class = cache / "ZafuThesis.cls"
    if cached_class.exists():
        data = cached_class.read_bytes()
    else:
        request = urllib.request.Request(CLASS_URL, headers={"User-Agent": "Paper-Formatter-Skill"})
        with urllib.request.urlopen(request, timeout=30) as response:
            data = response.read(MAX_CLASS_BYTES + 1)
    if len(data) > MAX_CLASS_BYTES or hashlib.sha256(data).hexdigest() != CLASS_SHA256:
        raise ValueError("ZafuThesis.cls integrity check failed; inspect or remove the local cache before retrying")
    if not cached_class.exists():
        with tempfile.NamedTemporaryFile(dir=cache, suffix=".tmp", delete=False) as temporary:
            temporary.write(data)
            temporary_path = Path(temporary.name)
        try:
            os.replace(temporary_path, cached_class)
        finally:
            temporary_path.unlink(missing_ok=True)
    shutil.copyfile(cached_class, project / "ZafuThesis.cls")
    provenance = {
        "repository": REPOSITORY, "revision": REVISION, "sourceUrl": CLASS_URL,
        "path": "ZafuThesis.cls", "sha256": CLASS_SHA256, "modified": False,
        "license": "Upstream redistribution license not confirmed; not covered by this project's Apache-2.0 license.",
    }
    (project / "UPSTREAM.json").write_text(json.dumps(provenance, indent=2), encoding="utf-8")
    return provenance
