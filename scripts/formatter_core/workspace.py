from __future__ import annotations

import hashlib
import json
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


def safe_name(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value).strip(" .")
    cleaned = re.sub(r"\s+", "-", cleaned)[:64].rstrip(" .") or "thesis"
    if cleaned.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{index}" for index in range(1, 10)), *(f"LPT{index}" for index in range(1, 10))}:
        cleaned = f"thesis-{cleaned}"
    return cleaned


def file_digest(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


@dataclass(frozen=True)
class Workspace:
    root: Path
    project: str
    profile: str
    source: Path

    @classmethod
    def create(cls, source: Path, profile: str, output_base: Path | None = None, project_name: str | None = None) -> Workspace:
        source = source.resolve(strict=True)
        if not source.is_file():
            raise ValueError(f"Input is not a file: {source}")
        project = safe_name(project_name or source.stem)
        base = (output_base or source.parent / "thesis-output").resolve()
        run_id = datetime.now().astimezone().strftime("%Y%m%d-%H%M%S-%f") + "-" + uuid.uuid4().hex[:6]
        root = base / project / run_id
        root.mkdir(parents=True, exist_ok=False)
        workspace = cls(root.resolve(), project, safe_name(profile), source)
        for directory in (workspace.work, workspace.reports, workspace.deliverables):
            directory.mkdir()
        return workspace

    @property
    def work(self) -> Path:
        return self.root / "work"

    @property
    def reports(self) -> Path:
        return self.root / "reports"

    @property
    def deliverables(self) -> Path:
        return self.root / "deliverables"

    @property
    def manifest_path(self) -> Path:
        return self.root / "manifest.json"

    def filename(self, backend: str, extension: str) -> str:
        return f"{self.project}__{self.profile}__{safe_name(backend)}.{extension.lstrip('.')}"

    def relative(self, path: Path) -> str:
        resolved = path.resolve()
        if not resolved.is_relative_to(self.root):
            raise ValueError(f"Artifact escapes workspace: {path}")
        return resolved.relative_to(self.root).as_posix()

    def move(self, source: Path, destination: Path) -> Path:
        self.relative(source)
        self.relative(destination)
        if source.resolve() == self.source:
            raise ValueError("Cannot move the original input")
        if destination.exists():
            raise FileExistsError(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(source), str(destination))
        return destination

    def publish(self, source: Path, backend: str, extension: str) -> Path:
        return self.move(source, self.deliverables / self.filename(backend, extension))

    def portable_paths(self, mapping: dict[str, object]) -> dict[str, object]:
        result: dict[str, object] = {}
        for key, value in mapping.items():
            if isinstance(value, dict):
                result[key] = self.portable_paths(value)
            elif isinstance(value, list):
                result[key] = [self.portable_paths({"value": item})["value"] for item in value]
            elif isinstance(value, str) and Path(value).is_absolute() and Path(value).resolve().is_relative_to(self.root):
                result[key] = self.relative(Path(value))
            else:
                result[key] = value
        return result

    def remap(self, value: object, relocations: dict[str, str]) -> object:
        if isinstance(value, dict):
            return {key: self.remap(item, relocations) for key, item in value.items()}
        if isinstance(value, list):
            return [self.remap(item, relocations) for item in value]
        if isinstance(value, str):
            return relocations.get(value, value)
        return value

    def rewrite_references(self, relocations: dict[str, str]) -> None:
        for path in self.root.rglob("*.json"):
            if path == self.manifest_path:
                continue
            try:
                original = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            rewritten = self.remap(original, relocations)
            if rewritten != original:
                path.write_text(json.dumps(rewritten, ensure_ascii=True, indent=2), encoding="utf-8")

    def write_manifest(self, payload: dict[str, object]) -> None:
        manifest = dict(payload)
        manifest["schemaVersion"] = "1.0"
        manifest["runId"] = self.root.name
        manifest["projectName"] = self.project
        if "inputSha256" not in manifest:
            manifest["inputSha256"] = file_digest(self.source)
        manifest["visualReviewed"] = False
        for key in ("artifacts", "deliverables", "backendResults"):
            if isinstance(manifest.get(key), dict):
                manifest[key] = self.portable_paths(manifest[key])
        manifest["artifactIndex"] = [
            {"path": self.relative(path), "role": "deliverable" if path.is_relative_to(self.deliverables) else "report" if path.is_relative_to(self.reports) else "intermediate"}
            for path in sorted(self.root.rglob("*"))
            if path.is_file() and path != self.manifest_path
        ]
        temporary = self.root / "manifest.json.tmp"
        temporary.write_text(json.dumps(manifest, ensure_ascii=True, indent=2), encoding="utf-8")
        temporary.replace(self.manifest_path)
