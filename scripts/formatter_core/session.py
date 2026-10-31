"""Single-run, content-addressed analysis reuse. Cached results are never mutable aliases."""
from __future__ import annotations
import copy
import hashlib
import json
from pathlib import Path
from functools import wraps
from contextlib import contextmanager
import importlib
import sys
import time

_active = None

class AnalysisSession:
    def __init__(self, cache_dir: Path | None = None):
        self.cache_dir = cache_dir
        self.memory = {}
        self.metrics = {"cacheHits": 0, "cacheMisses": 0, "steps": []}

    def key(self, name, paths, options):
        digest = hashlib.sha256(name.encode())
        # Invalidate on implementation or profile changes, including transitive rule helpers.
        for path in [*paths, *sorted(Path(__file__).parents[1].glob("*.py")), *sorted(Path(__file__).parent.glob("*.py"))]:
            if path:
                path = Path(path)
                digest.update(path.read_bytes())
        digest.update(repr(options).encode())
        return digest.hexdigest()

    def get(self, name, paths, options, compute, persistent=False):
        key = self.key(name, paths, options)
        if key in self.memory:
            self.metrics["cacheHits"] += 1
            return copy.deepcopy(self.memory[key])
        destination = self.cache_dir / f"{key}.json" if persistent and self.cache_dir else None
        if destination and destination.is_file():
            try:
                value = json.loads(destination.read_text(encoding="utf-8"))
                self.metrics["cacheHits"] += 1
            except (OSError, ValueError):
                value = compute()
                self.metrics["cacheMisses"] += 1
        else:
            value = compute()
            self.metrics["cacheMisses"] += 1
        self.memory[key] = copy.deepcopy(value)
        if destination:
            destination.parent.mkdir(parents=True, exist_ok=True)
            # Concurrent writers publish complete values via unique temporary files.
            import uuid
            temp = destination.with_suffix(f".{uuid.uuid4().hex}.tmp")
            temp.write_text(json.dumps(value, ensure_ascii=True), encoding="utf-8")
            temp.replace(destination)
        return value

@contextmanager
def analysis_session(cache_dir=None):
    global _active
    previous = _active
    session = AnalysisSession(cache_dir)
    _active = session
    try:
        yield session
    finally:
        _active = previous

def cached_analysis(name, path_args, *, persistent=False):
    def decorate(function):
        import inspect
        signature = inspect.signature(function)
        @wraps(function)
        def call(*args, **kwargs):
            if _active is None:
                return function(*args, **kwargs)
            bound = signature.bind(*args, **kwargs)
            bound.apply_defaults()
            paths = [bound.arguments.get(key) for key in path_args]
            options = tuple((key, value) for key, value in bound.arguments.items() if key not in path_args and not key.startswith("_"))
            # Reports carry the document's source identity, while equivalent rules
            # copied into a new workspace should still share the template cache.
            if bound.arguments.get("docx_path"):
                options += (("sourceIdentity", str(Path(bound.arguments["docx_path"]).resolve())),)
            return _active.get(name, paths, options, lambda: function(*args, **kwargs), persistent)
        return call
    return decorate

def run_in_process(command, root):
    """Compatibility adapter for existing CLI main functions; no child interpreter."""
    previous = sys.argv
    start = time.perf_counter()
    try:
        sys.argv = command[1:]
        module = importlib.import_module(Path(command[1]).stem)
        try:
            module.main()
        except SystemExit as error:
            if error.code not in (None, 0):
                raise
    finally:
        sys.argv = previous
        if _active:
            _active.metrics["steps"].append({"step": Path(command[1]).stem, "seconds": time.perf_counter() - start})
