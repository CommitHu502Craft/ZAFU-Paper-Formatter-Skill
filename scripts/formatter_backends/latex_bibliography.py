from __future__ import annotations

import re
import json
from pathlib import Path
from typing import Any

import bibtexparser
from bibtexparser.bparser import BibTexParser

from formatter_core.workspace import file_digest
from formatter_backends.external_class import ROOT, install_verified_resource


CITATION = re.compile(r"\[@[A-Za-z0-9][A-Za-z0-9_.:+/-]*(?:\s*;\s*@[A-Za-z0-9][A-Za-z0-9_.:+/-]*)*\]")
CITATION_INLINE = re.compile(r"\$\$[\s\S]*?\$\$|(?<!\\)\$[^$\n]+?\$|\\\([\s\S]*?\\\)|\\\[[\s\S]*?\\\]|\*\*[^*]+\*\*|(?<!\*)\*[^*\n]+\*|`[^`\n]+`|" + CITATION.pattern)
SAFE_KEY = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.:+/-]*$")
SAFE_COMMANDS = {"&", "%", "_", "#", "$", "{", "}", "'", '"', "`", "^", "~", "=", ".", "c", "v", "u", "H", "o", "O", "l", "L", "ae", "AE", "oe", "OE", "ss", "textit", "textbf", "emph", "url", "LaTeX", "TeX"}
STYLE_REVISION = "f7e42092047bebfdec0d2dedf46b61eab2195fb1"
STYLE_RESOURCES = {
    "gbt7714-author-year.bst": "3b0753310c397e7c24d62ebe2ca83c3bd726bb777cfb72adaa4993979c13ee69",
    "LICENSE": "5f05fcf6ef25a6c31bccd2df7c0c46b23107bbeb2ce5cdba74efb5cc357f4dbb",
    "gbt7714.dtx": "9e890731335174b10486a264bfcdc348df367e8a127b0e1f5112657b6f0089f2",
    "gbt7714.ins": "bed4b632f1258955fd008bea256843fd609a33ab3cf5a0af7bab50040c0640b9",
}


def install_bibliography_style(project: Path) -> dict[str, Any]:
    resources = []
    for filename, digest in STYLE_RESOURCES.items():
        url = f"https://raw.githubusercontent.com/zepinglee/gbt7714-bibtex-style/{STYLE_REVISION}/{filename}"
        relative = filename if filename.endswith(".bst") else f"upstream/gbt7714/{filename}"
        install_verified_resource(url, digest, ROOT / ".local/dependencies/gbt7714-2015" / filename, project / relative)
        resources.append({"path": relative, "sha256": digest, "sourceUrl": url})
    provenance = {"repository": "zepinglee/gbt7714-bibtex-style", "revision": STYLE_REVISION, "version": "2.1.4", "standard": "GB/T 7714-2015", "license": "LPPL-1.3c-or-later", "modified": False, "resources": resources}
    original_style = (project / "gbt7714-author-year.bst").read_text(encoding="utf-8")
    replacements = {"#1 'year.after.author :=": "#0 'year.after.author :=", "#0 'period.after.author :=": "#1 'period.after.author :="}
    derived_style = original_style
    for original, replacement in replacements.items():
        if derived_style.count(original) != 1:
            raise ValueError("Pinned BibTeX style configuration no longer matches the verified adapter")
        derived_style = derived_style.replace(original, replacement)
    derived_path = project / "zafu-gbt7714-2015-authoryear.bst"
    derived_path.write_text("%% Generated school variant; see UPSTREAM-BIBTEX.json for source and exact changes.\n" + derived_style, encoding="utf-8")
    provenance["schoolVariant"] = {"path": derived_path.name, "sha256": file_digest(derived_path), "modified": True, "changes": replacements, "purpose": "Keep author-year citation labels but place publication year in the bibliography as prescribed by the user's school examples."}
    (project / "UPSTREAM-BIBTEX.json").write_text(json.dumps(provenance, indent=2), encoding="utf-8")
    return provenance


def prepare_bibliography(path: Path | None, project: Path) -> dict[str, Any]:
    if path is None:
        return {"enabled": False, "keys": [], "errors": [], "warnings": []}
    parser = BibTexParser(common_strings=True)
    parser.ignore_nonstandard_types = False
    data = path.read_text(encoding="utf-8-sig")
    database = bibtexparser.loads(data, parser=parser)
    entries = database.entries
    errors = []
    warnings = []
    keys = [str(entry.get("ID") or "") for entry in entries]
    declared = re.findall(r"@([A-Za-z]+)\s*[{(]", data)
    if not entries or len(entries) != sum(kind.lower() not in {"comment", "string", "preamble"} for kind in declared):
        errors.append({"kind": "invalid_bibliography", "message": "Every BibTeX record must parse completely."})
    if len(keys) != len(set(keys)) or any(not SAFE_KEY.fullmatch(key) for key in keys):
        errors.append({"kind": "invalid_or_duplicate_citation_key"})
    if database.preambles:
        errors.append({"kind": "unsafe_bibliography_preamble"})
    for entry in entries:
        for field, value in entry.items():
            if any(command not in SAFE_COMMANDS for command in re.findall(r"\\([A-Za-z]+|.)", str(value))):
                errors.append({"kind": "unsafe_bibliography_command", "key": entry["ID"], "field": field})
        if not entry.get("title") or not (entry.get("author") or entry.get("editor")) or not (entry.get("year") or entry.get("date")):
            warnings.append(f"Bibliography record {entry['ID']} lacks title, responsibility or publication year; metadata was not invented.")
        if re.search(r"[\u4e00-\u9fff]", entry.get("author", "")) and not entry.get("key"):
            errors.append({"kind": "missing_chinese_sort_key", "key": entry["ID"], "message": "Supply the author's verified pinyin in the BibTeX key field for Chinese sorting."})
    if not errors:
        (project / "references.bib").write_text(bibtexparser.dumps(database), encoding="utf-8")
    return {"enabled": True, "keys": keys, "entryCount": len(entries), "sourceSha256": file_digest(path), "errors": errors, "warnings": warnings, "style": "GB/T 7714-2015 author-year", "citationManagement": "latex-only"}


def render_citations(text: str, keys: set[str], used: set[str], render, issues: list[dict[str, Any]], identifier: str) -> str:
    chunks = []
    cursor = 0
    for match in CITATION_INLINE.finditer(text):
        chunks.append(render(text[cursor:match.start()]))
        token = match.group()
        if token.startswith("**"):
            chunks.append(r"\textbf{" + render_citations(token[2:-2], keys, used, render, issues, identifier) + "}")
        elif token.startswith("*"):
            chunks.append(r"\emph{" + render_citations(token[1:-1], keys, used, render, issues, identifier) + "}")
        elif not CITATION.fullmatch(token):
            chunks.append(render(token))
        else:
            cited = re.findall(r"@([A-Za-z0-9][A-Za-z0-9_.:+/-]*)", match.group())
            missing = sorted(set(cited) - keys)
            if missing:
                issues.append({"kind": "missing_citation_key", "blockId": identifier, "keys": missing})
                chunks.append(render(match.group()))
            else:
                used.update(cited)
                chunks.append(r"\citep{" + ",".join(cited) + "}")
        cursor = match.end()
    chunks.append(render(text[cursor:]))
    return "".join(chunks)
