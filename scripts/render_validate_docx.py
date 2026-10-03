#!/usr/bin/env python3
"""Render a DOCX to PDF and run a semantic + visual review pass.

Outputs (under --output-dir):
  <stem>.pdf                    rendered PDF (backend permitting)
  semantic_pages.json           per-page text summaries + located regions
  critical_pages/               hi-res PNG for located key pages (+/- 1 page)
  all_pages/                    low-res PNG thumbnails for every page
  contact_sheet.png             one or more grid sheets with page numbers
  visual_review_manifest.json   everything an agent needs for visual review
  render_validation_report.json legacy-compatible report (when run as CLI)

Pages are located by page text markers, never by fixed page-number guessing;
page-number fallbacks are used only when a marker is genuinely absent.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional

from render_pdf_backends import convert_docx_to_pdf, detect_libreoffice


def _json_text(payload: Any) -> str:
    """Serialize render diagnostics when PDF extraction returns lone surrogates."""
    return json.dumps(payload, ensure_ascii=True, indent=2)


def detect_pdftoppm() -> Optional[str]:
    return shutil.which("pdftoppm")


def detect_pdftotext() -> Optional[str]:
    return shutil.which("pdftotext")


# --------------------------------------------------------------------------
# Page text extraction
# --------------------------------------------------------------------------

def extract_page_texts(pdf_path: Path) -> List[str]:
    """Return one text string per page; prefer pypdf, fall back to pdftotext."""
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(pdf_path))
        return [(page.extract_text() or "") for page in reader.pages]
    except Exception:
        pass

    pdftotext = detect_pdftotext()
    if not pdftotext:
        return []
    completed = subprocess.run(
        [pdftotext, "-layout", str(pdf_path), "-"],
        check=False,
        capture_output=True,
    )
    if completed.returncode != 0:
        return []
    text = completed.stdout.decode("utf-8", errors="replace")
    return text.split("\f")


def extract_page_text_boxes(pdf_path: Path) -> List[Dict[str, Any]]:
    """Per-page text fragments with coordinates via pypdf's visitor API.

    Returns one entry per page: {"width", "height", "fragments": [{"text",
    "x", "y"}]} where y is measured from the page BOTTOM (PDF space).
    Empty list when pypdf is unavailable or the PDF resists extraction.
    """
    try:
        from pypdf import PdfReader
    except ImportError:
        return []
    pages: List[Dict[str, Any]] = []
    try:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            box = page.mediabox
            fragments: List[Dict[str, Any]] = []

            def visitor(text: str, cm, tm, font_dict, font_size) -> None:
                stripped = (text or "").strip()
                if not stripped:
                    return
                fragments.append({"text": stripped[:80], "x": round(float(tm[4]), 1), "y": round(float(tm[5]), 1)})

            page.extract_text(visitor_text=visitor)
            pages.append(
                {
                    "width": float(box.width),
                    "height": float(box.height),
                    "fragments": fragments,
                }
            )
    except Exception:
        return []
    return pages


def count_pdf_pages(pdf_path: Path) -> Optional[int]:
    if not pdf_path.exists():
        return None
    try:
        from pypdf import PdfReader

        return len(PdfReader(str(pdf_path)).pages)
    except Exception:
        data = pdf_path.read_bytes()
        matches = re.findall(rb"/Type\s*/Page\b", data)
        return len(matches) or None


# --------------------------------------------------------------------------
# Semantic page location
# --------------------------------------------------------------------------

def _squash(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


REGION_MARKERS: List[Dict[str, Any]] = [
    {
        "key": "integrity_statement",
        "label": "诚信承诺书",
        "patterns": [r"诚信承诺", r"承诺书", r"原创性声明", r"学位论文版权"],
        "searchIn": "head",
    },
    {
        "key": "toc",
        "label": "目录",
        "patterns": [r"^目录", r"^目次"],
        "searchIn": "head",
    },
    {
        "key": "cn_abstract",
        "label": "中文摘要",
        "patterns": [r"摘要[:：]?", r"^中文摘要"],
        "searchIn": "head",
    },
    {
        "key": "en_abstract",
        "label": "英文摘要",
        "patterns": [r"(?i)abstract"],
        "searchIn": "head",
    },
    {
        "key": "body_start",
        "label": "正文首页",
        "patterns": [
            r"^第[一二三四五六七八九十1-9]\s*[章部分]",
            r"^1\s*[一-鿿]",
            r"^引言", r"^绪论", r"^前言",
        ],
        "searchIn": "head",
    },
    {
        "key": "references",
        "label": "参考文献",
        "patterns": [r"^参考文献", r"(?i)^references"],
        "searchIn": "head",
    },
    {
        "key": "acknowledgements",
        "label": "致谢",
        "patterns": [r"^致谢", r"^致\s*谢", r"(?i)^acknowledg"],
        "searchIn": "head",
    },
    {
        "key": "appendix",
        "label": "附录",
        "patterns": [r"^附录", r"(?i)^appendix"],
        "searchIn": "head",
    },
]


def _looks_like_toc_entry(line: str) -> bool:
    """TOC entry lines carry dot leaders and/or trailing page numbers."""
    if re.search(r"[.…·]{5,}", line):
        return True
    if re.search(r"[.…·]{3,}\s*\d+\s*$", line):
        return True
    return bool(re.match(r"^\S{1,24}\s+\d{1,3}\s*$", line.strip()))


def locate_semantic_pages(page_texts: List[str]) -> Dict[str, Any]:
    """Locate regions by markers in the first lines of each page."""
    page_count = len(page_texts)
    page_heads: List[List[str]] = []
    for text in page_texts:
        lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
        page_heads.append(lines[:6])

    regions: Dict[str, Dict[str, Any]] = {}

    if page_count:
        regions["cover"] = {"label": "封面", "page": 1, "method": "first_page", "confidence": 0.9}
        regions["last_page"] = {"label": "末页", "page": page_count, "method": "last_page", "confidence": 0.95}

    claimed: Dict[int, str] = {}
    for marker in REGION_MARKERS:
        key = marker["key"]
        found = None
        for page_index, head_lines in enumerate(page_heads):
            page_number = page_index + 1
            if page_number == 1 and key not in {"integrity_statement"}:
                # Region headings never live on the cover.
                continue
            for line in head_lines:
                if key != "toc" and _looks_like_toc_entry(line):
                    continue
                normalized = _squash(line)
                for pattern in marker["patterns"]:
                    if re.search(pattern, line) or re.search(pattern, normalized):
                        found = {
                            "label": marker["label"],
                            "page": page_number,
                            "method": "text_marker",
                            "matchedText": line[:80],
                            "confidence": 0.9,
                        }
                        break
                if found:
                    break
            if found:
                # cn_abstract marker also matches TOC lines like "摘要...1"; require
                # the page not already claimed by toc and the match near line start.
                prior = claimed.get(found["page"])
                if prior and prior != key:
                    found["confidence"] = 0.6
                    found["note"] = f"page_shared_with_{prior}"
                claimed.setdefault(found["page"], key)
                break
        if found:
            regions[key] = found

    # en_abstract found on the same page as cn_abstract usually means the
    # "Abstract" word only appears inside the Chinese abstract page's footer
    # keywords; keep it but lower confidence.
    cn = regions.get("cn_abstract")
    en = regions.get("en_abstract")
    if cn and en and cn.get("page") == en.get("page"):
        en["confidence"] = min(float(en.get("confidence") or 0.9), 0.55)
        en["note"] = "same_page_as_cn_abstract"

    # Page-number fallbacks only where markers failed.
    fallbacks: List[str] = []
    if page_count:
        def fallback(key: str, label: str, page: int) -> None:
            if key not in regions and 1 <= page <= page_count:
                regions[key] = {
                    "label": label,
                    "page": page,
                    "method": "fallback_estimate",
                    "confidence": 0.3,
                }
                fallbacks.append(key)

        toc_page = (regions.get("toc") or {}).get("page")
        body_page = (regions.get("body_start") or {}).get("page")
        ref_page = (regions.get("references") or {}).get("page")
        fallback("toc", "目录", 2 if page_count >= 2 else 1)
        if body_page is None:
            anchor = toc_page or 2
            fallback("body_start", "正文首页", min(int(anchor) + 3, page_count))
        if ref_page is None:
            fallback("references", "参考文献", max(page_count - 2, 1))

    return {
        "pageCount": page_count,
        "regions": regions,
        "fallbackRegions": fallbacks,
        "unlocatedRegions": [m["key"] for m in REGION_MARKERS if m["key"] not in regions],
    }


def build_page_summaries(page_texts: List[str]) -> List[Dict[str, Any]]:
    summaries = []
    for index, text in enumerate(page_texts):
        stripped = (text or "").strip()
        lines = [line.strip() for line in stripped.splitlines() if line.strip()]
        summaries.append(
            {
                "page": index + 1,
                "charCount": len(_squash(stripped)),
                "lineCount": len(lines),
                "firstLines": lines[:3],
            }
        )
    return summaries


# --------------------------------------------------------------------------
# Page image export
# --------------------------------------------------------------------------

def export_pages_png(
    pdf_path: Path,
    out_dir: Path,
    first: int,
    last: int,
    dpi: int,
    prefix: str,
    timeout_sec: int,
) -> List[Path]:
    pdftoppm = detect_pdftoppm()
    if not pdftoppm:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    base = out_dir / prefix
    command = [
        pdftoppm, "-png", "-r", str(dpi),
        "-f", str(first), "-l", str(last),
        str(pdf_path), str(base),
    ]
    completed = subprocess.run(command, check=False, capture_output=True, timeout=timeout_sec)
    if completed.returncode != 0:
        return []
    return sorted(out_dir.glob(f"{prefix}-*.png")) or sorted(out_dir.glob(f"{prefix}*.png"))


def export_all_page_thumbnails(pdf_path: Path, out_dir: Path, page_count: int, timeout_sec: int) -> List[Path]:
    if not page_count:
        return []
    return export_pages_png(pdf_path, out_dir, 1, page_count, dpi=50, prefix="page", timeout_sec=timeout_sec)


def export_critical_pages(
    pdf_path: Path,
    out_dir: Path,
    semantic: Dict[str, Any],
    page_count: int,
    timeout_sec: int,
) -> List[Dict[str, Any]]:
    """Hi-res export of each located region page plus its neighbors."""
    pdftoppm = detect_pdftoppm()
    if not pdftoppm or not page_count:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    wanted: Dict[int, List[str]] = {}
    for key, region in (semantic.get("regions") or {}).items():
        page = region.get("page")
        if not isinstance(page, int):
            continue
        for neighbor in (page - 1, page, page + 1):
            if 1 <= neighbor <= page_count:
                tag = key if neighbor == page else f"{key}_ctx"
                wanted.setdefault(neighbor, []).append(tag)

    entries: List[Dict[str, Any]] = []
    for page in sorted(wanted):
        labels = wanted[page]
        primary = next((l for l in labels if not l.endswith("_ctx")), labels[0])
        prefix = f"{primary}_p{page:03d}"
        images = export_pages_png(pdf_path, out_dir, page, page, dpi=140, prefix=prefix, timeout_sec=timeout_sec)
        entries.append(
            {
                "page": page,
                "labels": labels,
                "imagePath": str(images[0]) if images else None,
                "status": "exported" if images else "failed",
            }
        )
    return entries


def build_contact_sheets(
    thumbnails: List[Path],
    out_dir: Path,
    columns: int = 5,
    rows: int = 6,
) -> List[Path]:
    """Grid sheets of numbered page thumbnails; multiple sheets for long docs."""
    if not thumbnails:
        return []
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return []

    per_sheet = columns * rows
    cell_w, cell_h, pad, label_h = 240, 340, 8, 22
    sheets: List[Path] = []
    for sheet_index in range(0, len(thumbnails), per_sheet):
        batch = thumbnails[sheet_index : sheet_index + per_sheet]
        sheet_w = columns * (cell_w + pad) + pad
        used_rows = (len(batch) + columns - 1) // columns
        sheet_h = used_rows * (cell_h + label_h + pad) + pad
        sheet = Image.new("RGB", (sheet_w, sheet_h), "white")
        draw = ImageDraw.Draw(sheet)
        for offset, thumb_path in enumerate(batch):
            page_number = sheet_index + offset + 1
            col, row = offset % columns, offset // columns
            x = pad + col * (cell_w + pad)
            y = pad + row * (cell_h + label_h + pad)
            try:
                thumb = Image.open(thumb_path)
                thumb.thumbnail((cell_w, cell_h))
                sheet.paste(thumb, (x + (cell_w - thumb.width) // 2, y + (cell_h - thumb.height) // 2))
                draw.rectangle([x, y, x + cell_w, y + cell_h], outline="#999999")
            except OSError:
                draw.rectangle([x, y, x + cell_w, y + cell_h], outline="#cc0000")
            draw.text((x + cell_w // 2 - 14, y + cell_h + 4), f"p.{page_number}", fill="black")
        suffix = "" if sheet_index == 0 else f"_{sheet_index // per_sheet + 1}"
        sheet_path = out_dir / f"contact_sheet{suffix}.png"
        sheet.save(sheet_path)
        sheets.append(sheet_path)
    return sheets


# --------------------------------------------------------------------------
# Lightweight visual heuristics (warnings only, never delivery blockers)
# --------------------------------------------------------------------------

def analyze_page_geometry(thumbnails: List[Path]) -> List[Dict[str, Any]]:
    """Pixel-level warnings from low-res page thumbnails."""
    try:
        from PIL import Image
    except ImportError:
        return []
    findings: List[Dict[str, Any]] = []
    blank_streak: List[int] = []
    for index, thumb_path in enumerate(thumbnails):
        page = index + 1
        try:
            img = Image.open(thumb_path).convert("L")
        except OSError:
            continue
        w, h = img.size
        pixels = img.load()
        dark_rows = []
        for y in range(h):
            dark = sum(1 for x in range(0, w, 2) if pixels[x, y] < 176)
            dark_rows.append(dark)
        total_dark = sum(dark_rows)
        content_rows = [y for y, dark in enumerate(dark_rows) if dark >= 2]
        if total_dark < w * h * 0.0008 or not content_rows:
            findings.append({"type": "near_blank_page", "page": page, "severity": "warning"})
            blank_streak.append(page)
            continue
        if len(blank_streak) >= 2:
            findings.append({
                "type": "consecutive_blank_pages",
                "pages": list(blank_streak),
                "severity": "warning",
            })
        blank_streak = []
        top_ratio = content_rows[0] / h
        bottom_ratio = (h - 1 - content_rows[-1]) / h
        if top_ratio > 0.45:
            findings.append({"type": "large_top_whitespace", "page": page, "topWhitespaceRatio": round(top_ratio, 2), "severity": "warning"})
        if bottom_ratio < 0.02:
            findings.append({"type": "content_touches_bottom_edge", "page": page, "severity": "warning"})
        if top_ratio < 0.015:
            findings.append({"type": "content_touches_top_edge", "page": page, "severity": "warning"})
    if len(blank_streak) >= 2:
        findings.append({"type": "consecutive_blank_pages", "pages": list(blank_streak), "severity": "warning"})
    return findings


def analyze_text_geometry(
    text_boxes: List[Dict[str, Any]],
    margins_cm: Optional[Dict[str, float]] = None,
) -> List[Dict[str, Any]]:
    """Coordinate-level warnings from PDF text boxes (no image reading).

    Checks content against the expected printable area and catches heading
    orphans by their true vertical position instead of line order alone.
    """
    findings: List[Dict[str, Any]] = []
    if not text_boxes:
        return findings
    cm_to_pt = 72.0 / 2.54
    margins = margins_cm or {}
    top_margin = float(margins.get("top") or 2.0) * cm_to_pt
    bottom_margin = float(margins.get("bottom") or 2.0) * cm_to_pt
    side_margin = float(margins.get("left") or 2.0) * cm_to_pt
    heading_re = re.compile(r"^(第[一二三四五六七八九十1-9]\s*[章部分]|\d+(?:\.\d+){0,2}\s+\S)")

    # pypdf x-coordinates are not always absolute (transform-dependent), so the
    # left-edge check is calibrated against the document's own typical left
    # edge rather than the nominal margin.
    per_page_min_x: List[float] = []
    for page in text_boxes:
        xs = [float(f.get("x") or 0) for f in (page.get("fragments") or []) if len(str(f.get("text") or "")) > 4]
        if xs:
            per_page_min_x.append(min(xs))
    typical_left = sorted(per_page_min_x)[len(per_page_min_x) // 2] if per_page_min_x else 0.0

    for index, page in enumerate(text_boxes):
        page_number = index + 1
        width = float(page.get("width") or 595)
        height = float(page.get("height") or 842)
        fragments = page.get("fragments") or []
        if not fragments:
            continue
        top_limit = height - top_margin * 0.4
        for fragment in fragments:
            y = float(fragment.get("y") or 0)
            text = str(fragment.get("text") or "")
            if y > top_limit and len(text) > 1:
                findings.append({
                    "type": "content_above_top_margin",
                    "page": page_number,
                    "text": text[:40],
                    "severity": "warning",
                })
                break
        page_min_xs = [float(f.get("x") or 0) for f in fragments if len(str(f.get("text") or "")) > 4]
        if page_min_xs and typical_left > 20 and min(page_min_xs) < typical_left - side_margin * 0.5:
            offender = next(f for f in fragments if float(f.get("x") or 0) == min(page_min_xs))
            findings.append({
                "type": "content_outside_left_margin",
                "page": page_number,
                "text": str(offender.get("text") or "")[:40],
                "severity": "warning",
            })
        # Heading orphan by geometry: heading fragment in the lowest ~12% of
        # the text area with nothing below it on the same page.
        body_fragments = [f for f in fragments if len(str(f.get("text") or "")) > 1]
        if body_fragments:
            lowest_y = min(float(f.get("y") or 0) for f in body_fragments)
            for fragment in body_fragments:
                y = float(fragment.get("y") or 0)
                text = str(fragment.get("text") or "")
                if heading_re.match(text) and y <= bottom_margin + (height - top_margin - bottom_margin) * 0.12:
                    below = [f for f in body_fragments if float(f.get("y") or 0) < y - 2]
                    if not below and index + 1 < len(text_boxes):
                        findings.append({
                            "type": "heading_orphan_at_page_bottom_geometry",
                            "page": page_number,
                            "headingText": text[:60],
                            "yFromBottomPt": round(y, 1),
                            "severity": "warning",
                        })
                    break

    # Deduplicate per (type, page).
    seen = set()
    unique: List[Dict[str, Any]] = []
    for finding in findings:
        key = (finding.get("type"), finding.get("page"))
        if key in seen:
            continue
        seen.add(key)
        unique.append(finding)
    return unique


def analyze_text_layout(
    page_texts: List[str],
    semantic: Dict[str, Any],
    source_page_hint: Optional[int],
) -> List[Dict[str, Any]]:
    findings: List[Dict[str, Any]] = []
    page_count = len(page_texts)
    heading_re = re.compile(r"^(第[一二三四五六七八九十1-9]\s*[章部分]|\d+(?:\.\d+){0,2}\s+\S)")
    caption_re = re.compile(r"^[图表]\s*\d")
    for index, text in enumerate(page_texts):
        lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
        if not lines:
            continue
        # Heading stranded at the bottom of a page: last non-empty line is a
        # heading and the next page continues with body text.
        if heading_re.match(lines[-1]) and index + 1 < page_count:
            next_lines = [l.strip() for l in (page_texts[index + 1] or "").splitlines() if l.strip()]
            if next_lines and not heading_re.match(next_lines[0]):
                findings.append({
                    "type": "heading_orphan_at_page_bottom",
                    "page": index + 1,
                    "headingText": lines[-1][:60],
                    "severity": "warning",
                })
        if caption_re.match(lines[0]) and index > 0:
            findings.append({
                "type": "caption_possibly_split_from_asset",
                "page": index + 1,
                "captionText": lines[0][:60],
                "severity": "warning",
            })
    if source_page_hint and page_count and page_count > max(source_page_hint * 1.8, source_page_hint + 12):
        findings.append({
            "type": "page_count_explosion",
            "renderedPages": page_count,
            "sourcePageHint": source_page_hint,
            "severity": "warning",
        })
    for key in semantic.get("unlocatedRegions") or []:
        if key in {"integrity_statement", "appendix", "acknowledgements", "en_abstract"}:
            continue  # genuinely optional in many theses
        findings.append({"type": "key_region_not_located", "region": key, "severity": "warning"})
    return findings


# --------------------------------------------------------------------------
# Main entry
# --------------------------------------------------------------------------

def suggested_review_pages(semantic: Dict[str, Any]) -> Dict[str, Optional[int]]:
    regions = semantic.get("regions") or {}

    def page_of(key: str) -> Optional[int]:
        value = (regions.get(key) or {}).get("page")
        return int(value) if isinstance(value, int) else None

    return {
        "cover": page_of("cover"),
        "toc": page_of("toc"),
        "abstract": page_of("cn_abstract"),
        "firstBodyPage": page_of("body_start"),
        "references": page_of("references"),
    }


def run_render_validation(
    docx_path: str,
    output_dir: str,
    timeout_sec: int = 180,
    source_page_hint: Optional[int] = None,
    *,
    analyze_layout: bool = True,
    export_previews: bool = False,
) -> Dict[str, Any]:
    source = Path(docx_path).resolve()
    render_dir = Path(output_dir).resolve()
    render_dir.mkdir(parents=True, exist_ok=True)

    report: Dict[str, Any] = {
        "sourceDocx": str(source),
        "outputDir": str(render_dir),
        "available": False,
        "backend": None,
        "backendPath": None,
        "status": "unavailable",
        "reason": None,
        "pdfPath": None,
        "pageCount": None,
        "previewDir": str(render_dir / "critical_pages"),
        "suggestedReviewPages": {"cover": None, "toc": None, "abstract": None, "firstBodyPage": None, "references": None},
        "semanticPagesPath": None,
        "contactSheets": [],
        "criticalPages": [],
        "visualFindings": [],
        "visualReviewManifestPath": None,
        "previewExport": {"available": False, "backend": None, "backendPath": None, "manifestPath": None, "images": [], "reason": "Page previews were not requested or conversion was unavailable"},
        "layoutAnalyzed": False,
        "previewsRequested": export_previews,
        "visualReviewed": False,
    }

    conversion = convert_docx_to_pdf(source, render_dir, timeout_sec=timeout_sec)
    report["available"] = bool(conversion.get("available"))
    report["backend"] = conversion.get("backend")
    report["backendPath"] = conversion.get("backendPath")
    report["conversionAttempts"] = conversion.get("attempts")
    if not conversion.get("available"):
        report["reason"] = conversion.get("reason")
        return report

    pdf_path = Path(conversion["pdfPath"])
    report["pdfPath"] = str(pdf_path)
    return analyze_pdf_document(pdf_path, render_dir, report=report, timeout_sec=timeout_sec, source_page_hint=source_page_hint, analyze_layout=analyze_layout, export_previews=export_previews)


def analyze_pdf_document(
    pdf_path: Path,
    render_dir: Path,
    *,
    report: Optional[Dict[str, Any]] = None,
    timeout_sec: int = 180,
    source_page_hint: Optional[int] = None,
    analyze_layout: bool = True,
    export_previews: bool = False,
) -> Dict[str, Any]:
    render_dir.mkdir(parents=True, exist_ok=True)
    if report is None:
        report = {"available": True, "pdfPath": str(pdf_path), "contactSheets": [], "criticalPages": [], "visualFindings": [], "layoutAnalyzed": False, "previewsRequested": export_previews, "visualReviewed": False}
    page_count = count_pdf_pages(pdf_path) or 0
    report["pageCount"] = page_count

    if not analyze_layout and not export_previews:
        report["status"] = "ok"
        return report

    page_texts = extract_page_texts(pdf_path)
    semantic = locate_semantic_pages(page_texts)
    semantic["pageSummaries"] = build_page_summaries(page_texts)
    semantic_path = render_dir / "semantic_pages.json"
    semantic_path.write_text(_json_text(semantic), encoding="utf-8")
    report["semanticPagesPath"] = str(semantic_path)
    report["suggestedReviewPages"] = suggested_review_pages(semantic)

    critical_entries = export_critical_pages(pdf_path, render_dir / "critical_pages", semantic, page_count, timeout_sec) if export_previews else []
    report["criticalPages"] = critical_entries

    thumbnails = export_all_page_thumbnails(pdf_path, render_dir / "all_pages", page_count, timeout_sec) if export_previews else []
    sheets = build_contact_sheets(thumbnails, render_dir) if thumbnails else []
    report["contactSheets"] = [str(p) for p in sheets]

    findings = analyze_page_geometry(thumbnails)
    findings.extend(analyze_text_layout(page_texts, semantic, source_page_hint))
    text_boxes = extract_page_text_boxes(pdf_path)
    geometry_findings = analyze_text_geometry(text_boxes)
    existing_orphan_pages = {f.get("page") for f in findings if str(f.get("type", "")).startswith("heading_orphan")}
    for finding in geometry_findings:
        if finding.get("type") == "heading_orphan_at_page_bottom_geometry" and finding.get("page") in existing_orphan_pages:
            continue
        findings.append(finding)
    report["visualFindings"] = findings
    report["layoutAnalyzed"] = True

    manifest = {
        "pdfPath": str(pdf_path),
        "pageCount": page_count,
        "semanticPages": str(semantic_path),
        "contactSheets": [str(p) for p in sheets],
        "criticalPages": critical_entries,
        "allPagesDir": str(render_dir / "all_pages") if thumbnails else None,
        "visualFindings": findings,
        "regions": semantic.get("regions"),
        "unlocatedRegions": semantic.get("unlocatedRegions"),
        "visualReviewed": False,
        "reviewInstructions": [
            "有读图能力时:先看 contact_sheet.png 检查整体分页(空白页、异常留白、标题落单),再看 critical_pages/ 的封面、目录、摘要、正文首页和参考文献页。",
            "没有读图能力时:不要尝试读图。直接依据本文件的 visualFindings 与 regions/unlocatedRegions(均来自 PDF 文本和页面几何,无需看图),或运行 scripts/suggest_visual_refinements.py 自动起草修复计划;并把 contact_sheet.png 与 critical_pages/ 的路径告诉用户请其人工浏览。",
            "发现明显排版问题时,生成 visual_refinement_plan.json 并执行一次 apply_visual_refinements.py。",
            "visualFindings 中的项目均为 warning,不阻止交付;确认无误后即可交付最终文件。",
        ],
    }
    manifest_path = render_dir / "visual_review_manifest.json"
    manifest_path.write_text(_json_text(manifest), encoding="utf-8")
    report["visualReviewManifestPath"] = str(manifest_path)

    report["status"] = "ok"
    report["reason"] = None
    # Legacy compatibility with prior previewExport consumers.
    report["previewExport"] = {
        "available": bool(critical_entries),
        "backend": "pdftoppm" if critical_entries else None,
        "backendPath": detect_pdftoppm(),
        "manifestPath": str(manifest_path),
        "images": [
            {
                "label": (entry.get("labels") or ["page"])[0],
                "page": entry.get("page"),
                "status": entry.get("status"),
                "imagePath": entry.get("imagePath"),
            }
            for entry in critical_entries
        ],
        "reason": None if critical_entries else "pdftoppm not found in PATH",
    }
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a DOCX to PDF, locate key pages semantically, and export review images.")
    parser.add_argument("docx", help="DOCX file to render")
    parser.add_argument("--output-dir", required=True, help="Directory for PDF and render artifacts")
    parser.add_argument("--output", "-o", help="Write JSON report to this path")
    parser.add_argument("--check", choices=["pdf", "layout", "visual"], default="layout", help="Page images are generated only with visual")
    parser.add_argument("--timeout-sec", type=int, default=180, help="Renderer timeout in seconds")
    parser.add_argument("--source-page-hint", type=int, help="Approximate page count of the source document for explosion detection")
    args = parser.parse_args()

    report = run_render_validation(args.docx, args.output_dir, timeout_sec=args.timeout_sec, source_page_hint=args.source_page_hint, analyze_layout=args.check != "pdf", export_previews=args.check == "visual")
    text = _json_text(report)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
