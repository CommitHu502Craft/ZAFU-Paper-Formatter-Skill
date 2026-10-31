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
import xml.etree.ElementTree as ET
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
    pdftotext = detect_pdftotext()
    if pdftotext:
        try:
            completed = subprocess.run([pdftotext, "-bbox-layout", "-enc", "UTF-8", str(pdf_path), "-"], capture_output=True, check=False, timeout=60)
            if completed.returncode == 0:
                root = ET.fromstring(completed.stdout)
                pages = []
                for node in root.iter():
                    if node.tag.rsplit("}", 1)[-1] != "page":
                        continue
                    height = float(node.get("height"))
                    fragments = []
                    for line in node.iter():
                        if line.tag.rsplit("}", 1)[-1] != "line":
                            continue
                        words = [word for word in line if word.tag.rsplit("}", 1)[-1] == "word"]
                        text = " ".join("".join(word.itertext()) for word in words)
                        if not text.strip():
                            continue
                        x0, x1 = float(line.get("xMin")), float(line.get("xMax"))
                        y0, y1 = height - float(line.get("yMax")), height - float(line.get("yMin"))
                        fragments.append({"text": text, "x": x0, "y": y0, "fontSize": y1 - y0, "boundingBox": [x0, y0, x1, y1], "coordinateAccuracy": "pdf_text_bbox"})
                    pages.append({"width": float(node.get("width")), "height": height, "fragments": fragments})
                if pages:
                    return pages
        except (OSError, ValueError, ET.ParseError, subprocess.TimeoutExpired):
            pass
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
                x = float(tm[4]) * cm[0] + float(tm[5]) * cm[2] + cm[4]
                y = float(tm[4]) * cm[1] + float(tm[5]) * cm[3] + cm[5]
                fragments.append({"text": stripped, "x": round(x, 1), "y": round(y, 1), "fontSize": float(font_size), "boundingBox": [round(x, 1), round(y, 1), None, round(y + float(font_size), 1)], "coordinateAccuracy": "approximate_baseline_no_glyph_width"})

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
    # Render consecutive pages as one range instead of one Poppler process per page.
    page_images = {}
    ordered = sorted(wanted)
    ranges = []
    for page in ordered:
        if ranges and page == ranges[-1][1] + 1:
            ranges[-1][1] = page
        else:
            ranges.append([page, page])
    for first, last in ranges:
        images = export_pages_png(pdf_path, out_dir, first, last, dpi=140, prefix=f"critical_{first}_{last}", timeout_sec=timeout_sec)
        for image in images:
            match = re.search(r"-(\d+)\.png$", image.name)
            if match:
                page_images[int(match.group(1))] = image
    for page in ordered:
        labels = wanted[page]
        primary = next((l for l in labels if not l.endswith("_ctx")), labels[0])
        images = [page_images[page]] if page in page_images else []
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
        fragments = [f for f in page.get("fragments") or [] if not str(f.get("coordinateAccuracy", "")).startswith("approximate")]
        if not fragments:
            continue
        measured = [f for f in fragments if f.get("coordinateAccuracy") == "pdf_text_bbox"]
        for fragment in measured:
            x0, y0, x1, y1 = fragment["boundingBox"]
            if x0 < -2 or y0 < -2 or x1 > width + 2 or y1 > height + 2:
                findings.append({"type": "text_outside_page_bounds", "page": page_number, "text": fragment["text"][:80], "boundingBox": fragment["boundingBox"], "confidence": 0.95, "classification": "geometry", "severity": "warning"})
                break
        # Bounding boxes can flag overlap, but mathematical stacking and intentional
        # overprinting still require inspection. Never auto-repair this hint.
        overlap = None
        for i, first in enumerate(measured):
            if len(first["text"]) < 4:
                continue
            a = first["boundingBox"]
            for second in measured[i + 1:]:
                b = second["boundingBox"]
                area = max(0, min(a[2], b[2]) - max(a[0], b[0])) * max(0, min(a[3], b[3]) - max(a[1], b[1]))
                smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
                if smaller > 0 and area / smaller > 0.7 and len(second["text"]) >= 4:
                    overlap = first
                    break
            if overlap:
                break
        if overlap:
            findings.append({"type": "possible_text_overlap", "page": page_number, "text": overlap["text"][:80], "boundingBox": overlap["boundingBox"], "confidence": 0.75, "classification": "advisory", "severity": "warning"})
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
        body_fragments = [f for f in fragments if len(str(f.get("text") or "")) > 1 and float(f.get("y") or 0) >= bottom_margin * 0.7]
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
    from formatter_core.page_layout import running_lines, content_lines
    repeated = running_lines(page_texts)
    for index, text in enumerate(page_texts):
        lines = content_lines(text or "", repeated)
        if not lines:
            continue
        # Heading stranded at the bottom of a page: last non-empty line is a
        # heading and the next page continues with body text.
        if heading_re.match(lines[-1]) and index + 1 < page_count:
            next_lines = content_lines(page_texts[index + 1] or "", repeated)
            if next_lines and not heading_re.match(next_lines[0]):
                findings.append({
                    "type": "heading_orphan_at_page_bottom",
                    "page": index + 1,
                    "headingText": lines[-1][:60],
                    "severity": "warning",
                })
        # Text order alone cannot establish whether a caption belongs to an
        # image on this page. Image/text geometry is checked separately.
        if index > 0:
            previous = content_lines(page_texts[index - 1] or "", repeated)
            toc_lines = [line for line in lines if _looks_like_toc_entry(line)]
            if 0 < len(toc_lines) <= 3 and len(lines) <= 4 and sum(_looks_like_toc_entry(line) for line in previous) >= 5:
                findings.append({"type": "sparse_toc_continuation", "page": index + 1, "entryCount": len(toc_lines), "severity": "warning"})
        for line in lines:
            if re.match(r"^(?:Abstract|Key\s*words?)[:：]\S", line, re.I):
                findings.append({"type": "english_front_label_spacing", "page": index + 1, "text": line[:100], "severity": "warning"})
            elif re.match(r"^(?:Abstract|Key\s*words?)\s+[^:：]", line, re.I):
                findings.append({"type": "english_front_label_missing_colon", "page": index + 1, "text": line[:100], "severity": "warning"})
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
    backend: str = "auto",
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
        "rendered": False,
        "previewGenerated": False,
        "visualReviewStatus": "render_unavailable",
        "unresolvedVisualIssues": [],
    }

    from formatter_core.workspace import file_digest
    candidate_hash = file_digest(source)
    # Render artifacts live under their candidate hash, so final edits cannot reuse stale pages.
    candidate_dir = render_dir / candidate_hash
    candidate_dir.mkdir(parents=True, exist_ok=True)
    import hashlib
    import render_pdf_backends
    converter_hash = hashlib.sha256(Path(render_pdf_backends.__file__).read_bytes()).hexdigest()[:16]
    cache_path = candidate_dir / f"conversion-{backend}-{converter_hash}.json"
    conversion = None
    if cache_path.is_file():
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
        if cached.get("pdfPath") and Path(cached["pdfPath"]).is_file():
            conversion = cached
    if conversion is None:
        conversion = convert_docx_to_pdf(source, candidate_dir, timeout_sec=timeout_sec, backend=backend)
        if conversion.get("available"):
            cache_path.write_text(_json_text(conversion), encoding="utf-8")
    report["candidateSha256"] = candidate_hash
    report["available"] = bool(conversion.get("available"))
    report["backend"] = conversion.get("backend")
    report["backendPath"] = conversion.get("backendPath")
    report["conversionAttempts"] = conversion.get("attempts")
    if not conversion.get("available"):
        report["reason"] = conversion.get("reason")
        return report

    pdf_path = Path(conversion["pdfPath"])
    report["pdfPath"] = str(pdf_path)
    result = analyze_pdf_document(pdf_path, candidate_dir, report=report, timeout_sec=timeout_sec, source_page_hint=source_page_hint, analyze_layout=analyze_layout, export_previews=export_previews, source_docx=source)
    from formatter_core.visual import anchor_findings
    boxes_path = result.get("textGeometryPath")
    text_boxes = json.loads(Path(boxes_path).read_text(encoding="utf-8")) if boxes_path and Path(boxes_path).is_file() else []
    ir_path = source.parent / "thesis_ir.json"
    ir = json.loads(ir_path.read_text(encoding="utf-8")) if ir_path.is_file() else None
    result["visualFindings"] = anchor_findings(source, result.get("visualFindings") or [], text_boxes, ir)
    result["rendered"] = result.get("status") == "ok" and bool(result.get("pageCount"))
    result["previewGenerated"] = any(item.get("status") == "exported" for item in result.get("criticalPages") or [])
    result["visualReviewStatus"] = "render_unavailable" if not result["rendered"] else "pending_agent_review" if result["previewGenerated"] else "pending_manual_review"
    result["unresolvedVisualIssues"] = result["visualFindings"]
    if result.get("visualReviewManifestPath"):
        review_path = Path(result["visualReviewManifestPath"])
        review = json.loads(review_path.read_text(encoding="utf-8"))
        review.update({key: result[key] for key in ("candidateSha256", "visualFindings", "rendered", "previewGenerated", "visualReviewStatus", "unresolvedVisualIssues")})
        review_path.write_text(_json_text(review), encoding="utf-8")
    return result


def analyze_pdf_document(
    pdf_path: Path,
    render_dir: Path,
    *,
    report: Optional[Dict[str, Any]] = None,
    timeout_sec: int = 180,
    source_page_hint: Optional[int] = None,
    analyze_layout: bool = True,
    export_previews: bool = False,
    source_docx: Optional[Path] = None,
) -> Dict[str, Any]:
    render_dir.mkdir(parents=True, exist_ok=True)
    from formatter_core.workspace import file_digest
    cache_key = f"{file_digest(pdf_path) if pdf_path.is_file() else 'missing'}-{file_digest(Path(__file__))}-{file_digest(Path(__file__).parent / 'formatter_core' / 'page_layout.py')}-{int(analyze_layout)}-{int(export_previews)}-{source_page_hint}-{detect_pdftoppm()}"
    import hashlib
    cache_path = render_dir / ("preview-" + hashlib.sha256(cache_key.encode()).hexdigest() + ".json")
    if cache_path.is_file():
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
        artifacts = [*cached.get("contactSheets", []), *[p.get("imagePath") for p in cached.get("criticalPages", [])]]
        if all(path and Path(path).is_file() for path in artifacts):
            return {**(report or {}), **cached}
    if report is None:
        report = {"available": True, "pdfPath": str(pdf_path), "contactSheets": [], "criticalPages": [], "visualFindings": [], "layoutAnalyzed": False, "previewsRequested": export_previews, "visualReviewed": False}
    page_count = count_pdf_pages(pdf_path) or 0
    report["pageCount"] = page_count
    if page_count < 1:
        report["status"] = "failed"
        report["reason"] = "PDF contains no readable pages; visual acceptance is incomplete"
        return report

    if not analyze_layout and not export_previews:
        report["status"] = "ok"
        return report

    text_boxes = extract_page_text_boxes(pdf_path)
    page_texts = ["\n".join(f["text"] for f in page.get("fragments", [])) for page in text_boxes] if len(text_boxes) == page_count else extract_page_texts(pdf_path)
    boxes_path = render_dir / "text_geometry.json"
    boxes_path.write_text(_json_text(text_boxes), encoding="utf-8")
    report["textGeometryPath"] = str(boxes_path)
    report["textGeometryStatus"] = "measured" if any(f.get("coordinateAccuracy") == "pdf_text_bbox" for p in text_boxes for f in p.get("fragments", [])) else "approximate" if text_boxes else "unavailable"
    semantic = locate_semantic_pages(page_texts)
    semantic["pageSummaries"] = build_page_summaries(page_texts)
    semantic_path = render_dir / "semantic_pages.json"
    semantic_path.write_text(_json_text(semantic), encoding="utf-8")
    report["semanticPagesPath"] = str(semantic_path)
    report["suggestedReviewPages"] = suggested_review_pages(semantic)

    thumbnails = export_all_page_thumbnails(pdf_path, render_dir / "all_pages", page_count, timeout_sec) if export_previews else []
    sheets = build_contact_sheets(thumbnails, render_dir) if thumbnails else []
    report["contactSheets"] = [str(p) for p in sheets]

    findings = analyze_page_geometry(thumbnails)
    findings.extend(analyze_text_layout(page_texts, semantic, source_page_hint))
    from formatter_core.page_layout import extract_image_geometry, caption_geometry_findings
    try:
        image_pages = extract_image_geometry(pdf_path)
    except Exception as exc:
        image_pages = []
        report["imageGeometryWarning"] = str(exc)
    image_path = render_dir / "image_geometry.json"
    image_path.write_text(_json_text(image_pages), encoding="utf-8")
    report["imageGeometryPath"] = str(image_path)
    findings.extend(caption_geometry_findings(text_boxes, image_pages))
    if source_docx:
        from formatter_core.visual import structural_layout_findings
        findings.extend(structural_layout_findings(source_docx, page_texts))
    geometry_findings = analyze_text_geometry(text_boxes)
    existing_orphan_pages = {f.get("page") for f in findings if str(f.get("type", "")).startswith("heading_orphan")}
    for finding in geometry_findings:
        if finding.get("type") == "heading_orphan_at_page_bottom_geometry" and finding.get("page") in existing_orphan_pages:
            continue
        findings.append(finding)
    report["visualFindings"] = findings
    report["layoutAnalyzed"] = True
    for finding_index, finding in enumerate(findings):
        if isinstance(finding.get("page"), int):
            semantic.setdefault("regions", {})[f"anomaly_{finding_index}"] = {"page": finding["page"], "label": finding.get("type")}
    # Select chapter starts and representative figure/table pages using text anchors.
    chapter = re.compile(r"^(第.+章|[1-9]\d*\s+\S)")
    for index, text in enumerate(page_texts):
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if any(chapter.match(line) and not _looks_like_toc_entry(line) for line in lines[:6]):
            semantic["regions"][f"chapter_{index + 1}"] = {"page": index + 1, "label": "章节首页"}
        if any(re.match(r"^(?:[图表]\s*\d|Fig(?:ure)?\.?\s*\d|Table\s*\d)", line, re.I) for line in lines) or index < len(image_pages) and image_pages[index].get("images"):
            semantic["regions"][f"asset_{index + 1}"] = {"page": index + 1, "label": "图表页"}
    reference_start = (semantic["regions"].get("references") or {}).get("page")
    if reference_start:
        following = [(semantic["regions"].get(key) or {}).get("page") for key in ("acknowledgements", "appendix")]
        reference_end = min((page for page in following if page and page > reference_start), default=page_count + 1)
        for page in range(reference_start, reference_end):
            semantic["regions"][f"reference_page_{page}"] = {"page": page, "label": "参考文献页"}
    semantic_path.write_text(_json_text(semantic), encoding="utf-8")
    critical_entries = export_critical_pages(pdf_path, render_dir / "critical_pages", semantic, page_count, timeout_sec) if export_previews else []
    report["criticalPages"] = critical_entries
    report["previewDir"] = str(render_dir / "critical_pages")

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
            "异常是程序提示，不代表已确认错误。AI 必须实际查看总览和关键/异常页，并把审阅记录绑定到 candidateSha256；未审阅只能交付待复核候选成品。",
        ],
    }
    manifest_path = render_dir / "visual_review_manifest.json"
    manifest_path.write_text(_json_text(manifest), encoding="utf-8")
    report["visualReviewManifestPath"] = str(manifest_path)

    report["status"] = "ok"
    report["reason"] = None
    # Legacy compatibility with prior previewExport consumers.
    report["previewExport"] = {
        "available": bool(critical_entries) and all(item.get("status") == "exported" for item in critical_entries),
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
    cache_path.write_text(_json_text(report), encoding="utf-8")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a DOCX to PDF, locate key pages semantically, and export review images.")
    parser.add_argument("docx", help="DOCX file to render")
    parser.add_argument("--output-dir", required=True, help="Directory for PDF and render artifacts")
    parser.add_argument("--output", "-o", help="Write JSON report to this path")
    parser.add_argument("--check", choices=["pdf", "layout", "visual"], default="layout", help="Page images are generated only with visual")
    parser.add_argument("--timeout-sec", type=int, default=180, help="Renderer timeout in seconds")
    parser.add_argument("--render-backend", choices=["auto", "word", "wps", "libreoffice"], default="auto")
    parser.add_argument("--source-page-hint", type=int, help="Approximate page count of the source document for explosion detection")
    args = parser.parse_args()

    report = run_render_validation(args.docx, args.output_dir, timeout_sec=args.timeout_sec, source_page_hint=args.source_page_hint, analyze_layout=args.check != "pdf", export_previews=args.check == "visual", backend=args.render_backend)
    text = _json_text(report)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
