"""Image geometry and caption checks; never infer image position from text order."""
from __future__ import annotations

import re
from collections import Counter


def _multiply(a, b):
    return (
        a[0]*b[0] + a[1]*b[2], a[0]*b[1] + a[1]*b[3],
        a[2]*b[0] + a[3]*b[2], a[2]*b[1] + a[3]*b[3],
        a[4]*b[0] + a[5]*b[2] + b[4], a[4]*b[1] + a[5]*b[3] + b[5],
    )


def extract_image_geometry(pdf_path):
    """Measure image XObjects, including nested forms, in unrotated PDF space.

    Rotated/cropped pages and inline images are left for visual review. Clipping
    can reduce the visible rectangle; these boxes describe painted image bounds.
    """
    from pypdf import PdfReader
    from pypdf.generic import ContentStream
    reader = PdfReader(str(pdf_path))
    pages = []
    identity = (1, 0, 0, 1, 0, 0)

    def walk(stream, resources, matrix, result, depth=0):
        if stream is None or depth > 8:
            return
        stack = []
        for operands, operator in ContentStream(stream, reader).operations:
            if operator == b"q":
                stack.append(matrix)
            elif operator == b"Q":
                matrix = stack.pop() if stack else identity
            elif operator == b"cm":
                matrix = _multiply(tuple(float(v) for v in operands), matrix)
            elif operator == b"Do":
                objects = resources.get("/XObject", {})
                objects = objects.get_object() if hasattr(objects, "get_object") else objects
                obj = objects.get(operands[0])
                if obj is None:
                    continue
                obj = obj.get_object()
                if obj.get("/Subtype") == "/Form":
                    transform = tuple(float(v) for v in obj.get("/Matrix", identity))
                    walk(obj, obj.get("/Resources", resources), _multiply(transform, matrix), result, depth + 1)
                elif obj.get("/Subtype") == "/Image":
                    corners = [(x*matrix[0] + y*matrix[2] + matrix[4], x*matrix[1] + y*matrix[3] + matrix[5]) for x, y in ((0,0), (0,1), (1,0), (1,1))]
                    xs, ys = zip(*corners)
                    box = [min(xs), min(ys), max(xs), max(ys)]
                    if box[2] - box[0] >= 24 and box[3] - box[1] >= 24:
                        result.append({"boundingBox": box, "coordinateAccuracy": "pdf_image_transform"})

    for page in reader.pages:
        result = {"images": [], "status": "measured"}
        if page.rotation or list(page.cropbox) != list(page.mediabox) or float(page.mediabox.left) != 0 or float(page.mediabox.bottom) != 0:
            result["status"] = "needs_visual_review"
        else:
            try:
                walk(page.get_contents(), page.get("/Resources", {}), identity, result["images"])
            except (ValueError, KeyError, TypeError, RecursionError):
                result = {"images": [], "status": "unavailable"}
        pages.append(result)
    return pages


def running_lines(page_texts):
    counts = Counter()
    for text in page_texts:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        counts.update(set(lines[:1] + lines[-1:]))
    return {line for line, count in counts.items() if count >= max(2, len(page_texts) * .4)}


def content_lines(text, repeated):
    return [line.strip() for line in text.splitlines() if line.strip() and line.strip() not in repeated and not re.fullmatch(r"(?:\d+|[ivxlcdm]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)", line.strip(), re.I)]


def caption_geometry_findings(text_pages, image_pages):
    findings = []
    for index, (page, images) in enumerate(zip(text_pages, image_pages), 1):
        for fragment in page.get("fragments", []):
            if not re.match(r"^(?:图\s*\d|Fig(?:ure)?\.?\s*\d)", fragment.get("text", ""), re.I):
                continue
            box = fragment.get("boundingBox")
            if fragment.get("coordinateAccuracy") != "pdf_text_bbox" or not box or any(v is None for v in box):
                continue
            matches = []
            for item in images.get("images", []):
                image = item["boundingBox"]
                overlap_x = min(box[2], image[2]) - max(box[0], image[0])
                gap = max(image[1] - box[3], box[1] - image[3], 0)
                if overlap_x > 0 and gap <= 72:
                    matches.append((gap, image))
            matches.sort(key=lambda item: item[0])
            if not matches or len(matches) > 1 and matches[1][0] - matches[0][0] < 12:
                continue  # ambiguous pairing must not become an automatic repair
            gap, image = matches[0]
            kind = None
            if box[1] >= image[3]:
                kind = "figure_caption_above_image"
            elif box[3] > image[1] and box[1] < image[3]:
                kind = "figure_caption_overlaps_image"
            elif gap > 36:
                kind = "figure_caption_large_gap"
            if kind:
                findings.append({"type": kind, "page": index, "captionText": fragment["text"], "boundingBox": box, "imageBoundingBox": image, "severity": "warning", "mappingStatus": "geometry_pair_needs_visual_confirmation"})
    return findings
