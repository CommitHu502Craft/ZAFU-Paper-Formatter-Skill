from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

from lxml import etree as LET

try:
    import latex2mathml.converter as latex2mathml_converter
except ImportError:
    latex2mathml_converter = None


INLINE_MATH_RE = re.compile(r"(\\\(.+?\\\)|(?<!\\)\$(.+?)(?<!\\)\$)")
EQUATION_NUMBER_RE = re.compile(r"(?P<number>[（(]\d+(?:[-.．]\d+)*[）)])$")
LATEX_BLOCK_EQ_RE = re.compile(
    r"^\\begin\{(?P<env>equation\*?|align\*?|gather\*?|multline\*?)\}(?P<body>.+)\\end\{(?P=env)\}$",
    re.S,
)
LATEX_BLOCK_BEGIN_RE = re.compile(r"^\\begin\{(?P<env>equation\*?|align\*?|gather\*?|multline\*?)\}\s*$")
LATEX_BLOCK_END_RE = re.compile(r"^\\end\{(?P<env>equation\*?|align\*?|gather\*?|multline\*?)\}\s*(?P<suffix>.*)$")
TEXTUAL_MATH_LIKE_RE = re.compile(
    r"^[\s\dA-Za-zα-ωΑ-ΩπΔδΣσμνρτφχψΩ×÷+\-*/=<>≤≥≈∑∏∫√^_(),.;:，。；：（）〔〕\[\]{}|%&·•~—\-]+$"
)
STATISTICAL_EXPR_RE = re.compile(r"^\s*[+\-]?\d+(?:\.\d+)?(?:\*{1,3})?\s*\(\s*[+\-]?\d+(?:\.\d+)?\s*\)\s*$")
INLINE_EQUATION_HINT_RE = re.compile(r"(?:=|≈|≤|≥|<|>|\\frac|\\sqrt|\\sum|\\int|\^|_|\\pi|\\alpha|\\beta|\\gamma|\\lambda)")

MML2OMML_XSL_CANDIDATES = [
    Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
]

_MML2OMML_TRANSFORM: Optional[LET.XSLT] = None


def find_mml2omml_xsl() -> Optional[Path]:
    for candidate in MML2OMML_XSL_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def get_mml2omml_transform() -> Optional[LET.XSLT]:
    global _MML2OMML_TRANSFORM
    if _MML2OMML_TRANSFORM is not None:
        return _MML2OMML_TRANSFORM
    xsl_path = find_mml2omml_xsl()
    if xsl_path is None:
        return None
    _MML2OMML_TRANSFORM = LET.XSLT(LET.parse(str(xsl_path)))
    return _MML2OMML_TRANSFORM


def strip_latex_block_environment(text: str) -> Optional[str]:
    stripped = (text or "").strip()
    block_match = LATEX_BLOCK_EQ_RE.match(stripped)
    if block_match is None:
        return None
    return block_match.group("body").strip()


def is_single_line_display_math(text: str) -> Optional[str]:
    stripped = (text or "").strip()
    if stripped.startswith("$$") and stripped.endswith("$$") and len(stripped) >= 4:
        return stripped[2:-2].strip()
    if stripped.startswith(r"\[") and stripped.endswith(r"\]") and len(stripped) >= 4:
        return stripped[2:-2].strip()
    return None


def split_display_math_and_number(text: str) -> Tuple[Optional[str], Optional[str]]:
    stripped = (text or "").strip()
    number_match = EQUATION_NUMBER_RE.search(stripped)
    number = number_match.group("number") if number_match else None
    candidate = stripped[: number_match.start()].rstrip() if number_match else stripped
    block_body = strip_latex_block_environment(candidate)
    if block_body is not None:
        return block_body, number
    math_body = is_single_line_display_math(candidate)
    if math_body is not None:
        return math_body, number
    return None, None


def extract_inline_math_latex(token: str) -> str:
    return token[2:-2] if token.startswith(r"\(") else token[1:-1]


def extract_standalone_formula_latex(text: str) -> Optional[str]:
    stripped = re.sub(r"\s+", " ", (text or "").strip())
    if not stripped or STATISTICAL_EXPR_RE.match(stripped):
        return None
    display_latex, _ = split_display_math_and_number(stripped)
    if display_latex is not None:
        return display_latex
    block_body = strip_latex_block_environment(stripped)
    if block_body is not None:
        return block_body
    inline_match = INLINE_MATH_RE.fullmatch(stripped)
    if inline_match is not None:
        return extract_inline_math_latex(inline_match.group(0)).strip()
    if INLINE_EQUATION_HINT_RE.search(stripped) and TEXTUAL_MATH_LIKE_RE.match(stripped):
        if stripped.count("=") > 2 or len(stripped) > 100:
            return None
        if len(stripped) <= 120:
            return stripped
    return None


def paragraph_looks_like_formula(text: str) -> bool:
    stripped = re.sub(r"\s+", " ", (text or "").strip())
    if not stripped or STATISTICAL_EXPR_RE.match(stripped):
        return False
    if INLINE_MATH_RE.search(stripped):
        return True
    return extract_standalone_formula_latex(stripped) is not None


def normalize_latex_for_conversion(latex: str) -> str:
    block_body = strip_latex_block_environment(latex)
    if block_body is not None:
        return block_body
    return (latex or "").strip()


def detect_latex_block_begin(text: str) -> Optional[str]:
    match = LATEX_BLOCK_BEGIN_RE.match((text or "").strip())
    if match is None:
        return None
    return match.group("env")


def detect_latex_block_end(text: str, expected_env: str) -> Optional[str]:
    match = LATEX_BLOCK_END_RE.match((text or "").strip())
    if match is None or match.group("env") != expected_env:
        return None
    return match.group("suffix").strip()


def latex_to_omml_xml(latex: str, display: str = "inline") -> Optional[bytes]:
    if latex2mathml_converter is None:
        return None
    transform = get_mml2omml_transform()
    if transform is None:
        return None
    normalized = normalize_latex_for_conversion(latex)
    if not normalized:
        return None
    try:
        mathml = latex2mathml_converter.convert(normalized, display=display)
        mathml_root = LET.fromstring(mathml.encode("utf-8"))
        omml_tree = transform(mathml_root)
        return LET.tostring(omml_tree.getroot(), encoding="utf-8")
    except Exception:
        return None
