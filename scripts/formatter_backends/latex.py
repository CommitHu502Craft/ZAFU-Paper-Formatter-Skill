from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from typing import Any

import yaml
from PIL import Image
from pypdf import PdfReader

from formatter_core.checks import CheckPolicy
from formatter_core.workspace import Workspace, file_digest
from formatter_backends.external_class import install_external_class, install_cover_assets
from formatter_backends.latex_frontmatter import load_metadata, prepare_frontmatter
from formatter_backends.latex_bibliography import CITATION, prepare_bibliography, render_citations, install_bibliography_style


ROOT = Path(__file__).resolve().parents[2]
TEXT_ESCAPES = {
    "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$",
    "#": r"\#", "_": r"\_", "{": r"\{", "}": r"\}",
    "~": r"\textasciitilde{}", "^": r"\textasciicircum{}",
}
MATH_COMMANDS = set("""
frac dfrac tfrac sqrt left right big Big bigg Bigg bigl bigr Bigl Bigr
sum prod int iint iiint oint lim log ln exp sin cos tan cot sec csc arcsin arccos arctan
min max sup inf det gcd mod bmod pmod operatorname limits nolimits
alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda
mu nu xi pi varpi rho varrho sigma varsigma tau upsilon phi varphi chi psi omega
Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega
infty partial nabla ell hbar imath jmath Re Im
cdot times div pm mp ast star circ bullet cap cup land lor neg
le leq ge geq neq ne approx sim simeq equiv propto ll gg
in notin subset supset subseteq supseteq forall exists emptyset
to rightarrow leftarrow leftrightarrow Rightarrow Leftarrow Leftrightarrow mapsto
uparrow downarrow langle rangle lvert rvert vert Vert lVert rVert
ldots cdots vdots ddots dots quad qquad , ; : ! space
mathrm mathbf mathit mathsf mathtt mathcal mathbb boldsymbol text textrm textbf
overline underline vec hat bar tilde widehat widetilde dot ddot overset underset
underbrace overbrace binom dbinom tbinom begin end cases aligned matrix pmatrix bmatrix
""".split())
MATH_ENVIRONMENTS = {"aligned", "gathered", "cases", "matrix", "pmatrix", "bmatrix", "vmatrix", "Vmatrix", "smallmatrix"}
INLINE_TOKEN = re.compile(r"\$\$[\s\S]*?\$\$|(?<!\\)\$[^$\n]+?\$|\\\([\s\S]*?\\\)|\\\[[\s\S]*?\\\]|\*\*[^*]+\*\*|(?<!\*)\*[^*\n]+\*|`[^`\n]+`|https?://[^\s{}<>\\（）《》]+")


def escape_text(text: str) -> str:
    return "".join(TEXT_ESCAPES.get(character, character) for character in text)


def safe_math(expression: str) -> bool:
    if any(character in expression for character in ("%", "#", "\x00")):
        return False
    commands = re.findall(r"\\([A-Za-z]+|.)", expression)
    if any(command not in MATH_COMMANDS and command not in {"\\", "{", "}", "|", "_", " ", ",", ";", ":", "!"} for command in commands):
        return False
    environments = re.findall(r"\\(?:begin|end)\s*\{([^}]+)\}", expression)
    if any(environment not in MATH_ENVIRONMENTS for environment in environments):
        return False
    depth = 0
    for character in re.sub(r"\\[{}]", "", expression):
        if character == "{":
            depth += 1
        elif character == "}":
            depth -= 1
        if depth < 0:
            return False
    return depth == 0


def render_inline(text: str, issues: list[dict[str, Any]], block_id: str, *, code_font: str = "monospace", break_code: bool = False) -> str:
    def prose(part: str) -> str:
        if not break_code:
            return escape_text(part)
        # Bare accessions/contrast names occur outside code markup too.
        pieces = []
        cursor = 0
        for token_match in re.finditer(r"[A-Za-z][A-Za-z0-9_./=\-]{11,}|(?:FD|MD)\d+_vs_(?:FD|MD)\d+|[A-Za-z][A-Za-z0-9]*(?:/[A-Za-z0-9]+)+", part):
            pieces.append(escape_text(part[cursor:token_match.start()]))
            token = token_match.group()
            if re.search(r"[0-9_./=\-]|[a-z][A-Z]", token):
                pieces.append("".join(escape_text(char) + (r"\allowbreak{}" if char in "_./=-" or len(token) > 12 and (i + 1) % 6 == 0 else "") for i, char in enumerate(token)))
            else:
                # Ordinary English words keep TeX's linguistic hyphenation.
                pieces.append(escape_text(token))
            cursor = token_match.end()
        pieces.append(escape_text(part[cursor:]))
        return "".join(pieces)
    visible_text = re.sub(r"`[^`\n]*`", "", text)
    if re.search(r"!\[[^]]*\]\([^)]*\)", visible_text):
        issues.append({"kind": "inline_image_requires_own_block", "blockId": block_id, "message": "Place image markup on its own line so an asset block can be extracted; inline markup is preserved as text."})
    chunks: list[str] = []
    cursor = 0
    for match in INLINE_TOKEN.finditer(text):
        chunks.append(prose(text[cursor:match.start()]))
        token = match.group()
        if token.startswith(("http://", "https://")):
            url = token.rstrip(".,;，。；")
            chunks.append(r"\url{" + url + "}" + escape_text(token[len(url):]))
        elif token.startswith(("$", r"\(", r"\[")):
            display = token.startswith(("$$", r"\["))
            expression = token[2:-2] if token.startswith(("$$", r"\(", r"\[")) else token[1:-1]
            if safe_math(expression):
                chunks.append((r"\[" if display else r"\(") + expression + (r"\]" if display else r"\)"))
            else:
                issues.append({"kind": "unsafe_or_unsupported_math", "blockId": block_id, "message": "Formula preserved as escaped text; unsupported TeX commands are never executed."})
                chunks.append(escape_text(token))
        elif token.startswith("**"):
            chunks.append(r"\textbf{" + render_inline(token[2:-2], issues, block_id, code_font=code_font, break_code=break_code) + "}")
        elif token.startswith("*"):
            chunks.append(r"\emph{" + render_inline(token[1:-1], issues, block_id, code_font=code_font, break_code=break_code) + "}")
        else:
            code = token[1:-1]
            allow_code_break = break_code and not re.fullmatch(r"log\d*\([^()\n]+\)", code)
            def code_boundary(i: int, char: str) -> bool:
                if char == "-":
                    # Protect every flag prefix, including subsequent flags.
                    return i > 0 and code[i - 1] != "-" and (i + 1 == len(code) or code[i + 1] != "-")
                return char in "_=,/()." or len(code) > 18 and (i + 1) % 6 == 0
            def code_character(char: str) -> str:
                # Roman text normally ligates -- and quotes into typography.
                # Literal command flags must retain their source characters.
                if code_font == "roman" and char in {'-', '"'}:
                    return r"{\char" + str(ord(char)) + "}"
                return escape_text(char)
            encoded = "".join(code_character(char) + (r"\allowbreak{}" if allow_code_break and code_boundary(i, char) else "") for i, char in enumerate(code))
            chunks.append((r"\textrm{" if code_font == "roman" else r"\texttt{") + encoded + "}")
        cursor = match.end()
    chunks.append(prose(text[cursor:]))
    return "".join(chunks)


def load_config(profile_dir: Path) -> dict[str, Any]:
    config_path = profile_dir / "latex.yaml"
    if config_path.exists():
        return yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    return {"template": "assets/latex/main.tex", "engine": "xelatex", "compliance": "draft", "limitations": ["No school-specific LaTeX adapter is configured; generic draft formatting only."]}


def font_setting(command: str, name: str, fallback: str) -> str:
    if any(character in name + fallback for character in "\\{}%#\n\r"):
        raise ValueError("Unsafe font name in LaTeX profile")
    return rf"\IfFontExistsTF{{{name}}}{{\{command}{{{name}}}}}{{\{command}{{{fallback}}}\PackageWarning{{thesis-formatter}}{{Font fallback used for {command}}}}}"


def numeric_rule(value: Any, default: float) -> str:
    number = float(value if value is not None else default)
    if not 0 <= number <= 100:
        raise ValueError(f"Invalid numeric LaTeX formatting rule: {value}")
    return f"{number:g}"


def build_settings(rules: dict[str, Any], config: dict[str, Any]) -> str:
    page = rules.get("page") or {}
    styles = rules.get("styles") or {}
    body = styles.get("body_text") or {}
    fonts = config.get("fonts") or {}
    settings = [
        font_setting("setmainfont", str(fonts.get("latin") or "Times New Roman"), str(fonts.get("fallback_latin") or "TeX Gyre Termes")),
        font_setting("setCJKmainfont", str(fonts.get("body") or "宋体"), str(fonts.get("fallback_body") or "FandolSong-Regular")),
        font_setting("setCJKsansfont", str(fonts.get("heading") or "楷体"), str(fonts.get("fallback_heading") or "FandolKai-Regular")),
        r"\geometry{a4paper," + ",".join(f"{side}={numeric_rule(page.get(key), 2.7)}cm" for side, key in (("top", "margin_top_cm"), ("bottom", "margin_bottom_cm"), ("left", "margin_left_cm"), ("right", "margin_right_cm"))) + "}",
        rf"\AtBeginDocument{{\fontsize{{{numeric_rule(body.get('size_pt'), 10.5)}}}{{{numeric_rule(body.get('line_spacing_pt'), 20)}}}\selectfont}}",
        rf"\setlength{{\parindent}}{{{numeric_rule(body.get('indent_chars'), 2)}em}}",
        r"\setlength{\parskip}{0pt}",
    ]
    for level, section in ((1, "section"), (2, "subsection"), (3, "subsubsection")):
        style = styles.get(f"heading{level}") or {}
        size = numeric_rule(style.get("size_pt"), 14 if level == 1 else 12)
        leading = numeric_rule(style.get("line_spacing_pt"), 20)
        before = numeric_rule(style.get("spacing_before_pt"), 12)
        after = numeric_rule(style.get("spacing_after_pt"), 6)
        weight = r"\bfseries" if style.get("bold", True) else r"\mdseries"
        settings.append(rf"\ctexset{{{section}={{format=\sffamily{weight}\fontsize{{{size}}}{{{leading}}}\selectfont,beforeskip={before}pt,afterskip={after}pt}}}}")
    return "\n".join(settings)


def render_table(block: dict[str, Any], issues: list[dict[str, Any]], renderer=None, *, config=None, caption="") -> str:
    lines = (block.get("attributes") or {}).get("lines") or []
    rows = [re.split(r"(?<!\\)\|", line.strip().strip("|")) for index, line in enumerate(lines) if index != 1]
    if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
        issues.append({"kind": "unsupported_table", "blockId": block["id"], "message": "Malformed table retained as escaped text."})
        return "\n\n".join(escape_text(str(line)) for line in lines)
    columns = len(rows[0])
    config = config or {}
    render = renderer or (lambda text, identifier: render_inline(text, issues, identifier))
    identifier_pattern = re.compile(r"(?<![A-Za-z0-9])(?:LOC\d+|[A-Z]{1,4}(?:_|\\_)\d+(?:\.\d+)?|(?:FD|MD)\d+(?:_|\\_)vs(?:_|\\_)(?:FD|MD)\d+)(?![A-Za-z0-9])")
    def render_cell(cell: str) -> str:
        cell = cell.strip().replace(r"\|", "|")
        rendered = render(cell, block["id"])
        # Render Markdown first, so matching an ID inside backticks/emphasis
        # never splits its formatting delimiters into literal source markers.
        if r"\url{" in rendered:
            return rendered
        return identifier_pattern.sub(lambda match: r"\allowbreak{}".join(
            escape_text(character) for character in match.group().replace(r"\_", "_")), rendered)
    rendered_rows = [" & ".join(render_cell(cell) for cell in row) + r" \\" for row in rows]
    spacing = r"\setlength{\tabcolsep}{3pt}" if columns >= 7 else ""
    alignment = r"\centering" if config.get("table_alignment") == "center" else r"\raggedright"
    # Use modest content-based widths, not equal-width columns or font shrinking.
    lengths = [max(len(re.sub(r"[`*]", "", row[i].strip())) for row in rows) for i in range(columns)]
    weights = [max(1.0, min(2.4, length / 10)) for length in lengths]
    for i in range(columns):
        # Dense unsplit letter+digit accessions need more width than a
        # similarly long descriptive field with natural word boundaries.
        if any(re.fullmatch(r"[A-Za-z]{2,6}\d{6,}", cell.strip().strip("`*")) for cell in (row[i] for row in rows[1:])):
            weights[i] = max(weights[i], min(2.4, lengths[i] / 8))
    total = sum(weights)
    spec = "".join(rf">{{\hsize={weight * columns / total:.4f}\hsize\linewidth=\hsize{alignment}\arraybackslash}}X" for weight in weights)
    wide = columns >= int(config.get("wide_table_columns", 100))
    start = [r"\clearpage\begin{landscape}"] if wide else []
    finish = [r"\end{landscape}\clearpage"] if wide else []
    return "\n".join([*start, r"\begin{table}[H]\centering", r"\begingroup", r"\renewcommand{\tabularxcolumn}[1]{m{#1}}", r"\renewcommand{\arraystretch}{1.25}", spacing, r"\caption*{" + caption + "}" if caption else "", rf"\begin{{tabularx}}{{\linewidth}}{{@{{}}{spec}@{{}}}}", r"\toprule", rendered_rows[0], r"\midrule", *rendered_rows[1:], r"\bottomrule", r"\end{tabularx}", r"\endgroup\end{table}", *finish])


def render_image(block: dict[str, Any], source: Path, project: Path, issues: list[dict[str, Any]], assets: list[dict[str, Any]], renderer=None, *, config=None, caption="") -> str:
    attributes = block.get("attributes") or {}
    raw_path = str(attributes.get("path") or "")
    image_path = (source.parent / raw_path).resolve()
    allowed = bool(raw_path) and not Path(raw_path).is_absolute() and image_path.is_relative_to(source.parent.resolve())
    extension = image_path.suffix.lower()
    if not allowed or not image_path.is_file() or extension not in {".png", ".jpg", ".jpeg", ".pdf"}:
        issues.append({"kind": "unsupported_or_missing_image", "blockId": block["id"], "path": raw_path, "message": "Only existing PNG/JPEG/PDF assets within the input directory are supported."})
        return r"\noindent " + escape_text(f"[图片未转换：{attributes.get('alt') or block.get('text') or raw_path}]")
    try:
        if extension == ".pdf":
            if not PdfReader(image_path).pages:
                raise ValueError("Empty PDF image")
        else:
            with Image.open(image_path) as image:
                image.verify()
    except Exception as error:
        issues.append({"kind": "invalid_image", "blockId": block["id"], "message": str(error)})
        return escape_text(f"[无效图片：{raw_path}]")
    digest = file_digest(image_path)
    relative = f"figures/{digest[:16]}{extension}"
    destination = project / relative
    destination.parent.mkdir(exist_ok=True)
    if not destination.exists():
        shutil.copyfile(image_path, destination)
    assets.append({"blockId": block["id"], "path": relative, "sha256": digest})
    render = renderer or (lambda text, identifier: render_inline(text, issues, identifier))
    alt = render(str(attributes.get("alt") or block.get("text") or ""), block["id"])
    config = config or {}
    width = float(config.get("image_width_fraction", 0.85))
    height = float(config.get("image_height_fraction", 0.65))
    if not 0 < width <= 1 or not 0 < height <= 0.9:
        raise ValueError("Image size fractions must fit the printable page")
    return "\n".join([r"\begin{figure}[H]\centering", rf"\includegraphics[width={width:g}\linewidth,height={height:g}\textheight,keepaspectratio]{{{relative}}}", r"\caption*{" + (caption or alt) + "}" if caption or alt else "", r"\end{figure}"])


def build_project(ir: dict[str, Any], source: Path, rules: dict[str, Any], profile_dir: Path, project: Path, *, metadata_path: Path | None = None, bibliography_path: Path | None = None) -> dict[str, Any]:
    if (ir.get("source") or {}).get("format") not in {"markdown", "text"}:
        raise ValueError("LaTeX MVP only accepts Markdown/TXT ThesisIR")
    project.mkdir(parents=True, exist_ok=True)
    config = load_config(profile_dir)
    issues: list[dict[str, Any]] = []
    warnings = list(config.get("limitations") or [])
    assets: list[dict[str, Any]] = []
    metadata = load_metadata(metadata_path)
    if metadata.get("frontMatterPdf"):
        front_path = (metadata_path.parent / metadata["frontMatterPdf"]).resolve(strict=True)
        if front_path.suffix.lower() != ".pdf" or not PdfReader(front_path).pages:
            raise ValueError("frontMatterPdf must name an existing nonempty PDF")
        shutil.copyfile(front_path, project / "front-matter.pdf")
        assets.append({"kind": "external_frontmatter", "path": "front-matter.pdf", "sha256": file_digest(front_path)})
    bibliography = prepare_bibliography(bibliography_path, project)
    issues.extend(bibliography["errors"])
    warnings.extend(bibliography["warnings"])
    used_citations: set[str] = set()
    def plain_inline(text: str, identifier: str) -> str:
        return render_inline(text, issues, identifier, code_font=str(config.get("inline_code_font", "monospace")), break_code=bool(config.get("break_inline_code", False)))
    def inline(text: str, identifier: str) -> str:
        if bibliography["enabled"]:
            return render_citations(text, set(bibliography["keys"]), used_citations, lambda part: plain_inline(part, identifier), issues, identifier)
        if CITATION.search(re.sub(r"`[^`]*`", "", text)):
            issues.append({"kind": "citation_requires_bibliography", "blockId": identifier, "message": "Supply --latex-bibliography for explicit citation markers."})
        return plain_inline(text, identifier)
    frontmatter = None
    if config.get("external_class") == "zafu":
        frontmatter = prepare_frontmatter(ir, metadata, escape_text, inline)
        issues.extend(frontmatter["errors"])
        warnings.extend(frontmatter["warnings"])
    elif metadata or bibliography["enabled"]:
        raise ValueError("LaTeX metadata and bibliography adapters require the external ZAFU class profile")
    blocks = ir.get("semanticBlocks") or []
    content: list[str] = []
    block_index: list[dict[str, Any]] = []
    title = str((ir.get("frontMatter") or {}).get("title") or "")
    toc_inserted = bool(frontmatter and frontmatter["tocInserted"])
    references_started = False
    bibliography_inserted = False
    literal_reference_section = False
    bibliography_tex = "\n".join([r"\clearpage\phantomsection", r"\addcontentsline{toc}{section}{参考文献}", r"\bibliography{references}"])
    attached_captions = set()
    table_captions = {}
    def adjacent_caption(item, label):
        if item.get("kind") == "caption":
            return True
        visible = str(item.get("text") or "").replace("**", "").lstrip()
        return item.get("kind") == "paragraph" and bool(re.match(rf"^{label}\s*A?\d+(?:\s|[：:、])", visible))
    for index, item in enumerate(blocks):
        if item.get("kind") == "table" and index and adjacent_caption(blocks[index - 1], "表"):
            preceding = blocks[index - 1]
            table_captions[item["id"]] = preceding
            attached_captions.add(preceding["id"])
    for index, block in enumerate(blocks):
        block_id = str(block.get("id") or "unknown")
        if block_id in attached_captions:
            block_index.append({"blockId": block_id, "status": "attached_caption"})
            continue
        if frontmatter and block_id in frontmatter["consumedIds"]:
            block_index.append({"blockId": block_id, "status": "rendered_frontmatter_macro", "sourceAnchor": block.get("sourceAnchor")})
            continue
        kind = str(block.get("kind") or "")
        role = str(block.get("role") or "")
        text = str(block.get("text") or "")
        if role == "template_example":
            block_index.append({"blockId": block_id, "status": "excluded_template_example"})
            continue
        if not toc_inserted and text != title:
            content.extend([r"\tableofcontents", r"\clearpage"])
            toc_inserted = True
        reference_heading = role == "references_heading" or (kind == "heading" and text.strip() == "参考文献")
        if kind == "heading" and not reference_heading:
            literal_reference_section = False
        if reference_heading:
            literal_reference_section = True
        if bibliography["enabled"] and (kind == "reference" or literal_reference_section):
            if not bibliography_inserted:
                content.append(bibliography_tex)
                bibliography_inserted = True
            block_index.append({"blockId": block_id, "status": "replaced_by_explicit_bibliography", "sourceAnchor": block.get("sourceAnchor")})
            continue
        rendered = inline(text, block_id)
        if kind == "image":
            following = blocks[index + 1] if index + 1 < len(blocks) else {}
            caption = ""
            if adjacent_caption(following, "图"):
                caption = inline(str(following.get("text") or ""), str(following["id"]))
                attached_captions.add(following["id"])
            rendered = render_image(block, source, project, issues, assets, inline, config=config, caption=caption)
        elif kind == "table":
            caption_block = table_captions.get(block_id, {})
            rendered = render_table(block, issues, inline, config=config, caption=inline(str(caption_block.get("text") or ""), block_id))
        elif kind == "page_break":
            rendered = r"\clearpage"
        elif kind == "equation":
            expression = re.sub(r"^\$\$|\$\$$", "", text.strip())
            if safe_math(expression):
                rendered = r"\[" + expression + r"\]"
            else:
                issues.append({"kind": "unsafe_or_unsupported_math", "blockId": block_id})
                rendered = escape_text(text)
        elif kind == "heading":
            level = int(block.get("level") or 1)
            heading = re.sub(r"^(?:\d+(?:\.\d+)*[.、]?\s+|第[一二三四五六七八九十百\d]+章\s*)", "", text)
            is_references = role == "references_heading" or heading.strip() == "参考文献"
            references_started = references_started or is_references
            if text == title and role == "front_matter":
                rendered = r"\begin{center}\Large " + rendered + r"\end{center}"
            elif role == "front_matter" or is_references or role in {"acknowledgements_heading", "appendix_heading"}:
                rendered = r"\clearpage\section*{" + rendered + "}" + r"\addcontentsline{toc}{section}{" + rendered + "}"
            elif level > 3:
                issues.append({"kind": "unsupported_heading_level", "blockId": block_id, "level": level})
                rendered = r"\paragraph*{" + rendered + "}"
            else:
                command = {1: "section", 2: "subsection", 3: "subsubsection"}[level]
                rendered = "\\" + command + "*{" + rendered + "}" + rf"\addcontentsline{{toc}}{{{command}}}{{" + rendered + "}"
        elif kind in {"paragraph", "reference", "caption"}:
            if references_started or kind == "reference":
                rendered = r"\noindent\hangindent=2em\hangafter=1 " + rendered + r"\par"
        else:
            issues.append({"kind": "unsupported_block_kind", "blockId": block_id, "value": kind})
        content.append(rendered)
        block_index.append({"blockId": block_id, "status": "rendered", "kind": kind, "sourceAnchor": block.get("sourceAnchor")})
    if bibliography["enabled"] and not bibliography_inserted:
        content.append(bibliography_tex)
    bibliography["usedKeys"] = sorted(used_citations)
    if bibliography["enabled"] and not used_citations:
        warnings.append("No explicit citation markers were used; uncited BibTeX entries are not automatically included.")
    template_path = (ROOT / str(config.get("template") or "assets/latex/main.tex")).resolve()
    if not template_path.is_relative_to(ROOT):
        raise ValueError("LaTeX template must be a repository resource")
    template = template_path.read_text(encoding="utf-8")
    if template.count("@@SETTINGS@@") != 1 or template.count("@@CONTENT@@") != 1:
        raise ValueError("LaTeX template must contain exactly one settings and one content placeholder")
    external_class = None
    if config.get("external_class") == "zafu":
        external_class = install_external_class(project)
        settings = r"\hypersetup{hidelinks}" + "\n" + r"\urlstyle{same}"
        if config.get("emergency_stretch_em") is not None:
            settings += "\n" + rf"\setlength{{\emergencystretch}}{{{numeric_rule(config['emergency_stretch_em'], 0)}em}}"
        if frontmatter and (frontmatter["cover"] or frontmatter["statement"]) and not frontmatter["errors"]:
            external_class["assets"] = install_cover_assets(project)
            (project / "UPSTREAM.json").write_text(json.dumps(external_class, indent=2), encoding="utf-8")
        if bibliography["enabled"]:
            if not bibliography["errors"]:
                bibliography["upstreamStyle"] = install_bibliography_style(project)
            settings += "\n" + "\n".join([
                r"\citestyle{authoryear}", r"\setcitestyle{authoryear,round,semicolon}",
                r"\bibliographystyle{zafu-gbt7714-2015-authoryear}",
                r"\renewcommand{\bibfont}{\songti\zihao{5}}",
            ])
    else:
        settings = build_settings(rules, config)
    replacements = {"@@SETTINGS@@": settings, "@@CONTENT@@": "\n\n".join(content)}
    if frontmatter:
        for placeholder, field in (("@@METADATA@@", "metadataTex"), ("@@FRONTMATTER@@", "frontmatterTex")):
            if template.count(placeholder) != 1:
                raise ValueError(f"External class adapter needs exactly one {placeholder} placeholder")
            replacements[placeholder] = frontmatter[field]
    document = re.sub(r"@@(?:SETTINGS|CONTENT|METADATA|FRONTMATTER)@@", lambda match: replacements[match.group()], template)
    (project / "main.tex").write_text(document, encoding="utf-8")
    (project / "README.txt").write_text(
        "Generated thesis draft. Compile with XeLaTeX (no shell escape), twice or until references stabilize.\n"
        "External CLS front matter is independent of Word. No visual review or automatic signature was performed.\n"
        "Missing prescribed fonts can fail compilation; retained family defaults are reported as warnings.\n"
        "With explicit references.bib: XeLaTeX, then bibtex build/main, then XeLaTeX twice; Tectonic runs BibTeX internally.\n"
        "Without explicit BibTeX input, references remain literal source text.\n", encoding="utf-8",
    )
    report = {
        "backend": "latex", "sourceFormat": ir["source"]["format"], "projectDir": str(project),
        "mainTex": str(project / "main.tex"), "profileCompliance": config.get("compliance", "draft"),
        "templateSha256": file_digest(template_path), "blockCount": len(blocks), "blockIndex": block_index,
        "assets": assets, "errors": issues, "warnings": warnings,
        "externalClass": external_class,
        "frontMatter": {key: value for key, value in (frontmatter or {}).items() if key not in {"metadataTex", "frontmatterTex", "consumedIds", "errors", "warnings"}},
        "metadataSourceSha256": file_digest(metadata_path) if metadata_path else None,
        "bibliography": bibliography,
        "passed": not issues, "visualReviewed": False,
    }
    portable_report = dict(report, projectDir=".", mainTex="main.tex")
    (project / "conversion.json").write_text(json.dumps(portable_report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def compile_project(project: Path, timeout: int = 120, no_compile: bool = False, engine_name: str = "auto") -> dict[str, Any]:
    if no_compile:
        return {"status": "skipped", "reason": "--no-compile requested", "pdfPath": None}
    if engine_name not in {"auto", "xelatex", "tectonic"}:
        raise ValueError(f"Unsupported LaTeX engine: {engine_name}")
    engine = shutil.which("xelatex") if engine_name in {"auto", "xelatex"} else None
    selected_engine = "xelatex"
    if not engine and engine_name in {"auto", "tectonic"}:
        engine = shutil.which("tectonic")
        selected_engine = "tectonic"
    if not engine:
        return {"status": "unavailable", "reason": f"No compatible {engine_name} TeX engine was found on PATH; source project is still available", "pdfPath": None}
    needs_bibtex = (project / "references.bib").is_file()
    bibtex = shutil.which("bibtex") if needs_bibtex and selected_engine == "xelatex" else None
    if needs_bibtex and selected_engine == "xelatex" and not bibtex:
        return {"status": "failed", "reason": "XeLaTeX bibliography compilation requires bibtex on PATH; no PDF will be published", "pdfPath": None}
    build_dir = project / "build"
    build_dir.mkdir(exist_ok=True)
    command = ([engine, "-interaction=nonstopmode", "-halt-on-error", "-no-shell-escape", "-output-directory=build", "main.tex"] if selected_engine == "xelatex" else [engine, "-X", "compile", "--untrusted", "--keep-logs", "--print", "--outdir", "build", "main.tex"])
    transcripts: list[str] = []
    result: dict[str, Any] = {"status": "failed", "engine": engine, "engineName": selected_engine, "command": command, "pdfPath": None}
    for pass_number in range(1, 4):
        try:
            completed = subprocess.run(command, cwd=str(project), capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout, check=False)
        except subprocess.TimeoutExpired:
            result["reason"] = f"{selected_engine} pass {pass_number} timed out after {timeout}s"
            break
        except OSError as error:
            result["reason"] = str(error)
            break
        transcripts.append(completed.stdout + "\n" + completed.stderr)
        result["passes"] = pass_number
        result["returnCode"] = completed.returncode
        if completed.returncode != 0:
            result["reason"] = f"{selected_engine} exited unsuccessfully; no PDF will be published"
            break
        if needs_bibtex and selected_engine == "tectonic" and "errors were issued by BibTeX" in completed.stdout + completed.stderr:
            result["reason"] = "Tectonic reported BibTeX errors despite a successful engine exit; no PDF will be published"
            break
        if bibtex and pass_number == 1:
            environment = os.environ.copy()
            environment["BIBINPUTS"] = str(project.resolve()) + os.pathsep + environment.get("BIBINPUTS", "")
            try:
                bibliography_run = subprocess.run([bibtex, "main"], cwd=str(build_dir), env=environment, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout, check=False)
            except (OSError, subprocess.TimeoutExpired) as error:
                result["reason"] = f"BibTeX compilation failed: {error}"
                break
            transcripts.append(bibliography_run.stdout + "\n" + bibliography_run.stderr)
            result["bibtexReturnCode"] = bibliography_run.returncode
            if bibliography_run.returncode != 0:
                result["reason"] = "BibTeX failed; no PDF will be published"
                break
        rerun_needed = bool(re.search(r"Rerun to get|Label\(s\) may have changed|rerunfilecheck Warning|There were undefined citations|Citation .*undefined", completed.stdout))
        if selected_engine == "tectonic" or (pass_number >= 2 and not rerun_needed):
            final_log = (build_dir / "main.log").read_text(encoding="utf-8", errors="replace") if (build_dir / "main.log").exists() else completed.stdout
            if needs_bibtex and re.search(r"There were undefined citations|Citation .*undefined", final_log):
                result["reason"] = "Unresolved citations remain after compilation; no PDF will be published"
                break
            candidate = build_dir / "main.pdf"
            try:
                page_count = len(PdfReader(candidate).pages)
                if page_count < 1:
                    raise ValueError("Compiled PDF contains no pages")
            except Exception as error:
                result["reason"] = f"Compiled PDF validation failed: {error}"
            else:
                result.update(status="ok", pdfPath=str(candidate), pageCount=page_count)
            break
    else:
        result["reason"] = "References did not stabilize within three passes"
    transcript = "\n\n".join(transcripts)
    log_path = build_dir / "compiler-output.log"
    log_path.write_text(transcript, encoding="utf-8")
    result["logPath"] = str(log_path)
    result["warnings"] = list(dict.fromkeys(line.strip() for line in transcript.splitlines() if "Warning" in line or "Overfull" in line or "Underfull" in line))[:30]
    return result


def run_latex(ir: dict[str, Any], source: Path, rules: dict[str, Any], profile_dir: Path, workspace: Workspace, policy: CheckPolicy, *, no_compile: bool = False, timeout: int = 120, engine_name: str = "auto", metadata_path: Path | None = None, bibliography_path: Path | None = None) -> dict[str, Any]:
    project = workspace.work / "latex"
    report = build_project(ir, source, rules, profile_dir, project, metadata_path=metadata_path, bibliography_path=bibliography_path)
    compilation = compile_project(project, timeout, no_compile, engine_name) if report["passed"] else {"status": "blocked", "reason": "Conversion errors must be resolved before compilation", "pdfPath": None}
    report["compilation"] = compilation
    report["checkPolicy"] = policy.summary()
    deliverables: dict[str, Any] = {}
    source_zip = workspace.work / "latex-source.zip"
    with zipfile.ZipFile(source_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(project.rglob("*")):
            if path.is_file() and not path.is_relative_to(project / "build"):
                archive.write(path, path.relative_to(project).as_posix())
    deliverables["latexSource"] = str(workspace.publish(source_zip, "latex-source", "zip"))
    if compilation.get("status") == "ok":
        pdf = Path(compilation["pdfPath"])
        if policy.level != "structural":
            from render_validate_docx import analyze_pdf_document

            report["layoutValidation"] = analyze_pdf_document(pdf, workspace.work / "latex-preview", export_previews=policy.level == "visual")
        deliverables["latexPdf"] = str(workspace.publish(pdf, "latex", "pdf"))
        relocations = {str(pdf): deliverables["latexPdf"]}
        workspace.rewrite_references(relocations)
        report = workspace.remap(report, relocations)
        compilation["pdfPath"] = deliverables["latexPdf"]
    report["deliverables"] = deliverables
    report_path = workspace.reports / "latex.json"
    report_path.write_text(json.dumps(workspace.portable_paths(report), ensure_ascii=True, indent=2), encoding="utf-8")
    report["reportPath"] = str(report_path)
    return report
