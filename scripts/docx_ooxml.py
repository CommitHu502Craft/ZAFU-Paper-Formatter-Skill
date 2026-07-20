#!/usr/bin/env python3
from __future__ import annotations

import copy
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import yaml

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

FONT_SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")
FONT_THEME_ATTRS = {
    "ascii": ("asciiTheme",),
    "hAnsi": ("hAnsiTheme",),
    "eastAsia": ("eastAsiaTheme",),
    "cs": ("cstheme", "csTheme"),
}
PUNCTUATION_ENDINGS = "。；;，,：:！？!?）)"
IDEOGRAPHIC_SPACE = "\u3000"

SCIENCE_DECIMAL_RE = re.compile(r"^(?P<prefix>\d+(?:\.\d+){0,3})(?P<tail>(?:[.、])?\s*)")
CN_LEVEL1_RE = re.compile(r"^(?P<prefix>[一二三四五六七八九十百千]+、)(?P<tail>\s*)")
CN_LEVEL2_RE = re.compile(r"^(?P<prefix>（[一二三四五六七八九十百千]+）)(?P<tail>\s*)")
CN_NUMERIC_LEVEL_RE = re.compile(r"^(?P<prefix>\d+、)(?P<tail>\s*)")
CN_NUMERIC_PAREN_RE = re.compile(r"^(?P<prefix>（\d+）)(?P<tail>\s*)")
CHAPTER_RE = re.compile(r"^(?P<prefix>第[一二三四五六七八九十百千]+章)(?P<tail>\s*)")
SECTION_RE = re.compile(r"^(?P<prefix>第[一二三四五六七八九十百千]+节)(?P<tail>\s*)")
CAPTION_LABEL_TOKEN_RE = r"(?:\d+|[一二三四五六七八九十百千]+)(?:[a-zA-Z])?(?:[-—.．](?:\d+|[一二三四五六七八九十百千]+))?"
FIGURE_CAPTION_RE = re.compile(rf"^图\s*{CAPTION_LABEL_TOKEN_RE}(?:\s+|[：:．。]|$)")
TABLE_CAPTION_RE = re.compile(rf"^表\s*{CAPTION_LABEL_TOKEN_RE}(?:\s+|[：:．。]|$)")
EQUATION_NUMBER_RE = re.compile(r"[（(]\d+(?:[-.．]\d+)?[）)]$")
MANUAL_TOC_RE = re.compile(r"[.．·•]{4,}\s*\d+$")
CROSS_REF_RE = re.compile(rf"(?P<kind>图|表)\s*(?P<label>{CAPTION_LABEL_TOKEN_RE})")

CN_DIGITS = {
    "零": 0,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


def qn(tag: str) -> str:
    return f"{W}{tag}"


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def read_xml(zf: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    if name not in zf.namelist():
        return None
    return ET.fromstring(zf.read(name))


def safe_int(value: Any) -> Optional[int]:
    if value is None or value == "":
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def safe_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def twips_to_pt(value: Any) -> Optional[float]:
    number = safe_int(value)
    if number is None:
        return None
    return round(number / 20.0, 2)


def twips_to_cm(value: Any) -> Optional[float]:
    number = safe_int(value)
    if number is None:
        return None
    return round(number / 1440.0 * 2.54, 3)


def pt_to_half_points(value: Any) -> Optional[int]:
    number = safe_float(value)
    if number is None:
        return None
    return int(round(number * 2))


def half_points_to_pt(value: Any) -> Optional[float]:
    number = safe_int(value)
    if number is None:
        return None
    return round(number / 2.0, 2)


def onoff_value(node: Optional[ET.Element]) -> Optional[bool]:
    if node is None:
        return None
    val = node.attrib.get(qn("val"))
    if val is None:
        return True
    return val not in {"0", "false", "False", "off"}


def contains_cjk(text: str) -> bool:
    return bool(re.search(r"[\u3400-\u9fff\uF900-\uFAFF]", text or ""))


def paragraph_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS))


def iter_text_runs(paragraph: ET.Element) -> Iterable[ET.Element]:
    for child in paragraph:
        tag = local_name(child.tag)
        if tag == "r":
            yield child
        elif tag == "hyperlink":
            for run in child.findall("w:r", NS):
                yield run
        elif tag in {"smartTag", "sdt"}:
            for run in child.findall(".//w:r", NS):
                yield run


def run_text(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.findall(".//w:t", NS))


def flatten_attrs(node: Optional[ET.Element], keys: Iterable[str]) -> Dict[str, Any]:
    if node is None:
        return {}
    out: Dict[str, Any] = {}
    for key in keys:
        value = node.attrib.get(qn(key))
        if value is not None:
            out[key] = value
    return out


def deep_merge(base: Dict[str, Any], overlay: Dict[str, Any]) -> Dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in overlay.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def load_rules(path: Optional[str]) -> Dict[str, Any]:
    if not path:
        return {}
    return yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}


DEFAULT_PROFILE_STYLE_TARGETS = {
    "title_cn": "zafu_title_cn",
    "title_en": "zafu_title_en",
    "abstract_heading": "zafu_heading1",
    "abstract_text_cn": "zafu_abstract_text_cn",
    "abstract_text_en": "zafu_abstract_text_en",
    "keywords_cn": "zafu_keywords_cn",
    "keywords_en": "zafu_keywords_en",
    "toc_title": "zafu_toc_title",
    "toc_body": "zafu_toc_body",
    "heading1": "zafu_heading1",
    "heading2": "zafu_heading2",
    "heading3": "zafu_heading3",
    "body_text": "zafu_body",
    "caption": "zafu_caption",
}


def resolve_profile_style_targets(style_map: Dict[str, Any]) -> Dict[str, str]:
    style_roles = (style_map.get("style_roles") or {}) if isinstance(style_map, dict) else {}
    targets = dict(DEFAULT_PROFILE_STYLE_TARGETS)
    role_overrides = {
        "heading1": "heading1",
        "heading2": "heading2",
        "heading3": "heading3",
        "body_text": "body_text",
        "caption": "caption",
    }
    for target_key, role_key in role_overrides.items():
        value = style_roles.get(role_key)
        if isinstance(value, str) and value.strip() and value != "template_derived":
            targets[target_key] = value.strip()
    return targets


def front_matter_policy_value(front_matter_policy: Dict[str, Any], key: str, default: Any) -> Any:
    if not isinstance(front_matter_policy, dict):
        return default
    policy = front_matter_policy.get("policy") or {}
    if not isinstance(policy, dict):
        return default
    return policy.get(key, default)


def normalize_rel_target_word(base_part: str, target: str) -> str:
    raw = (Path(base_part).parent / target).as_posix()
    parts: List[str] = []
    for token in raw.split("/"):
        if token in {"", "."}:
            continue
        if token == "..":
            if parts:
                parts.pop()
        else:
            parts.append(token)
    return "/".join(parts)


def parse_relationships(zf: zipfile.ZipFile, part: str) -> Dict[str, Dict[str, str]]:
    root = read_xml(zf, part)
    if root is None:
        return {}
    rels: Dict[str, Dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        if not rel_id:
            continue
        rels[rel_id] = {
            "target": normalize_rel_target_word(part, rel.attrib.get("Target", "")),
            "type": rel.attrib.get("Type", ""),
            "targetMode": rel.attrib.get("TargetMode", ""),
        }
    return rels


def extract_theme(theme_root: Optional[ET.Element]) -> Dict[str, Any]:
    if theme_root is None:
        return {}

    def parse_font_scheme_node(node: Optional[ET.Element]) -> Dict[str, Any]:
        if node is None:
            return {}
        out = {
            "latin": None,
            "eastAsia": None,
            "cs": None,
            "scripts": {},
        }
        latin = node.find("a:latin", NS)
        ea = node.find("a:ea", NS)
        cs = node.find("a:cs", NS)
        if latin is not None:
            out["latin"] = latin.attrib.get("typeface")
        if ea is not None:
            out["eastAsia"] = ea.attrib.get("typeface")
        if cs is not None:
            out["cs"] = cs.attrib.get("typeface")
        for font in node.findall("a:font", NS):
            script = font.attrib.get("script")
            typeface = font.attrib.get("typeface")
            if script and typeface is not None:
                out["scripts"][script] = typeface
        return out

    def parse_color_scheme_node(node: Optional[ET.Element]) -> Dict[str, Any]:
        if node is None:
            return {}
        out: Dict[str, Any] = {}
        for child in list(node):
            name = local_name(child.tag)
            srgb = child.find("a:srgbClr", NS)
            sys_clr = child.find("a:sysClr", NS)
            if srgb is not None:
                out[name] = {"type": "srgb", "value": srgb.attrib.get("val")}
            elif sys_clr is not None:
                out[name] = {
                    "type": "system",
                    "value": sys_clr.attrib.get("lastClr"),
                    "system": sys_clr.attrib.get("val"),
                }
        return out

    font_scheme = theme_root.find(".//a:fontScheme", NS)
    color_scheme = theme_root.find(".//a:clrScheme", NS)
    return {
        "fontSchemeName": font_scheme.attrib.get("name") if font_scheme is not None else None,
        "major": parse_font_scheme_node(font_scheme.find("a:majorFont", NS) if font_scheme is not None else None),
        "minor": parse_font_scheme_node(font_scheme.find("a:minorFont", NS) if font_scheme is not None else None),
        "colors": parse_color_scheme_node(color_scheme),
    }


def pick_script_from_lang(lang: Optional[str]) -> Optional[str]:
    if not lang:
        return None
    lowered = lang.lower()
    if lowered.startswith("zh"):
        if lowered in {"zh-cn", "zh-sg"}:
            return "Hans"
        return "Hant"
    if lowered.startswith("ja"):
        return "Jpan"
    if lowered.startswith("ko"):
        return "Hang"
    return None


def resolve_theme_font(theme_value: Optional[str], slot: str, theme: Dict[str, Any], lang: Dict[str, Any]) -> Optional[str]:
    if not theme_value:
        return None
    lowered = theme_value.lower()
    family = "minor" if lowered.startswith("minor") else "major" if lowered.startswith("major") else None
    if not family:
        return None
    scheme = theme.get(family) or {}
    if slot in {"ascii", "hAnsi"}:
        return scheme.get("latin")
    if slot == "cs":
        return scheme.get("cs") or scheme.get("latin")
    if slot == "eastAsia":
        direct = scheme.get("eastAsia")
        if direct:
            return direct
        script = pick_script_from_lang(lang.get("eastAsia") if isinstance(lang, dict) else None)
        if script:
            return scheme.get("scripts", {}).get(script)
        return scheme.get("latin")
    return None


def extract_font_table(font_root: Optional[ET.Element]) -> Dict[str, Any]:
    if font_root is None:
        return {"fonts": []}
    fonts = []
    for font in font_root.findall("w:font", NS):
        name = font.attrib.get(qn("name"))
        if not name:
            continue
        entry = {"name": name}
        alt = font.find("w:altName", NS)
        charset = font.find("w:charset", NS)
        family = font.find("w:family", NS)
        if alt is not None:
            entry["altName"] = alt.attrib.get(qn("val"))
        if charset is not None:
            entry["charset"] = charset.attrib.get(qn("val"))
        if family is not None:
            entry["family"] = family.attrib.get(qn("val"))
        fonts.append(entry)
    return {"fonts": fonts, "names": [item["name"] for item in fonts]}


def parse_rpr(node: Optional[ET.Element]) -> Dict[str, Any]:
    if node is None:
        return {}
    out: Dict[str, Any] = {}
    rfonts = node.find("w:rFonts", NS)
    if rfonts is not None:
        out["fonts"] = {local_name(key): value for key, value in rfonts.attrib.items()}
    for tag, target in [("b", "bold"), ("bCs", "boldCs"), ("i", "italic"), ("iCs", "italicCs")]:
        value = onoff_value(node.find(f"w:{tag}", NS))
        if value is not None:
            out[target] = value
    sz = node.find("w:sz", NS)
    szcs = node.find("w:szCs", NS)
    if sz is not None and qn("val") in sz.attrib:
        out["sizeHalfPoints"] = safe_int(sz.attrib.get(qn("val")))
    if szcs is not None and qn("val") in szcs.attrib:
        out["sizeCsHalfPoints"] = safe_int(szcs.attrib.get(qn("val")))
    color = node.find("w:color", NS)
    if color is not None:
        out["color"] = {
            "val": color.attrib.get(qn("val")),
            "themeColor": color.attrib.get(qn("themeColor")),
            "themeTint": color.attrib.get(qn("themeTint")),
            "themeShade": color.attrib.get(qn("themeShade")),
        }
    lang = node.find("w:lang", NS)
    if lang is not None:
        out["lang"] = {
            "val": lang.attrib.get(qn("val")),
            "eastAsia": lang.attrib.get(qn("eastAsia")),
            "bidi": lang.attrib.get(qn("bidi")),
        }
    rstyle = node.find("w:rStyle", NS)
    if rstyle is not None and qn("val") in rstyle.attrib:
        out["rStyle"] = rstyle.attrib.get(qn("val"))
    return out


def parse_ppr(node: Optional[ET.Element]) -> Dict[str, Any]:
    if node is None:
        return {}
    out: Dict[str, Any] = {}
    spacing = node.find("w:spacing", NS)
    if spacing is not None:
        out["spacing"] = flatten_attrs(
            spacing,
            ["before", "after", "line", "lineRule", "beforeLines", "afterLines"],
        )
    ind = node.find("w:ind", NS)
    if ind is not None:
        out["ind"] = flatten_attrs(
            ind,
            ["left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars"],
        )
    jc = node.find("w:jc", NS)
    if jc is not None:
        out["jc"] = jc.attrib.get(qn("val"))
    outline = node.find("w:outlineLvl", NS)
    if outline is not None:
        out["outlineLvl"] = safe_int(outline.attrib.get(qn("val")))
    for tag, target in [("keepNext", "keepNext"), ("keepLines", "keepLines"), ("pageBreakBefore", "pageBreakBefore")]:
        value = onoff_value(node.find(f"w:{tag}", NS))
        if value is not None:
            out[target] = value
    pstyle = node.find("w:pStyle", NS)
    if pstyle is not None and qn("val") in pstyle.attrib:
        out["pStyle"] = pstyle.attrib.get(qn("val"))
    return out


def extract_doc_defaults(styles_root: Optional[ET.Element]) -> Dict[str, Any]:
    if styles_root is None:
        return {"pPr": {}, "rPr": {}}
    defaults = styles_root.find("w:docDefaults", NS)
    if defaults is None:
        return {"pPr": {}, "rPr": {}}
    p_default = defaults.find("w:pPrDefault/w:pPr", NS)
    r_default = defaults.find("w:rPrDefault/w:rPr", NS)
    return {"pPr": parse_ppr(p_default), "rPr": parse_rpr(r_default)}


def extract_styles(styles_root: Optional[ET.Element]) -> Dict[str, Any]:
    styles: Dict[str, Any] = {}
    if styles_root is None:
        return styles
    for node in styles_root.findall("w:style", NS):
        style_id = node.attrib.get(qn("styleId"))
        if not style_id:
            continue
        name = node.find("w:name", NS)
        based_on = node.find("w:basedOn", NS)
        linked = node.find("w:link", NS)
        next_style = node.find("w:next", NS)
        ui_priority = node.find("w:uiPriority", NS)
        styles[style_id] = {
            "styleId": style_id,
            "type": node.attrib.get(qn("type")),
            "default": onoff_value(node if qn("default") in node.attrib else None),
            "customStyle": node.attrib.get(qn("customStyle")),
            "name": name.attrib.get(qn("val")) if name is not None else None,
            "basedOn": based_on.attrib.get(qn("val")) if based_on is not None else None,
            "link": linked.attrib.get(qn("val")) if linked is not None else None,
            "next": next_style.attrib.get(qn("val")) if next_style is not None else None,
            "uiPriority": safe_int(ui_priority.attrib.get(qn("val"))) if ui_priority is not None else None,
            "qFormat": node.find("w:qFormat", NS) is not None,
            "semiHidden": node.find("w:semiHidden", NS) is not None,
            "unhideWhenUsed": node.find("w:unhideWhenUsed", NS) is not None,
            "locked": node.find("w:locked", NS) is not None,
            "pPr": parse_ppr(node.find("w:pPr", NS)),
            "rPr": parse_rpr(node.find("w:rPr", NS)),
        }
    return styles


def build_style_chain(style_id: Optional[str], styles: Dict[str, Any]) -> List[Dict[str, Any]]:
    chain: List[Dict[str, Any]] = []
    current = style_id
    visited = set()
    while current and current not in visited and current in styles:
        visited.add(current)
        chain.append(styles[current])
        current = styles[current].get("basedOn")
    chain.reverse()
    return chain


def get_style_name(style_id: Optional[str], styles: Dict[str, Any]) -> Optional[str]:
    if not style_id:
        return None
    return (styles.get(style_id) or {}).get("name")


def extract_settings(settings_root: Optional[ET.Element]) -> Dict[str, Any]:
    if settings_root is None:
        return {}
    compat_settings = []
    compat = settings_root.find("w:compat", NS)
    if compat is not None:
        for setting in compat.findall("w:compatSetting", NS):
            compat_settings.append(
                {
                    "name": setting.attrib.get(qn("name")),
                    "uri": setting.attrib.get(qn("uri")),
                    "val": setting.attrib.get(qn("val")),
                }
            )
    return {
        "zoom": safe_int((settings_root.find("w:zoom", NS) or ET.Element("x")).attrib.get(qn("percent"))),
        "defaultTabStop": safe_int((settings_root.find("w:defaultTabStop", NS) or ET.Element("x")).attrib.get(qn("val"))),
        "embedSystemFonts": settings_root.find("w:embedSystemFonts", NS) is not None,
        "evenAndOddHeaders": settings_root.find("w:evenAndOddHeaders", NS) is not None,
        "characterSpacingControl": (settings_root.find("w:characterSpacingControl", NS) or ET.Element("x")).attrib.get(qn("val")),
        "compatSettings": compat_settings,
    }


def extract_numbering(numbering_root: Optional[ET.Element]) -> Dict[str, Any]:
    result = {"abstractNums": {}, "nums": {}, "styleBindings": {}}
    if numbering_root is None:
        return result
    for abstract in numbering_root.findall("w:abstractNum", NS):
        abstract_id = abstract.attrib.get(qn("abstractNumId"))
        if abstract_id is None:
            continue
        levels = {}
        for lvl in abstract.findall("w:lvl", NS):
            ilvl = safe_int(lvl.attrib.get(qn("ilvl")))
            if ilvl is None:
                continue
            pstyle = lvl.find("w:pStyle", NS)
            numfmt = lvl.find("w:numFmt", NS)
            lvltext = lvl.find("w:lvlText", NS)
            start = lvl.find("w:start", NS)
            suff = lvl.find("w:suff", NS)
            lvl_jc = lvl.find("w:lvlJc", NS)
            levels[str(ilvl)] = {
                "ilvl": ilvl,
                "start": safe_int(start.attrib.get(qn("val"))) if start is not None else None,
                "numFmt": numfmt.attrib.get(qn("val")) if numfmt is not None else None,
                "lvlText": lvltext.attrib.get(qn("val")) if lvltext is not None else None,
                "pStyle": pstyle.attrib.get(qn("val")) if pstyle is not None else None,
                "levelJc": lvl_jc.attrib.get(qn("val")) if lvl_jc is not None else None,
                "suff": suff.attrib.get(qn("val")) if suff is not None else None,
                "pPr": parse_ppr(lvl.find("w:pPr", NS)),
                "rPr": parse_rpr(lvl.find("w:rPr", NS)),
            }
            if pstyle is not None and pstyle.attrib.get(qn("val")):
                result["styleBindings"][pstyle.attrib.get(qn("val"))] = {"abstractNumId": abstract_id, "ilvl": ilvl}
        result["abstractNums"][abstract_id] = {
            "nsid": (abstract.find("w:nsid", NS) or ET.Element("x")).attrib.get(qn("val")),
            "tmpl": (abstract.find("w:tmpl", NS) or ET.Element("x")).attrib.get(qn("val")),
            "multiLevelType": (abstract.find("w:multiLevelType", NS) or ET.Element("x")).attrib.get(qn("val")),
            "levels": levels,
        }
    for num in numbering_root.findall("w:num", NS):
        num_id = num.attrib.get(qn("numId"))
        if not num_id:
            continue
        abstract_num = num.find("w:abstractNumId", NS)
        result["nums"][num_id] = {"abstractNumId": abstract_num.attrib.get(qn("val")) if abstract_num is not None else None}
    return result


def get_numbering_level(numbering: Dict[str, Any], num_id: Optional[str], ilvl: Optional[int]) -> Dict[str, Any]:
    if not num_id or ilvl is None:
        return {}
    instance = (numbering.get("nums") or {}).get(str(num_id))
    if not instance:
        return {}
    abstract = (numbering.get("abstractNums") or {}).get(str(instance.get("abstractNumId")))
    if not abstract:
        return {}
    return copy.deepcopy((abstract.get("levels") or {}).get(str(ilvl), {}))


def merge_rpr_chain(parts: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    result: Dict[str, Any] = {}
    for part in parts:
        if part:
            result = deep_merge(result, part)
    return result


def merge_ppr_chain(parts: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    result: Dict[str, Any] = {}
    for part in parts:
        if part:
            result = deep_merge(result, part)
    return result


def resolve_fonts(raw_fonts: Dict[str, Any], theme: Dict[str, Any], lang: Dict[str, Any], text: str) -> Dict[str, Any]:
    raw_fonts = raw_fonts or {}
    lang = lang or {}
    resolved: Dict[str, Optional[str]] = {}
    notes: List[str] = []
    for slot in FONT_SLOTS:
        direct_value = raw_fonts.get(slot)
        if direct_value:
            resolved[slot] = direct_value
            continue
        theme_value = None
        for theme_attr in FONT_THEME_ATTRS[slot]:
            if raw_fonts.get(theme_attr):
                theme_value = raw_fonts.get(theme_attr)
                break
        resolved[slot] = resolve_theme_font(theme_value, slot, theme, lang)
        if resolved[slot] and theme_value:
            notes.append(f"{slot}_from_theme:{theme_value}")
    if not resolved.get("hAnsi") and resolved.get("ascii"):
        resolved["hAnsi"] = resolved["ascii"]
    if not resolved.get("ascii") and resolved.get("hAnsi"):
        resolved["ascii"] = resolved["hAnsi"]
    if contains_cjk(text):
        if not raw_fonts.get("eastAsia"):
            if resolved.get("eastAsia"):
                notes.append("cjk_uses_inherited_eastAsia_font")
            else:
                notes.append("cjk_without_resolved_eastAsia_font")
        if (raw_fonts.get("ascii") or raw_fonts.get("hAnsi")) and not raw_fonts.get("eastAsia"):
            notes.append("latin_font_set_without_eastAsia_slot")
    return {"raw": raw_fonts, "resolved": resolved, "notes": notes}


def effective_spacing(spacing: Dict[str, Any]) -> Dict[str, Any]:
    if not spacing:
        return {}
    return {
        "beforeTwips": safe_int(spacing.get("before")),
        "afterTwips": safe_int(spacing.get("after")),
        "lineTwips": safe_int(spacing.get("line")),
        "lineRule": spacing.get("lineRule"),
        "beforePt": twips_to_pt(spacing.get("before")),
        "afterPt": twips_to_pt(spacing.get("after")),
        "linePt": twips_to_pt(spacing.get("line")),
    }


def effective_ind(ind: Dict[str, Any]) -> Dict[str, Any]:
    if not ind:
        return {}
    out = {
        "leftTwips": safe_int(ind.get("left")),
        "rightTwips": safe_int(ind.get("right")),
        "firstLineTwips": safe_int(ind.get("firstLine")),
        "hangingTwips": safe_int(ind.get("hanging")),
        "leftChars": safe_int(ind.get("leftChars")),
        "rightChars": safe_int(ind.get("rightChars")),
        "firstLineChars": safe_int(ind.get("firstLineChars")),
        "hangingChars": safe_int(ind.get("hangingChars")),
    }
    if out["firstLineChars"] is None and out["firstLineTwips"] is not None:
        out["firstLineCharsApprox"] = round(out["firstLineTwips"] / 200.0, 2)
    return {key: value for key, value in out.items() if value is not None}


def resolve_effective_rpr(
    doc_defaults: Dict[str, Any],
    paragraph_style_chain: List[Dict[str, Any]],
    run_style_chain: List[Dict[str, Any]],
    numbering_level: Dict[str, Any],
    direct_rpr: Dict[str, Any],
    theme: Dict[str, Any],
    text: str,
) -> Dict[str, Any]:
    merged = merge_rpr_chain(
        [
            doc_defaults.get("rPr") or {},
            *[item.get("rPr") or {} for item in paragraph_style_chain],
            numbering_level.get("rPr") or {},
            *[item.get("rPr") or {} for item in run_style_chain],
            direct_rpr or {},
        ]
    )
    lang = merged.get("lang") or {}
    fonts = resolve_fonts(merged.get("fonts") or {}, theme, lang, text)
    size_half_points = merged.get("sizeHalfPoints") or merged.get("sizeCsHalfPoints")
    size_cs_half_points = merged.get("sizeCsHalfPoints") or merged.get("sizeHalfPoints")
    return {
        "fonts": fonts,
        "sizeHalfPoints": size_half_points,
        "sizePt": half_points_to_pt(size_half_points),
        "sizeCsHalfPoints": size_cs_half_points,
        "sizeCsPt": half_points_to_pt(size_cs_half_points),
        "bold": merged.get("bold"),
        "italic": merged.get("italic"),
        "color": merged.get("color"),
        "lang": lang,
        "rawMerged": merged,
    }


def resolve_effective_ppr(
    doc_defaults: Dict[str, Any],
    paragraph_style_chain: List[Dict[str, Any]],
    numbering_level: Dict[str, Any],
    direct_ppr: Dict[str, Any],
) -> Dict[str, Any]:
    merged = merge_ppr_chain(
        [
            doc_defaults.get("pPr") or {},
            *[item.get("pPr") or {} for item in paragraph_style_chain],
            numbering_level.get("pPr") or {},
            direct_ppr or {},
        ]
    )
    return {
        "jc": merged.get("jc"),
        "spacing": effective_spacing(merged.get("spacing") or {}),
        "ind": effective_ind(merged.get("ind") or {}),
        "keepNext": merged.get("keepNext"),
        "keepLines": merged.get("keepLines"),
        "pageBreakBefore": merged.get("pageBreakBefore"),
        "outlineLvl": merged.get("outlineLvl"),
        "rawMerged": merged,
    }


def cn_to_int(value: Optional[str]) -> Optional[int]:
    if not value:
        return None
    if value == "十":
        return 10
    if "十" in value:
        left, _, right = value.partition("十")
        total = (CN_DIGITS.get(left, 1) if left else 1) * 10
        total += CN_DIGITS.get(right, 0) if right else 0
        return total
    total = 0
    for char in value:
        digit = CN_DIGITS.get(char)
        if digit is None or digit >= 10:
            return None
        total = total * 10 + digit
    return total


def int_to_cn(value: int) -> str:
    digits = "零一二三四五六七八九"
    if value <= 10:
        return "十" if value == 10 else digits[value]
    tens, ones = divmod(value, 10)
    prefix = "十" if value < 20 else f"{digits[tens]}十"
    if ones:
        prefix += digits[ones]
    return prefix


def parse_manual_numbering(text: str) -> Optional[Dict[str, Any]]:
    stripped = (text or "").strip()
    if not stripped:
        return None
    match = CHAPTER_RE.match(stripped)
    if match:
        value = cn_to_int(match.group("prefix")[1:-1])
        return {
            "family": "humanities_chapter",
            "level": 1,
            "prefix": match.group("prefix"),
            "separator": match.group("tail"),
            "values": [value] if value is not None else None,
        }
    match = SECTION_RE.match(stripped)
    if match:
        value = cn_to_int(match.group("prefix")[1:-1])
        return {
            "family": "humanities_chapter",
            "level": 2,
            "prefix": match.group("prefix"),
            "separator": match.group("tail"),
            "values": [value] if value is not None else None,
        }
    match = CN_LEVEL1_RE.match(stripped)
    if match:
        value = cn_to_int(match.group("prefix").rstrip("、"))
        return {
            "family": "humanities_cn",
            "level": 1,
            "prefix": match.group("prefix"),
            "separator": match.group("tail"),
            "values": [value] if value is not None else None,
        }
    match = CN_LEVEL2_RE.match(stripped)
    if match:
        value = cn_to_int(match.group("prefix").strip("（）"))
        return {
            "family": "humanities_cn",
            "level": 2,
            "prefix": match.group("prefix"),
            "separator": match.group("tail"),
            "values": [value] if value is not None else None,
        }
    match = SCIENCE_DECIMAL_RE.match(stripped)
    if match:
        prefix = match.group("prefix")
        values = [safe_int(part) for part in prefix.split(".")]
        if all(item is not None for item in values):
            return {
                "family": "science_decimal",
                "level": len(values),
                "prefix": prefix,
                "separator": match.group("tail"),
                "values": values,
            }
    match = CN_NUMERIC_LEVEL_RE.match(stripped)
    if match:
        value = safe_int(match.group("prefix").rstrip("、"))
        return {
            "family": "humanities_cn",
            "level": 3,
            "prefix": match.group("prefix"),
            "separator": match.group("tail"),
            "values": [value] if value is not None else None,
        }
    match = CN_NUMERIC_PAREN_RE.match(stripped)
    if match:
        value = safe_int(match.group("prefix").strip("（）"))
        return {
            "family": "humanities_cn",
            "level": 4,
            "prefix": match.group("prefix"),
            "separator": match.group("tail"),
            "values": [value] if value is not None else None,
        }
    return None


def normalize_front_matter_text(text: str) -> str:
    if not text:
        return ""
    normalized = text.replace("\u3000", "").replace(" ", "")
    normalized = normalized.replace("：", ":")
    return normalized


def dominant_run_summary(runs: List[Dict[str, Any]]) -> Dict[str, Any]:
    for run in runs:
        if (run.get("text") or "").strip():
            return run.get("effectiveFormat") or {}
    return {}


def classify_paragraph_role(paragraph: Dict[str, Any]) -> Dict[str, Any]:
    text = (paragraph.get("text") or "").strip()
    compact_text = normalize_front_matter_text(text)
    style_id = paragraph.get("styleId") or ""
    style_name = (paragraph.get("styleName") or "").lower()
    effective_ppr = paragraph.get("effectiveParagraph") or {}
    effective_run = paragraph.get("effectiveRunSummary") or {}
    manual_numbering = paragraph.get("manualNumbering")
    has_numpr = paragraph.get("numId") is not None

    if not text:
        return {"role": "blank", "confidence": 1.0, "signals": []}
    if style_name.startswith("toc") or MANUAL_TOC_RE.search(text):
        return {"role": "toc_entry", "confidence": 0.95, "signals": ["toc_style_or_leaders"]}
    if FIGURE_CAPTION_RE.match(text):
        return {"role": "figure_caption", "confidence": 0.98, "signals": ["caption_pattern"]}
    if TABLE_CAPTION_RE.match(text):
        return {"role": "table_caption", "confidence": 0.98, "signals": ["caption_pattern"]}
    if paragraph.get("hasMath") and EQUATION_NUMBER_RE.search(text):
        return {"role": "equation_block", "confidence": 0.92, "signals": ["omml_equation", "equation_number"]}
    if paragraph.get("hasMath"):
        return {"role": "equation_block", "confidence": 0.75, "signals": ["omml_equation"]}
    if style_id == "zafu_toc_title":
        return {"role": "toc_heading", "confidence": 0.99, "signals": ["repaired_toc_title_style"]}
    if style_id == "zafu_toc_body":
        return {"role": "toc_field", "confidence": 0.99, "signals": ["repaired_toc_body_style"]}
    if style_id == "zafu_title_cn":
        return {"role": "title_cn", "confidence": 0.99, "signals": ["repaired_title_cn_style"]}
    if style_id == "zafu_title_en":
        return {"role": "title_en", "confidence": 0.99, "signals": ["repaired_title_en_style"]}
    if paragraph.get("hasTocField"):
        return {"role": "toc_field", "confidence": 0.99, "signals": ["toc_field_code"]}
    if text.replace(" ", "") == "目录":
        return {"role": "toc_heading", "confidence": 0.99, "signals": ["exact_heading"]}
    if text.startswith("[图形占位]"):
        return {"role": "image_block", "confidence": 0.99, "signals": ["markdown_image_placeholder"]}
    if style_name == "title":
        chinese_count = len(re.findall(r"[\u4e00-\u9fff]", text))
        latin_count = len(re.findall(r"[A-Za-z]", text))
        if chinese_count and (latin_count == 0 or chinese_count >= latin_count):
            return {"role": "title_cn", "confidence": 0.99, "signals": ["source_title_style"]}
        if re.search(r"[A-Za-z]", text):
            return {"role": "title_en", "confidence": 0.99, "signals": ["source_title_style"]}
    if compact_text.startswith("摘要:"):
        return {"role": "abstract_paragraph_cn", "confidence": 0.98, "signals": ["abstract_label_inline_compact"]}
    if compact_text in {"摘要", "中文摘要"}:
        return {"role": "abstract_heading_cn", "confidence": 0.99, "signals": ["exact_heading_compact"]}
    if compact_text.startswith("关键词:"):
        return {"role": "keywords_cn", "confidence": 0.96, "signals": ["keywords_label_compact"]}
    if compact_text in {"中文摘要", "摘要", "摘要"} or text in {"中文摘要", "摘  要", "摘要"}:
        return {"role": "abstract_heading_cn", "confidence": 0.99, "signals": ["exact_heading"]}
    if text in {"English Abstract", "ABSTRACT"}:
        return {"role": "abstract_heading_en", "confidence": 0.99, "signals": ["exact_heading"]}
    if text.startswith("摘要：") or text.startswith("摘要:") or compact_text.startswith("摘要:"):
        return {"role": "abstract_paragraph_cn", "confidence": 0.98, "signals": ["abstract_label_inline"]}
    if text.startswith("Abstract:") or text.startswith("ABSTRACT:"):
        return {"role": "abstract_paragraph_en", "confidence": 0.98, "signals": ["abstract_label_inline"]}
    if text.startswith("关键词"):
        return {"role": "keywords_cn", "confidence": 0.96, "signals": ["keywords_label"]}
    if text.startswith("Key words") or text.startswith("Key Words") or text.startswith("KEY WORDS") or text.startswith("Keywords"):
        return {"role": "keywords_en", "confidence": 0.96, "signals": ["keywords_label"]}
    if text in {"参考文献"}:
        return {"role": "references_heading", "confidence": 0.97, "signals": ["exact_heading"]}
    if text in {"致谢", "致  谢"}:
        return {"role": "ack_heading", "confidence": 0.97, "signals": ["exact_heading"]}
    if text.startswith("附录") or text.startswith("附 录"):
        return {"role": "appendix_heading", "confidence": 0.94, "signals": ["exact_heading"]}

    def is_chinese_outline_heading() -> bool:
        if not manual_numbering:
            return False
        family = manual_numbering.get("family")
        level = manual_numbering.get("level") or 0
        if family not in {"humanities_cn", "humanities_chapter"}:
            return False
        if level not in {1, 2}:
            return False
        if len(text) > 60:
            return False
        if text.endswith(("；", ";", "。")):
            return False
        if text.endswith(("：", ":")):
            return True
        if text.startswith(("第", "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、", "（一）", "（二）", "（三）", "（四）", "（五）")):
            return True
        return False

    signals: List[str] = []
    heading_score = 0.0
    inferred_level = None
    inferred_family = None

    if "heading" in style_name:
        heading_score += 3.5
        signals.append("heading_style_name")
    if paragraph.get("outlineLvl") is not None:
        inferred_level = int(paragraph["outlineLvl"]) + 1
        heading_score += 3.0
        signals.append("outline_level")
    if manual_numbering:
        inferred_level = manual_numbering.get("level") or inferred_level
        inferred_family = manual_numbering.get("family")
        heading_score += 1.8
        signals.append("manual_numbering")
    if has_numpr:
        heading_score += 1.5
        signals.append("numPr")
    if len(text) <= 40:
        heading_score += 0.8
        signals.append("short_line")
    elif len(text) >= 90:
        heading_score -= 1.0
        signals.append("long_line")
    if not text.endswith(tuple(PUNCTUATION_ENDINGS)):
        heading_score += 0.6
        signals.append("no_sentence_ending")
    else:
        heading_score -= 1.0
        signals.append("sentence_ending")
    if effective_ppr.get("jc") == "center":
        heading_score += 0.7
        signals.append("centered")
    spacing = effective_ppr.get("spacing") or {}
    if (spacing.get("beforePt") or 0) > 0 or (spacing.get("afterPt") or 0) > 0:
        heading_score += 0.8
        signals.append("spacing_before_after")
    if effective_run.get("bold"):
        heading_score += 0.7
        signals.append("bold")
    structural_cue = any(signal in signals for signal in ["heading_style_name", "outline_level", "centered", "spacing_before_after"])
    if manual_numbering and not structural_cue:
        heading_score -= 1.2
        signals.append("manual_numbering_without_structural_cue")
    if manual_numbering and text.endswith(("：", ":")):
        heading_score -= 0.8
        signals.append("colon_suffix")
    if manual_numbering and len(text) > 60:
        heading_score -= 0.8
        signals.append("long_manual_enumeration")
    if text.endswith(("；", ";", "。")) and manual_numbering:
        heading_score -= 1.7
        signals.append("enumeration_like_punctuation")
    if is_chinese_outline_heading():
        heading_score += 2.2
        signals.append("chinese_outline_heading")
    if (
        manual_numbering
        and inferred_family == "science_decimal"
        and (manual_numbering.get("level") or 0) in {1, 2, 3}
        and len(text) <= 40
        and not text.endswith(tuple(PUNCTUATION_ENDINGS))
        and not text.endswith(("：", ":"))
    ):
        level = manual_numbering.get("level") or inferred_level or 1
        return {
            "role": f"heading{min(max(level, 1), 4)}",
            "confidence": 0.86,
            "signals": signals + ["science_decimal_short_heading"],
            "family": inferred_family,
        }

    if heading_score >= 4.5 and (structural_cue or "heading_style_name" in signals or "outline_level" in signals):
        level = inferred_level or 1
        return {
            "role": f"heading{min(max(level, 1), 4)}",
            "confidence": round(min(0.99, 0.55 + heading_score / 10.0), 2),
            "signals": signals,
            "family": inferred_family,
        }
    if manual_numbering:
        if is_chinese_outline_heading():
            level = inferred_level or 1
            return {
                "role": f"heading{min(max(level, 1), 4)}",
                "confidence": 0.84,
                "signals": signals + ["chinese_outline_heading_override"],
                "family": inferred_family,
            }
        return {
            "role": "body_enumeration",
            "confidence": 0.68,
            "signals": signals + ["manual_numbering_but_not_heading"],
            "family": inferred_family,
        }
    return {"role": "body", "confidence": 0.7, "signals": signals}
def format_prefix(family: str, values: List[int]) -> Optional[str]:
    if not family or not values:
        return None
    level = len(values)
    if family == "science_decimal":
        return ".".join(str(value) for value in values[:level])
    if family == "humanities_cn":
        if level == 1:
            return f"{int_to_cn(values[0])}、"
        if level == 2:
            return f"（{int_to_cn(values[1])}）"
        if level == 3:
            return f"{values[2]}、"
        if level >= 4:
            return f"（{values[3]}）"
    if family == "humanities_chapter":
        if level == 1:
            return f"第{int_to_cn(values[0])}章"
        if level == 2:
            return f"第{int_to_cn(values[1])}节"
        if level == 3:
            return f"{int_to_cn(values[2])}、"
        if level >= 4:
            return f"（{int_to_cn(values[3])}）"
    return None


def desired_heading_separator(item: Dict[str, Any]) -> str:
    if item.get("manualNumbering"):
        return "  "
    return ""


def analyze_styles(styles: Dict[str, Any]) -> Dict[str, Any]:
    heading_risks = []
    linked_heading_styles = []
    default_styles = []
    for style_id, style in styles.items():
        name = (style.get("name") or "").lower()
        if style.get("default"):
            default_styles.append(style_id)
        if "heading" in name or style_id in {"Heading1", "Heading2", "Heading3", "1", "2", "3"}:
            entry = {
                "styleId": style_id,
                "name": style.get("name"),
                "outlineLvl": style.get("pPr", {}).get("outlineLvl"),
                "color": style.get("rPr", {}).get("color"),
                "fonts": style.get("rPr", {}).get("fonts"),
                "sizePt": half_points_to_pt(style.get("rPr", {}).get("sizeHalfPoints")),
            }
            if style.get("link") or style.get("basedOn") == "a":
                linked_heading_styles.append(entry)
            if not style.get("rPr", {}).get("fonts"):
                heading_risks.append({**entry, "risk": "heading_style_without_explicit_fonts"})
            if style.get("rPr", {}).get("color", {}).get("themeColor"):
                heading_risks.append({**entry, "risk": "heading_style_uses_theme_color"})
    return {
        "defaultStyles": default_styles,
        "builtinHeadingRisks": heading_risks,
        "linkedHeadingStyles": linked_heading_styles,
    }


def analyze_font_slots(paragraphs: List[Dict[str, Any]]) -> Dict[str, Any]:
    suspicious = []
    explicit_slot_counts = Counter()
    for paragraph in paragraphs:
        for run in paragraph.get("runs", []):
            fonts = ((run.get("effectiveFormat") or {}).get("fonts") or {}).get("raw") or {}
            for key in FONT_SLOTS:
                if fonts.get(key):
                    explicit_slot_counts[key] += 1
            notes = ((run.get("effectiveFormat") or {}).get("fonts") or {}).get("notes") or []
            if notes and (run.get("text") or "").strip():
                suspicious.append(
                    {
                        "paragraphIndex": paragraph["index"],
                        "runIndex": run["index"],
                        "text": run["text"][:60],
                        "notes": notes,
                    }
                )
    return {"explicitSlotCounts": dict(explicit_slot_counts), "suspiciousRuns": suspicious[:200]}


def analyze_numbering_tree(paragraphs: List[Dict[str, Any]]) -> Dict[str, Any]:
    heading_candidates = []
    family_scores = Counter()
    mixed_manual_and_auto = []
    caption_items = []

    for paragraph in paragraphs:
        role = (paragraph.get("role") or {}).get("role")
        confidence = (paragraph.get("role") or {}).get("confidence", 0.0)
        manual = paragraph.get("manualNumbering")
        if manual and paragraph.get("numId") is not None:
            mixed_manual_and_auto.append(
                {
                    "paragraphIndex": paragraph["index"],
                    "text": paragraph["text"][:120],
                    "manualPrefix": manual.get("prefix"),
                    "numId": paragraph.get("numId"),
                    "ilvl": paragraph.get("ilvl"),
                }
            )
        if role in {"figure_caption", "table_caption"}:
            caption_items.append({"paragraphIndex": paragraph["index"], "role": role, "text": paragraph["text"][:120]})
        if role and role.startswith("heading"):
            heading_level = safe_int(role.replace("heading", "")) or (manual.get("level") if manual else 1)
            family = (paragraph.get("role") or {}).get("family") or (manual or {}).get("family")
            family_scores[family or "unknown"] += max(confidence, 0.1)
            heading_candidates.append(
                {
                    "paragraphIndex": paragraph["index"],
                    "text": paragraph["text"],
                    "level": heading_level or 1,
                    "family": family,
                    "confidence": confidence,
                    "manualNumbering": manual,
                    "numId": paragraph.get("numId"),
                    "ilvl": paragraph.get("ilvl"),
                }
            )

    dominant_family = None
    for family, _score in family_scores.most_common():
        if family and family != "unknown":
            dominant_family = family
            break

    counters = [0, 0, 0, 0]
    tree = []
    anomalies = []
    repair_actions = []
    for item in heading_candidates:
        level = min(max(item.get("level") or 1, 1), 4)
        counters[level - 1] += 1
        for deeper in range(level, len(counters)):
            counters[deeper] = 0
        expected_values = counters[:level]
        family = item.get("family") or dominant_family
        expected_prefix = format_prefix(family, expected_values) if family else None
        manual_numbering = item.get("manualNumbering") or {}
        actual_prefix = manual_numbering.get("prefix")
        actual_separator = manual_numbering.get("separator") or ""
        actual_values = manual_numbering.get("values")
        desired_separator = desired_heading_separator(item)
        mismatch = False
        if actual_values and len(actual_values) >= level:
            if actual_values[:level] != expected_values:
                mismatch = True
        elif actual_prefix and expected_prefix and actual_prefix != expected_prefix:
            mismatch = True
        if item.get("numId") is not None and actual_prefix:
            mismatch = True
            anomalies.append(
                {
                    "paragraphIndex": item["paragraphIndex"],
                    "kind": "manual_and_auto_numbering_mixed",
                    "text": item["text"][:120],
                }
            )
        if mismatch:
            anomalies.append(
                {
                    "paragraphIndex": item["paragraphIndex"],
                    "kind": "heading_sequence_or_prefix_mismatch",
                    "text": item["text"][:120],
                    "expectedPrefix": expected_prefix,
                    "actualPrefix": actual_prefix,
                }
            )
            if expected_prefix and actual_prefix and item.get("confidence", 0) >= 0.8:
                repair_actions.append(
                    {
                        "paragraphIndex": item["paragraphIndex"],
                        "action": "replace_heading_prefix",
                        "subtype": "prefix_value_or_separator",
                        "oldPrefix": actual_prefix + actual_separator,
                        "newPrefix": expected_prefix + desired_separator,
                        "confidence": item.get("confidence"),
                    }
                )
        elif actual_prefix and actual_separator != desired_separator and item.get("confidence", 0) >= 0.8:
            repair_actions.append(
                {
                    "paragraphIndex": item["paragraphIndex"],
                    "action": "replace_heading_prefix",
                    "subtype": "separator_only",
                    "oldPrefix": actual_prefix + actual_separator,
                    "newPrefix": actual_prefix + desired_separator,
                    "confidence": item.get("confidence"),
                }
            )
        tree.append(
            {
                "paragraphIndex": item["paragraphIndex"],
                "level": level,
                "text": item["text"][:160],
                "family": family,
                "actualPrefix": actual_prefix,
                "expectedPrefix": expected_prefix,
                "numId": item.get("numId"),
                "ilvl": item.get("ilvl"),
                "confidence": item.get("confidence"),
            }
        )

    families_present = [family for family, score in family_scores.items() if family and score > 0]
    issues = []
    normalized_families = {family for family in families_present if family != "unknown"}
    expected_mixed_set = {"humanities_chapter", "science_decimal"}
    if len(normalized_families) > 1 and normalized_families != expected_mixed_set:
        issues.append({"kind": "mixed_numbering_families", "families": families_present})
    if mixed_manual_and_auto:
        issues.append({"kind": "mixed_manual_and_auto_numbering", "count": len(mixed_manual_and_auto)})
    if anomalies:
        issues.append({"kind": "heading_sequence_anomalies", "count": len(anomalies)})

    return {
        "dominantFamily": dominant_family,
        "familyScores": dict(family_scores),
        "familiesPresent": families_present,
        "headingTree": tree,
        "issues": issues,
        "anomalies": anomalies,
        "mixedManualAndAuto": mixed_manual_and_auto,
        "repairActions": repair_actions,
        "captions": caption_items,
    }


def analyze_caption_layout(body_sequence: List[Dict[str, Any]]) -> Dict[str, Any]:
    issues = []
    caption_items = []
    for index, item in enumerate(body_sequence):
        if item.get("kind") != "paragraph":
            continue
        role = item.get("role")
        if role not in {"table_caption", "figure_caption"}:
            continue
        caption_items.append({"paragraphIndex": item.get("paragraphIndex"), "role": role, "text": item.get("text", "")[:120]})
        previous = next((candidate for candidate in reversed(body_sequence[:index]) if candidate.get("kind") != "blank"), None)
        next_item = next((candidate for candidate in body_sequence[index + 1 :] if candidate.get("kind") != "blank"), None)
        if role == "table_caption" and next_item and next_item.get("kind") == "table":
            issues.append(
                {
                    "kind": "caption_before_table",
                    "paragraphIndex": item.get("paragraphIndex"),
                    "text": item.get("text", "")[:120],
                }
            )
        if role == "table_caption" and not (previous and previous.get("kind") == "table"):
            issues.append(
                {
                    "kind": "table_caption_not_after_table",
                    "paragraphIndex": item.get("paragraphIndex"),
                    "text": item.get("text", "")[:120],
                }
            )
        previous_is_figure_block = bool(
            previous
            and (
                previous.get("role") == "image_block"
                or previous.get("hasDrawing")
                or previous.get("drawingCount", 0) > 0
            )
        )
        if role == "figure_caption" and not previous_is_figure_block:
            issues.append(
                {
                    "kind": "figure_caption_not_after_figure_block",
                    "paragraphIndex": item.get("paragraphIndex"),
                    "text": item.get("text", "")[:120],
                }
            )
    return {"issues": issues, "captions": caption_items}


def analyze_front_matter_structure(paragraphs: List[Dict[str, Any]]) -> Dict[str, Any]:
    role_positions: Dict[str, int] = {}
    for paragraph in paragraphs:
        role = (paragraph.get("role") or {}).get("role")
        if role and role not in role_positions:
            role_positions[role] = paragraph["index"]
    expected_order = ["toc_heading", "toc_field", "title_cn", "abstract_paragraph_cn", "keywords_cn", "title_en", "abstract_paragraph_en", "keywords_en"]
    issues = []
    present = [role for role in expected_order if role in role_positions]
    for earlier, later in zip(present, present[1:]):
        if role_positions[earlier] > role_positions[later]:
            issues.append(
                {
                    "kind": "front_matter_order_mismatch",
                    "earlierRole": earlier,
                    "laterRole": later,
                    "positions": {earlier: role_positions[earlier], later: role_positions[later]},
                }
            )
    required_roles = ["title_cn", "abstract_paragraph_cn", "keywords_cn", "title_en", "abstract_paragraph_en", "keywords_en"]
    missing = [role for role in required_roles if role not in role_positions]
    if missing:
        issues.append({"kind": "front_matter_missing_roles", "roles": missing})
    if "toc_heading" in role_positions and "abstract_paragraph_cn" in role_positions and role_positions["toc_heading"] > role_positions["abstract_paragraph_cn"]:
        issues.append(
            {
                "kind": "toc_after_abstract",
                "tocParagraph": role_positions["toc_heading"],
                "abstractParagraph": role_positions["abstract_paragraph_cn"],
            }
        )
    if "title_cn" not in role_positions:
        title_candidates = [
            paragraph["index"]
            for paragraph in paragraphs
            if (paragraph.get("text") or "").strip()
            and (paragraph.get("role") or {}).get("role") in {"heading1", "body"}
            and paragraph["index"] >= role_positions.get("toc_field", -1)
        ]
        if title_candidates:
            issues.append(
                {
                    "kind": "front_matter_title_candidate_missing",
                    "paragraphIndex": title_candidates[0],
                }
            )
    return {"rolePositions": role_positions, "issues": issues}


def analyze_cross_reference_candidates(paragraphs: List[Dict[str, Any]], captions: List[Dict[str, Any]]) -> Dict[str, Any]:
    caption_map = {}
    for item in captions:
        text = item.get("text") or ""
        match = CROSS_REF_RE.match(text)
        if match:
            caption_map[f"{match.group('kind')}{match.group('label').replace('—', '-').replace('．', '.').replace('.', '-')}"] = item
    candidates = []
    for paragraph in paragraphs:
        role = (paragraph.get("role") or {}).get("role")
        if role in {"figure_caption", "table_caption", "toc_heading", "toc_field", "title_cn", "title_en"}:
            continue
        text = paragraph.get("text") or ""
        for match in CROSS_REF_RE.finditer(text):
            label_key = f"{match.group('kind')}{match.group('label').replace('—', '-').replace('．', '.').replace('.', '-')}"
            candidates.append(
                {
                    "paragraphIndex": paragraph["index"],
                    "text": text[:160],
                    "reference": match.group(0),
                    "resolvedCaption": caption_map.get(label_key, {}).get("text"),
                    "resolved": label_key in caption_map,
                }
            )
    unresolved = [item for item in candidates if not item.get("resolved")]
    return {"candidates": candidates[:200], "unresolved": unresolved[:100]}


def extract_header_footer_parts(zf: zipfile.ZipFile) -> Dict[str, Any]:
    parts: Dict[str, Any] = {}
    for name in zf.namelist():
        if not name.startswith("word/"):
            continue
        if not (name.startswith("word/header") or name.startswith("word/footer")):
            continue
        if not name.endswith(".xml"):
            continue
        root = read_xml(zf, name)
        if root is None:
            continue
        paragraphs = root.findall(".//w:p", NS)
        texts = [paragraph_text(p) for p in paragraphs]
        parts[name] = {
            "text": "".join(texts).strip(),
            "paragraphCount": len(paragraphs),
            "paragraphs": [text.strip() for text in texts if text.strip()],
        }
    return parts


def parse_sections(
    document_root: Optional[ET.Element],
    document_rels: Dict[str, Dict[str, str]],
    parts_text: Dict[str, Any],
) -> List[Dict[str, Any]]:
    sections: List[Dict[str, Any]] = []
    if document_root is None:
        return sections

    def parse_sectpr(sect: ET.Element, paragraph_index: Optional[int]) -> Dict[str, Any]:
        page = {}
        pg_mar = sect.find("w:pgMar", NS)
        if pg_mar is not None:
            for key in ["top", "bottom", "left", "right", "header", "footer", "gutter"]:
                raw = pg_mar.attrib.get(qn(key))
                if raw is not None:
                    page[key] = {"twips": safe_int(raw), "cm": twips_to_cm(raw)}
        pg_sz = sect.find("w:pgSz", NS)
        if pg_sz is not None:
            page["size"] = {
                "widthTwips": safe_int(pg_sz.attrib.get(qn("w"))),
                "heightTwips": safe_int(pg_sz.attrib.get(qn("h"))),
                "orient": pg_sz.attrib.get(qn("orient"), "portrait"),
            }
        cols = sect.find("w:cols", NS)
        header_refs = []
        footer_refs = []
        for tag, container in [("headerReference", header_refs), ("footerReference", footer_refs)]:
            for ref in sect.findall(f"w:{tag}", NS):
                rel_id = ref.attrib.get(qn("id"))
                rel = document_rels.get(rel_id or "", {})
                target = rel.get("target")
                container.append(
                    {
                        "type": ref.attrib.get(qn("type"), "default"),
                        "rId": rel_id,
                        "target": target,
                        "text": parts_text.get(target, {}).get("text") if target else None,
                    }
                )
        pg_num_type = sect.find("w:pgNumType", NS)
        section_type = sect.find("w:type", NS)
        title_pg = sect.find("w:titlePg", NS) is not None
        return {
            "paragraphIndex": paragraph_index,
            "page": page,
            "columns": {
                "count": safe_int(cols.attrib.get(qn("num"))) if cols is not None else None,
                "space": safe_int(cols.attrib.get(qn("space"))) if cols is not None else None,
            },
            "headerReferences": header_refs,
            "footerReferences": footer_refs,
            "pageNumbering": {
                "start": safe_int(pg_num_type.attrib.get(qn("start"))) if pg_num_type is not None else None,
                "format": pg_num_type.attrib.get(qn("fmt")) if pg_num_type is not None else None,
            },
            "sectionType": section_type.attrib.get(qn("val")) if section_type is not None else None,
            "titlePg": title_pg,
        }

    paragraphs = document_root.findall(".//w:body/w:p", NS)
    for index, paragraph in enumerate(paragraphs):
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            continue
        sect = ppr.find("w:sectPr", NS)
        if sect is not None:
            sections.append(parse_sectpr(sect, index))
    body_sect = document_root.find(".//w:body/w:sectPr", NS)
    if body_sect is not None:
        sections.append(parse_sectpr(body_sect, None))
    return sections


def detect_immutable_prefix_region(paragraphs: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not paragraphs:
        return None
    page_break_index = None
    promise_index = None
    section_break_index = None
    for paragraph in paragraphs[:80]:
        text = (paragraph.get("text") or "").strip()
        if page_break_index is None and "page" in (paragraph.get("breakTypes") or []):
            page_break_index = paragraph["index"]
        if promise_index is None and ("诚信承诺书" in text or "承诺书" in text):
            promise_index = paragraph["index"]
        if paragraph.get("hasSectPr") and page_break_index is not None:
            if paragraph["index"] > page_break_index:
                section_break_index = paragraph["index"]
                break
    if page_break_index is None or promise_index is None or section_break_index is None:
        return None
    return {
        "startParagraph": 0,
        "endParagraphExclusive": section_break_index + 1,
        "reason": "cover_and_integrity_pages_should_be_preserved_as_immutable_prefix",
        "confidence": 0.98,
        "signals": {
            "pageBreakParagraph": page_break_index,
            "integrityParagraph": promise_index,
            "sectionBreakParagraph": section_break_index,
        },
    }


def preferred_header_text(template_baseline: Optional[Dict[str, Any]], rules: Dict[str, Any]) -> Optional[str]:
    if template_baseline:
        parts = template_baseline.get("headerFooterParts") or {}
        values = sorted({(item.get("text") or "").strip() for name, item in parts.items() if "header" in name and (item.get("text") or "").strip()})
        if len(values) == 1:
            return values[0]
        if not values:
            return None
    return (rules.get("page") or {}).get("header_text") if rules else None


def detect_template_prefix_count(template_baseline: Optional[Dict[str, Any]]) -> int:
    if not template_baseline:
        return 0
    immutable = ((template_baseline.get("preservationHints") or {}).get("immutablePrefix")) or {}
    immutable_end = safe_int(immutable.get("endParagraphExclusive")) or 0
    if immutable_end > 0:
        return immutable_end
    paragraphs = template_baseline.get("paragraphs") or template_baseline.get("paragraphProbe") or []
    paragraph_count = safe_int((template_baseline.get("summary") or {}).get("paragraphCount")) or 0
    if paragraph_count <= 0:
        probe = template_baseline.get("paragraphProbe") or []
        if probe:
            paragraph_count = (probe[-1].get("index") or 0) + 1
    probe_text = "".join((item.get("text") or "") for item in paragraphs[: min(40, len(paragraphs))])
    if paragraph_count > 0 and any(keyword in probe_text for keyword in ["诚信承诺书", "本科生毕业论文（设计）诚信承诺书", "本科生毕业论文诚信承诺书"]):
        return paragraph_count if paragraph_count <= 120 else 0
    return 0


def recommended_defaults_from_rules(rules: Dict[str, Any]) -> Dict[str, Any]:
    if not rules:
        return {}
    page = rules.get("page") or {}
    styles = rules.get("styles") or {}

    def style_defaults(key: str) -> Dict[str, Any]:
        style = styles.get(key) or {}
        fonts = {}
        if style.get("ascii_font"):
            fonts["ascii"] = style["ascii_font"]
            fonts["hAnsi"] = style["ascii_font"]
            fonts["cs"] = style["ascii_font"]
        if style.get("east_asia_font"):
            fonts["eastAsia"] = style["east_asia_font"]
        return {
            "fonts": fonts,
            "sizePt": style.get("size_pt"),
            "bold": style.get("bold"),
            "jc": style.get("align"),
            "spacing": {
                "beforePt": style.get("spacing_before_pt"),
                "afterPt": style.get("spacing_after_pt"),
                "linePt": style.get("line_spacing_pt"),
                "lineRule": "exact" if style.get("line_spacing_pt") else None,
            },
            "ind": {"firstLineChars": style.get("indent_chars")},
        }

    return {
        "page": {
            "paper": page.get("paper"),
            "top": {"cm": page.get("margin_top_cm")},
            "bottom": {"cm": page.get("margin_bottom_cm")},
            "left": {"cm": page.get("margin_left_cm")},
            "right": {"cm": page.get("margin_right_cm")},
            "header": {"cm": page.get("header_cm")},
            "footer": {"cm": page.get("footer_cm")},
            "headerText": page.get("header_text"),
        },
        "styles": {
            "body": style_defaults("body_text"),
            "heading1": style_defaults("heading1"),
            "heading2": style_defaults("heading2"),
            "heading3": style_defaults("heading3"),
            "title_cn": style_defaults("thesis_title"),
            "abstract_label_cn": style_defaults("abstract_label_cn"),
            "abstract_text_cn": style_defaults("abstract_text_cn"),
            "keywords_label_cn": style_defaults("keywords_label_cn"),
            "keywords_text_cn": style_defaults("keywords_text_cn"),
            "title_en": style_defaults("title_en"),
            "abstract_label_en": style_defaults("abstract_label_en"),
            "abstract_text_en": style_defaults("abstract_text_en"),
            "keywords_label_en": style_defaults("keywords_label_en"),
            "keywords_text_en": style_defaults("keywords_text_en"),
        },
        "numberingFamilies": rules.get("numbering_families") or {},
    }


def extract_role_style_targets(rules: Dict[str, Any]) -> Dict[str, Any]:
    return (recommended_defaults_from_rules(rules).get("styles") or {}) if rules else {}


def compare_rules_to_baseline(rules: Dict[str, Any], baseline: Dict[str, Any]) -> Dict[str, Any]:
    if not rules:
        return {}
    mismatches = []
    recommended = baseline.get("recommendedDefaults") or {}
    baseline_page = recommended.get("page") or {}
    rule_page = rules.get("page") or {}
    page_map = {
        "margin_top_cm": "top",
        "margin_bottom_cm": "bottom",
        "margin_left_cm": "left",
        "margin_right_cm": "right",
        "header_cm": "header",
        "footer_cm": "footer",
    }
    for rule_key, base_key in page_map.items():
        rule_value = rule_page.get(rule_key)
        baseline_value = ((baseline_page.get(base_key) or {}).get("cm")) if baseline_page.get(base_key) else None
        if rule_value is not None and baseline_value is not None and round(float(rule_value), 3) != round(float(baseline_value), 3):
            mismatches.append(
                {"scope": "page", "field": rule_key, "ruleValue": rule_value, "baselineValue": baseline_value}
            )

    style_expectations = {"body_text": "body", "heading1": "heading1", "heading2": "heading2", "heading3": "heading3"}
    baseline_styles = baseline.get("roleStyleTargets") or {}
    for rule_key, role_key in style_expectations.items():
        rule_style = (rules.get("styles") or {}).get(rule_key) or {}
        baseline_style = baseline_styles.get(role_key) or {}
        baseline_east_asia = (baseline_style.get("fonts") or {}).get("eastAsia")
        if rule_style.get("east_asia_font") and baseline_east_asia and rule_style["east_asia_font"] != baseline_east_asia:
            mismatches.append(
                {
                    "scope": "style",
                    "field": f"{rule_key}.east_asia_font",
                    "ruleValue": rule_style["east_asia_font"],
                    "baselineValue": baseline_east_asia,
                }
            )
    return {"mismatches": mismatches, "match": not mismatches}


def parse_document(
    docx_path: str,
    template_docx: Optional[str] = None,
    rules_path: Optional[str] = None,
) -> Dict[str, Any]:
    rules = load_rules(rules_path)
    with zipfile.ZipFile(docx_path) as zf:
        document_root = read_xml(zf, "word/document.xml")
        styles_root = read_xml(zf, "word/styles.xml")
        numbering_root = read_xml(zf, "word/numbering.xml")
        theme_root = read_xml(zf, "word/theme/theme1.xml")
        settings_root = read_xml(zf, "word/settings.xml")
        font_table_root = read_xml(zf, "word/fontTable.xml")
        footnotes_root = read_xml(zf, "word/footnotes.xml")
        endnotes_root = read_xml(zf, "word/endnotes.xml")

        theme = extract_theme(theme_root)
        doc_defaults = extract_doc_defaults(styles_root)
        styles = extract_styles(styles_root)
        numbering = extract_numbering(numbering_root)
        settings = extract_settings(settings_root)
        font_table = extract_font_table(font_table_root)
        document_rels = parse_relationships(zf, "word/_rels/document.xml.rels")
        header_footer_parts = extract_header_footer_parts(zf)

        paragraphs: List[Dict[str, Any]] = []
        tables: List[Dict[str, Any]] = []
        style_histogram = Counter()
        body_sequence = []
        if document_root is not None:
            body = document_root.find("w:body", NS)
            paragraph_index = 0
            table_index = 0
            for child in list(body) if body is not None else []:
                tag = local_name(child.tag)
                if tag == "p":
                    text = paragraph_text(child)
                    ppr_node = child.find("w:pPr", NS)
                    ppr = parse_ppr(ppr_node)
                    style_id = None
                    if ppr_node is not None and ppr_node.find("w:pStyle", NS) is not None:
                        style_id = ppr_node.find("w:pStyle", NS).attrib.get(qn("val"))
                    if style_id:
                        style_histogram[style_id] += 1
                    numpr = ppr_node.find("w:numPr", NS) if ppr_node is not None else None
                    num_id = None
                    ilvl = None
                    if numpr is not None:
                        if numpr.find("w:numId", NS) is not None:
                            num_id = safe_int(numpr.find("w:numId", NS).attrib.get(qn("val")))
                        if numpr.find("w:ilvl", NS) is not None:
                            ilvl = safe_int(numpr.find("w:ilvl", NS).attrib.get(qn("val")))
                    style_chain = build_style_chain(style_id, styles)
                    numbering_level = get_numbering_level(numbering, str(num_id) if num_id is not None else None, ilvl)
                    effective_paragraph = resolve_effective_ppr(doc_defaults, style_chain, numbering_level, ppr)
                    runs = []
                    for run_index, run in enumerate(iter_text_runs(child)):
                        direct_rpr = parse_rpr(run.find("w:rPr", NS))
                        run_style_id = direct_rpr.get("rStyle")
                        run_style_chain = build_style_chain(run_style_id, styles)
                        text_value = run_text(run)
                        effective_run = resolve_effective_rpr(
                            doc_defaults,
                            style_chain,
                            run_style_chain,
                            numbering_level,
                            direct_rpr,
                            theme,
                            text_value,
                        )
                        runs.append(
                            {
                                "index": run_index,
                                "text": text_value,
                                "styleId": run_style_id,
                                "styleName": get_style_name(run_style_id, styles),
                                "directRPr": direct_rpr,
                                "effectiveFormat": effective_run,
                            }
                        )
                    paragraph = {
                        "index": paragraph_index,
                        "text": text,
                        "styleId": style_id,
                        "styleName": get_style_name(style_id, styles),
                        "styleChain": [style["styleId"] for style in style_chain],
                        "numId": num_id,
                        "ilvl": ilvl,
                        "outlineLvl": effective_paragraph.get("outlineLvl"),
                        "directPPr": ppr,
                        "numberingLevel": numbering_level,
                        "effectiveParagraph": effective_paragraph,
                        "runs": runs,
                        "effectiveRunSummary": dominant_run_summary(runs),
                        "manualNumbering": parse_manual_numbering(text),
                        "hasMath": child.find(".//m:oMath", NS) is not None or child.find(".//m:oMathPara", NS) is not None,
                        "hasDrawing": child.find(".//w:drawing", NS) is not None,
                        "drawingCount": len(child.findall(".//w:drawing", NS)),
                        "hasFootnoteRef": child.find(".//w:footnoteReference", NS) is not None,
                        "hasEndnoteRef": child.find(".//w:endnoteReference", NS) is not None,
                        "breakTypes": [br.attrib.get(qn("type"), "textWrapping") for br in child.findall(".//w:br", NS)],
                        "renderedPageBreaks": len(child.findall(".//w:lastRenderedPageBreak", NS)),
                        "hasSectPr": ppr_node is not None and ppr_node.find("w:sectPr", NS) is not None,
                        "hasTocField": child.find(".//w:fldSimple", NS) is not None
                        and "TOC" in (((child.find(".//w:fldSimple", NS) or ET.Element("x")).attrib.get(qn("instr"))) or ""),
                    }
                    paragraph["role"] = classify_paragraph_role(paragraph)
                    paragraphs.append(paragraph)
                    body_sequence.append(
                        {
                            "kind": "paragraph",
                            "paragraphIndex": paragraph_index,
                            "role": "image_block"
                            if paragraph.get("hasDrawing") and not text.strip()
                            else (paragraph.get("role") or {}).get("role"),
                            "text": text,
                            "hasDrawing": paragraph.get("hasDrawing"),
                            "drawingCount": paragraph.get("drawingCount"),
                        }
                    )
                    paragraph_index += 1
                elif tag == "tbl":
                    tables.append({"index": table_index, "textPreview": paragraph_text(child)[:200]})
                    body_sequence.append({"kind": "table", "tableIndex": table_index, "textPreview": paragraph_text(child)[:120]})
                    table_index += 1

        sections = parse_sections(document_root, document_rels, header_footer_parts)
        numbering_analysis = analyze_numbering_tree(paragraphs)
        style_analysis = analyze_styles(styles)
        font_slot_analysis = analyze_font_slots(paragraphs)
        immutable_prefix = detect_immutable_prefix_region(paragraphs)
        caption_layout = analyze_caption_layout(body_sequence)
        front_matter_analysis = analyze_front_matter_structure(paragraphs)
        cross_reference_analysis = analyze_cross_reference_candidates(paragraphs, caption_layout.get("captions") or [])

        issues = []
        issues.extend(numbering_analysis.get("issues") or [])
        if style_analysis.get("builtinHeadingRisks"):
            issues.append({"kind": "builtin_heading_risks", "count": len(style_analysis["builtinHeadingRisks"])})
        if font_slot_analysis.get("suspiciousRuns"):
            issues.append({"kind": "font_slot_risks", "count": len(font_slot_analysis["suspiciousRuns"])})
        if caption_layout.get("issues"):
            issues.append({"kind": "caption_layout_issues", "count": len(caption_layout["issues"])})
        if front_matter_analysis.get("issues"):
            issues.append({"kind": "front_matter_structure_issues", "count": len(front_matter_analysis["issues"])})
        if cross_reference_analysis.get("unresolved"):
            issues.append({"kind": "unresolved_cross_references", "count": len(cross_reference_analysis["unresolved"])})

        template_baseline = None
        rules_comparison = {}
        if template_docx:
            template_baseline = extract_template_baseline(template_docx, rules_path=rules_path)
            rules_comparison = compare_rules_to_baseline(rules, template_baseline)

        return {
            "generatedAt": now_iso(),
            "source": str(Path(docx_path).resolve()),
            "templateDocx": str(Path(template_docx).resolve()) if template_docx else None,
            "rulesPath": str(Path(rules_path).resolve()) if rules_path else None,
            "summary": {
                "paragraphCount": len(paragraphs),
                "tableCount": len(tables),
                "sectionCount": len(sections),
                "styleCount": len(styles),
                "numberingInstanceCount": len(numbering.get("nums") or {}),
                "numberingAbstractCount": len(numbering.get("abstractNums") or {}),
                "footnoteCount": len(footnotes_root.findall(".//w:footnote", NS)) if footnotes_root is not None else 0,
                "endnoteCount": len(endnotes_root.findall(".//w:endnote", NS)) if endnotes_root is not None else 0,
            },
            "docDefaults": doc_defaults,
            "theme": theme,
            "settings": settings,
            "fontTable": font_table,
            "styles": styles,
            "styleAnalysis": style_analysis,
            "numbering": numbering,
            "sections": sections,
            "headerFooterParts": header_footer_parts,
            "styleHistogram": dict(style_histogram),
            "paragraphs": paragraphs,
            "tables": tables,
            "bodySequence": body_sequence,
            "numberingAnalysis": numbering_analysis,
            "fontSlotAnalysis": font_slot_analysis,
            "captionLayout": caption_layout,
            "frontMatterAnalysis": front_matter_analysis,
            "crossReferenceAnalysis": cross_reference_analysis,
            "issues": issues,
            "rulesComparison": rules_comparison,
            "recommendedDefaults": recommended_defaults_from_rules(rules),
            "roleStyleTargets": extract_role_style_targets(rules),
            "templateBaseline": template_baseline,
            "preservationHints": {
                "immutablePrefix": immutable_prefix,
                "preferredHeaderText": preferred_header_text(template_baseline, rules),
            },
        }


def extract_template_baseline(docx_path: str, rules_path: Optional[str] = None) -> Dict[str, Any]:
    document = parse_document(docx_path, rules_path=rules_path)
    baseline = {
        "generatedAt": document["generatedAt"],
        "source": document["source"],
        "summary": document.get("summary"),
        "docDefaults": document["docDefaults"],
        "theme": document["theme"],
        "settings": document["settings"],
        "fontTable": {
            "fontCount": len(document.get("fontTable", {}).get("fonts") or []),
            "names": document.get("fontTable", {}).get("names") or [],
        },
        "styles": document["styles"],
        "styleAnalysis": document["styleAnalysis"],
        "numbering": document["numbering"],
        "sections": document["sections"],
        "headerFooterParts": document["headerFooterParts"],
        "recommendedDefaults": document["recommendedDefaults"],
        "roleStyleTargets": document["roleStyleTargets"],
        "templateStructure": {
            "firstParagraphs": [
                {
                    "index": paragraph["index"],
                    "text": paragraph["text"][:120],
                    "styleId": paragraph["styleId"],
                    "styleName": paragraph["styleName"],
                    "role": paragraph["role"],
                }
                for paragraph in document["paragraphs"][:80]
                if paragraph["text"].strip()
            ],
            "headingCandidates": [
                {
                    "index": item["paragraphIndex"],
                    "level": item["level"],
                    "text": item["text"],
                    "family": item["family"],
                }
                for item in document["numberingAnalysis"]["headingTree"][:80]
            ],
        },
        "preservationHints": document.get("preservationHints") or {},
        "frontMatterAnalysis": document.get("frontMatterAnalysis") or {},
        "paragraphProbe": [
            {"index": paragraph["index"], "text": paragraph["text"][:200]}
            for paragraph in document.get("paragraphs", [])[:120]
        ],
    }
    if rules_path:
        baseline["rulesComparison"] = compare_rules_to_baseline(load_rules(rules_path), baseline)
    return baseline


SAFE_SUBSET_ACTION_TYPES = {
    "apply_paragraph_style",
    "apply_caption_style",
    "normalize_heading_prefix_separator",
    "normalize_header_text",
    "normalize_bibliography_indent",
    "preserve_or_insert_toc_field",
    "manual_review",
}

AUTO_EXECUTABLE_ACTION_TYPES = {
    "apply_paragraph_style",
    "apply_caption_style",
    "normalize_heading_prefix_separator",
    "normalize_header_text",
    "normalize_bibliography_indent",
    "preserve_or_insert_toc_field",
}

CONFIRMATION_FIRST_ACTION_TYPES = {
    "rebuild_front_matter",
    "rebuild_numbering_system",
    "convert_plain_text_cross_reference_to_field",
    "rewrite_equation_numbering",
}

ACTION_TYPE_WHITELIST = SAFE_SUBSET_ACTION_TYPES | CONFIRMATION_FIRST_ACTION_TYPES


def format_plan_target(paragraph_index: Optional[int]) -> Optional[str]:
    if paragraph_index is None:
        return None
    return f"p_{int(paragraph_index):04d}"


def infer_plan_document_risk_class(audit: Dict[str, Any]) -> str:
    numbering_analysis = audit.get("numberingAnalysis") or {}
    high_risk = False
    if len(numbering_analysis.get("anomalies") or []) >= 6:
        high_risk = True
    if len(numbering_analysis.get("mixedManualAndAuto") or []) >= 1:
        high_risk = True
    if len(((audit.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or [])) >= 25:
        high_risk = True
    if len((audit.get("sections") or [])) >= 6:
        high_risk = True
    if len(((audit.get("crossReferenceAnalysis") or {}).get("unresolved") or [])) >= 8:
        high_risk = True
    if high_risk:
        return "C"
    if audit.get("templateDocx") and ((audit.get("rulesComparison") or {}).get("match", False)):
        return "A"
    return "B"


def infer_plan_recommended_mode(audit: Dict[str, Any]) -> str:
    risk_class = audit.get("documentRiskClass") or infer_plan_document_risk_class(audit)
    return "audit-only" if risk_class == "C" else "conservative-repair"


def action_role_name(role_name: Optional[str]) -> Optional[str]:
    if not role_name:
        return None
    if role_name.startswith("heading"):
        return role_name.replace("heading", "heading_")
    return role_name


def style_action_risk(role_name: Optional[str], confidence: float) -> str:
    if role_name in {"heading1", "heading2", "heading3", "heading4", "figure_caption", "table_caption"} and confidence >= 0.85:
        return "low"
    if role_name in {"title_cn", "title_en", "abstract_paragraph_cn", "abstract_paragraph_en", "keywords_cn", "keywords_en", "toc_heading", "toc_field", "toc_entry"}:
        return "medium"
    if role_name == "body_enumeration":
        return "medium"
    return "low" if confidence >= 0.85 else "medium"


def style_action_type(role_name: Optional[str]) -> str:
    if role_name in {"figure_caption", "table_caption"}:
        return "apply_caption_style"
    return "apply_paragraph_style"


def build_gate_result(
    action: Dict[str, Any],
    *,
    paragraphs: List[Dict[str, Any]],
    preserve_regions: List[Dict[str, Any]],
    effective_mode: str,
    recommended_mode: str,
    document_risk_class: str,
) -> Dict[str, Any]:
    blocked_by: List[str] = []
    route = "execute"
    confidence_threshold = 0.85 if action.get("risk") == "low" else 0.95
    action_type = action.get("type")
    paragraph_index = safe_int(action.get("paragraphIndex"))
    target = action.get("target")
    if paragraph_index is not None:
        target_exists = 0 <= paragraph_index < len(paragraphs)
    elif action_type == "normalize_header_text":
        target_exists = bool(target)
    elif action_type == "rebuild_front_matter":
        target_exists = bool(action.get("target_region"))
    else:
        target_exists = bool(target)
    inside_preserved_region = False

    if paragraph_index is not None:
        for region in preserve_regions:
            start = safe_int(region.get("startParagraph")) or 0
            end = safe_int(region.get("endParagraphExclusive"))
            if end is not None and start <= paragraph_index < end:
                inside_preserved_region = True
                break

    if not action_type or action_type not in ACTION_TYPE_WHITELIST:
        blocked_by.append("shape_or_whitelist")
    required_fields = {"type", "target", "confidence", "risk", "reason"}
    missing_fields = sorted(field for field in required_fields if action.get(field) in (None, ""))
    if missing_fields:
        blocked_by.append("missing_required_fields")
    if not target or not target_exists:
        blocked_by.append("target_missing")
    if inside_preserved_region:
        blocked_by.append("preserved_region")

    confidence = float(action.get("confidence") or 0.0)
    if action_type != "manual_review" and confidence < confidence_threshold:
        blocked_by.append("confidence_threshold")

    if action_type not in SAFE_SUBSET_ACTION_TYPES:
        blocked_by.append("outside_safe_subset")
    if action_type in CONFIRMATION_FIRST_ACTION_TYPES or action.get("requires_confirmation"):
        blocked_by.append("confirmation_first")

    if effective_mode == "audit-only" and action_type != "manual_review":
        blocked_by.append("mode_audit_only")
    elif effective_mode == "conservative-repair":
        if action.get("risk") != "low" and action_type != "manual_review":
            blocked_by.append("risk_not_allowed_in_conservative_mode")
        if action_type not in AUTO_EXECUTABLE_ACTION_TYPES and action_type != "manual_review":
            blocked_by.append("not_auto_executable")
    elif effective_mode == "rebuild":
        if action.get("risk") == "high" and action_type != "manual_review":
            blocked_by.append("high_risk_blocked_in_rebuild_mode")

    modifies_text = bool(action.get("modifiesText"))
    if modifies_text:
        blocked_by.append("non_destructive_text_protection")

    if blocked_by:
        if "confirmation_first" in blocked_by or "non_destructive_text_protection" in blocked_by:
            route = "confirmationRequests"
        else:
            route = "manualReview"

    return {
        "passed": not blocked_by,
        "route": route,
        "blockedBy": blocked_by,
        "effectiveMode": effective_mode,
        "recommendedMode": recommended_mode,
        "documentRiskClass": document_risk_class,
        "targetExists": target_exists,
        "insidePreservedRegion": inside_preserved_region,
        "confidenceThreshold": confidence_threshold,
        "safeSubset": action_type in SAFE_SUBSET_ACTION_TYPES,
        "requiresConfirmation": action_type in CONFIRMATION_FIRST_ACTION_TYPES or bool(action.get("requires_confirmation")),
        "modifiesText": modifies_text,
    }


def make_confirmation_request_from_action(action: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "decisionId": f"confirm_{action.get('id')}",
        "priority": action.get("priority", 100),
        "requiresConfirmation": True,
        "scope": action.get("type"),
        "question": action.get("confirmationPrompt") or f"是否允许执行动作 {action.get('type')} 到 {action.get('target')}？",
        "reason": action.get("reason"),
        "impact": action.get("impact") or [],
        "recommendedDefault": "ask_user_first",
        "target": action.get("target"),
        "actionId": action.get("id"),
    }


def make_manual_review_from_action(action: Dict[str, Any], blocked_by: List[str]) -> Dict[str, Any]:
    return {
        "paragraphIndex": action.get("paragraphIndex"),
        "issue": "schema_gate_blocked_action",
        "actionId": action.get("id"),
        "actionType": action.get("type"),
        "target": action.get("target"),
        "blockedBy": blocked_by,
        "reason": action.get("reason"),
    }


def build_repair_plan(
    audit: Dict[str, Any],
    *,
    style_map_path: Optional[str] = None,
    front_matter_policy_path: Optional[str] = None,
) -> Dict[str, Any]:
    paragraphs = audit.get("paragraphs") or []
    numbering_analysis = audit.get("numberingAnalysis") or {}
    rules = load_rules(audit.get("rulesPath")) if audit.get("rulesPath") else {}
    style_map = load_rules(style_map_path) if style_map_path else {}
    front_matter_policy = load_rules(front_matter_policy_path) if front_matter_policy_path else {}
    style_targets = resolve_profile_style_targets(style_map)
    immutable_prefix = ((audit.get("preservationHints") or {}).get("immutablePrefix")) or None
    template_prefix = (((audit.get("templateBaseline") or {}).get("preservationHints") or {}).get("immutablePrefix")) or None
    template_prefix_count = detect_template_prefix_count(audit.get("templateBaseline"))
    template_prefix_end = template_prefix_count or (template_prefix["endParagraphExclusive"] if template_prefix else 0)
    preserve_template_prefix = bool(front_matter_policy_value(front_matter_policy, "preserve_cover_and_integrity_pages_from_template", True))
    source_prefix_drop_count = immutable_prefix["endParagraphExclusive"] if immutable_prefix else 0
    editable_start = 0 if (preserve_template_prefix and template_prefix_end > 0) else source_prefix_drop_count
    document_risk_class = audit.get("documentRiskClass") or infer_plan_document_risk_class(audit)
    recommended_mode = audit.get("recommendedMode") or infer_plan_recommended_mode(audit)
    effective_mode = "audit-only" if document_risk_class == "C" else recommended_mode

    style_mapping = {}
    paragraph_actions = []
    manual_review = []
    confirmation_requests = list(audit.get("confirmationRequests") or [])
    rebuild_regions = []
    role_map = {}
    numbering_actions = list(numbering_analysis.get("repairActions") or [])
    numbering_issue_kinds = {issue.get("kind") for issue in (numbering_analysis.get("issues") or [])}
    toc_entry_count = sum(1 for paragraph in paragraphs if (paragraph.get("role") or {}).get("role") == "toc_entry")

    first_heading_index = None
    abstract_context: Optional[str] = None
    for paragraph in paragraphs:
        if paragraph["index"] < editable_start:
            role_map[str(paragraph["index"])] = {"role": "preserved_prefix", "confidence": 1.0, "signals": ["immutable_prefix_region"]}
            continue
        role = paragraph.get("role") or {}
        role_name = role.get("role")
        confidence = role.get("confidence", 0.0)
        text = (paragraph.get("text") or "").strip()
        manual_numbering = paragraph.get("manualNumbering")
        role_map[str(paragraph["index"])] = role
        if role_name == "heading1" and first_heading_index is None and not manual_numbering:
            role = {"role": "title_cn", "confidence": 0.97, "signals": ["first_unnumbered_heading_title"]}
            role_name = "title_cn"
            paragraph["role"] = role
            role_map[str(paragraph["index"])] = role
        if role_name and role_name.startswith("heading") and first_heading_index is None:
            first_heading_index = paragraph["index"]
        if role_name == "title_cn":
            abstract_context = "cn"
            style_mapping[str(paragraph["index"])] = style_targets["title_cn"]
        elif role_name == "title_en":
            abstract_context = "en"
            style_mapping[str(paragraph["index"])] = style_targets["title_en"]
        elif role_name in {"abstract_heading_cn"} or text in {"中文摘要", "摘要", "摘  要"}:
            abstract_context = "cn"
            style_mapping[str(paragraph["index"])] = style_targets["abstract_heading"]
            paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_abstract_heading", "targetStyle": style_targets["abstract_heading"]})
        elif role_name in {"abstract_heading_en"} or text in {"English Abstract", "ABSTRACT"}:
            abstract_context = "en"
            style_mapping[str(paragraph["index"])] = style_targets["abstract_heading"]
            paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_abstract_heading", "targetStyle": style_targets["abstract_heading"]})
        elif role_name == "abstract_paragraph_cn":
            abstract_context = "cn"
            style_mapping[str(paragraph["index"])] = style_targets["abstract_text_cn"]
        elif role_name == "abstract_paragraph_en":
            abstract_context = "en"
            style_mapping[str(paragraph["index"])] = style_targets["abstract_text_en"]
        elif role_name == "keywords_cn":
            style_mapping[str(paragraph["index"])] = style_targets["keywords_cn"]
            abstract_context = None
        elif role_name == "keywords_en":
            style_mapping[str(paragraph["index"])] = style_targets["keywords_en"]
            abstract_context = None
        elif abstract_context == "cn" and role_name in {"body", "body_enumeration"} and text:
            style_mapping[str(paragraph["index"])] = style_targets["abstract_text_cn"]
        elif abstract_context == "en" and role_name in {"body", "body_enumeration"} and text:
            style_mapping[str(paragraph["index"])] = style_targets["abstract_text_en"]
        elif role_name == "toc_heading":
            abstract_context = None
            style_mapping[str(paragraph["index"])] = style_targets["toc_title"]
            paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_toc_heading", "targetStyle": style_targets["toc_title"]})
        elif role_name in {"toc_field", "toc_entry"}:
            abstract_context = None
            style_mapping[str(paragraph["index"])] = style_targets["toc_body"]
        elif role_name in {"heading1", "references_heading", "ack_heading", "appendix_heading"}:
            abstract_context = None
            style_mapping[str(paragraph["index"])] = style_targets["heading1"]
            paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_heading_style", "targetStyle": style_targets["heading1"]})
        elif role_name == "heading2":
            abstract_context = None
            style_mapping[str(paragraph["index"])] = style_targets["heading2"]
            paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_heading_style", "targetStyle": style_targets["heading2"]})
        elif role_name in {"heading3", "heading4"}:
            abstract_context = None
            style_mapping[str(paragraph["index"])] = style_targets["heading3"]
            paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_heading_style", "targetStyle": style_targets["heading3"]})
        elif role_name in {"body", "body_enumeration"} and paragraph["text"].strip():
            if role_name == "body_enumeration" and (paragraph.get("manualNumbering") or {}).get("family") in {"humanities_cn", "humanities_chapter"}:
                level = (paragraph.get("manualNumbering") or {}).get("level") or 1
                target_key = f"heading{min(max(level, 1), 3)}"
                style_mapping[str(paragraph["index"])] = style_targets[target_key]
                paragraph_actions.append({"paragraphIndex": paragraph["index"], "action": "normalize_heading_style", "targetStyle": style_targets[target_key]})
            else:
                style_mapping[str(paragraph["index"])] = style_targets["body_text"]
        elif role_name == "image_block":
            style_mapping[str(paragraph["index"])] = style_targets["caption"]
        elif role_name in {"figure_caption", "table_caption"}:
            style_mapping[str(paragraph["index"])] = style_targets["caption"]

        if confidence < 0.72 and role_name not in {"blank", "body"}:
            manual_review.append(
                {
                    "paragraphIndex": paragraph["index"],
                    "issue": "low_confidence_role",
                    "role": role_name,
                    "confidence": confidence,
                    "text": paragraph["text"][:140],
                }
            )
        if role_name == "body_enumeration":
            manual_review.append(
                {
                    "paragraphIndex": paragraph["index"],
                    "issue": "body_enumeration_may_be_heading",
                    "text": paragraph["text"][:140],
                    "confidence": confidence,
                }
            )
        if paragraph.get("hasMath"):
            manual_review.append(
                {
                    "paragraphIndex": paragraph["index"],
                    "issue": "equation_spacing_and_numbering_review",
                    "text": paragraph["text"][:140],
                }
            )

    for issue in numbering_analysis.get("mixedManualAndAuto") or []:
        manual_review.append(
            {
                "paragraphIndex": issue["paragraphIndex"],
                "issue": "manual_and_auto_numbering_mixed",
                "text": issue["text"],
            }
        )

    front_matter_end = first_heading_index if first_heading_index is not None else min(len(paragraphs), 20)
    front_matter_paragraphs = paragraphs[:front_matter_end]
    toc_like_count = sum(1 for item in front_matter_paragraphs if (item.get("role") or {}).get("role") == "toc_entry")
    if toc_like_count >= 3:
        rebuild_regions.append(
            {
                "startParagraph": 0,
                "endParagraph": front_matter_end,
                "mode": "template_rebuild_candidate",
                "reason": "front_matter_contains_toc_like_or_template_scaffold_blocks",
            }
        )

    caption_layout = audit.get("captionLayout") or {}
    for issue in caption_layout.get("issues") or []:
        manual_review.append(
            {
                "paragraphIndex": issue.get("paragraphIndex"),
                "issue": issue.get("kind"),
                "text": issue.get("text"),
            }
        )

    front_matter_analysis = audit.get("frontMatterAnalysis") or {}
    for issue in front_matter_analysis.get("issues") or []:
        manual_review.append(
            {
                "paragraphIndex": issue.get("abstractParagraph") or issue.get("tocParagraph"),
                "issue": issue.get("kind"),
                "details": issue,
            }
        )

    cross_reference_analysis = audit.get("crossReferenceAnalysis") or {}
    for item in (cross_reference_analysis.get("unresolved") or [])[:40]:
        manual_review.append(
            {
                "paragraphIndex": item.get("paragraphIndex"),
                "issue": "unresolved_cross_reference",
                "reference": item.get("reference"),
                "text": item.get("text"),
            }
        )

    if "mixed_numbering_families" in numbering_issue_kinds or toc_entry_count >= 8:
        retained_numbering_actions = [
            action for action in numbering_actions if action.get("subtype") == "separator_only"
        ]
        suppressed_actions = [action for action in numbering_actions if action not in retained_numbering_actions]
        if suppressed_actions:
            manual_review.append(
                {
                    "issue": "numbering_repairs_suppressed_due_to_mixed_families_or_template_like_structure",
                    "count": len(suppressed_actions),
                }
            )
        numbering_actions = retained_numbering_actions
    elif len(numbering_actions) > 12:
        retained_numbering_actions = [
            action for action in numbering_actions if action.get("subtype") == "separator_only"
        ]
        suppressed_actions = [action for action in numbering_actions if action not in retained_numbering_actions]
        if suppressed_actions:
            manual_review.append(
                {
                    "issue": "numbering_repairs_suppressed_due_to_high_change_volume",
                    "count": len(suppressed_actions),
                }
            )
        numbering_actions = retained_numbering_actions

    repair_mode = "hybrid" if rebuild_regions else "in_place"
    preserve_regions = []
    if not (preserve_template_prefix and template_prefix_end > 0) and immutable_prefix:
        preserve_regions = [immutable_prefix]
    if not preserve_regions and preserve_template_prefix and template_prefix_end > 0:
        preserve_regions = [
            {
                "startParagraph": 0,
                "endParagraphExclusive": template_prefix_end,
                "reason": "template_front_matter_prefix_from_profile_policy",
                "confidence": 0.98,
            }
        ]
    header_text = None
    for paragraph in paragraphs:
        if (paragraph.get("role") or {}).get("role") == "title_cn":
            candidate = (paragraph.get("text") or "").strip()
            if candidate:
                header_text = candidate
                break
    if not header_text:
        header_text = ((audit.get("preservationHints") or {}).get("preferredHeaderText"))
    blocked_auto_repairs = list(audit.get("blockedAutoRepairs") or [])

    actions: List[Dict[str, Any]] = []
    for paragraph_index_str, style_id in style_mapping.items():
        paragraph_index = safe_int(paragraph_index_str)
        if paragraph_index is None or not (0 <= paragraph_index < len(paragraphs)):
            continue
        role = role_map.get(paragraph_index_str) or {}
        role_name = role.get("role")
        confidence = float(role.get("confidence") or 0.0)
        action = {
            "id": f"style_{paragraph_index:04d}",
            "type": style_action_type(role_name),
            "target": format_plan_target(paragraph_index),
            "paragraphIndex": paragraph_index,
            "role": action_role_name(role_name),
            "style_id": style_id,
            "confidence": round(confidence, 2),
            "risk": style_action_risk(role_name, confidence),
            "reason": f"Paragraph role {role_name or 'unknown'} maps to style {style_id} under the active thesis profile.",
            "legacySource": "styleMapping",
            "modifiesText": False,
            "priority": 20 if role_name and role_name.startswith("heading") else 40,
        }
        actions.append(action)

    for action in numbering_actions:
        paragraph_index = safe_int(action.get("paragraphIndex"))
        target = format_plan_target(paragraph_index)
        subtype = action.get("subtype")
        action_type = "normalize_heading_prefix_separator" if subtype == "separator_only" else "rebuild_numbering_system"
        normalized = {
            "id": f"numbering_{(paragraph_index if paragraph_index is not None else 0):04d}",
            "type": action_type,
            "target": target,
            "paragraphIndex": paragraph_index,
            "confidence": round(float(action.get("confidence") or 0.0), 2),
            "risk": "low" if subtype == "separator_only" else "high",
            "reason": "Heading numbering drift was detected and a deterministic prefix normalization candidate was proposed.",
            "legacySource": "numberingActions",
            "expected_before": action.get("oldPrefix"),
            "expected_after": action.get("newPrefix"),
            "modifiesText": False,
            "priority": 10,
            "requires_confirmation": action_type in CONFIRMATION_FIRST_ACTION_TYPES,
            "impact": ["May alter visible heading numbering prefixes."],
            "confirmationPrompt": "检测到标题编号需要批量改写，是否允许执行编号修复？" if action_type in CONFIRMATION_FIRST_ACTION_TYPES else None,
        }
        actions.append(normalized)

    if header_text:
        actions.append(
            {
                "id": "header_text_0000",
                "type": "normalize_header_text",
                "target": "header_default",
                "confidence": 0.93,
                "risk": "low",
                "reason": "Known thesis template header text should match the detected Chinese thesis title.",
                "legacySource": "headerText",
                "header_text": header_text,
                "modifiesText": False,
                "priority": 60,
            }
        )

    for region_index, region in enumerate(rebuild_regions):
        rebuild_requires_confirmation = bool(front_matter_policy_value(front_matter_policy, "front_matter_rebuild_requires_confirmation", True))
        rebuild_action = {
            "id": f"rebuild_{region_index:04d}",
            "type": "rebuild_front_matter",
            "target": f"region_front_matter_{region_index}",
            "target_region": region,
            "confidence": 0.81,
            "risk": "high",
            "reason": "Front matter resembles template scaffold or TOC-heavy prefix and is a candidate for template-based rebuild.",
            "legacySource": "rebuildRegions",
            "requires_confirmation": rebuild_requires_confirmation,
            "modifiesText": False,
            "priority": 5,
            "impact": [
                "May replace the front-matter structure with template-based content regions.",
                "May shift paragraph positions in the first pages of the thesis.",
            ],
            "confirmationPrompt": "检测到前置部分可能需要按模板骨架重建，是否允许继续？" if rebuild_requires_confirmation else None,
        }
        actions.append(rebuild_action)

    preserve_regions = []
    if not (preserve_template_prefix and template_prefix_end > 0) and immutable_prefix:
        preserve_regions = [immutable_prefix]
    if not preserve_regions and preserve_template_prefix and template_prefix_end > 0:
        preserve_regions = [
            {
                "startParagraph": 0,
                "endParagraphExclusive": template_prefix_end,
                "reason": "template_front_matter_prefix_from_profile_policy",
                "confidence": 0.98,
            }
        ]

    gated_actions: List[Dict[str, Any]] = []
    action_counts = {"passed": 0, "blocked": 0, "confirmation": 0}
    for action in actions:
        gate = build_gate_result(
            action,
            paragraphs=paragraphs,
            preserve_regions=preserve_regions,
            effective_mode=effective_mode,
            recommended_mode=recommended_mode,
            document_risk_class=document_risk_class,
        )
        gated = dict(action)
        gated["schemaGate"] = gate
        gated_actions.append(gated)
        if gate["passed"]:
            action_counts["passed"] += 1
        else:
            action_counts["blocked"] += 1
            if gate["route"] == "confirmationRequests":
                action_counts["confirmation"] += 1
                confirmation_requests.append(make_confirmation_request_from_action(gated))
            else:
                manual_review.append(make_manual_review_from_action(gated, gate["blockedBy"]))
            if gate["blockedBy"]:
                for reason in gate["blockedBy"]:
                    entry = f"{action.get('type')}:{reason}"
                    if entry not in blocked_auto_repairs:
                        blocked_auto_repairs.append(entry)

    legacy_confirmation_ids = {item.get("decisionId") for item in confirmation_requests if item.get("decisionId")}
    deduped_confirmation_requests: List[Dict[str, Any]] = []
    for item in confirmation_requests:
        decision_id = item.get("decisionId")
        if decision_id and decision_id in legacy_confirmation_ids:
            legacy_confirmation_ids.discard(decision_id)
            deduped_confirmation_requests.append(item)
        elif not decision_id:
            deduped_confirmation_requests.append(item)

    return {
        "generatedAt": now_iso(),
        "source": audit.get("source"),
        "templateDocx": audit.get("templateDocx"),
        "rulesPath": audit.get("rulesPath"),
        "preset": "zafu_2022",
        "styleMapPath": str(Path(style_map_path).resolve()) if style_map_path else None,
        "frontMatterPolicyPath": str(Path(front_matter_policy_path).resolve()) if front_matter_policy_path else None,
        "profileStyleTargets": style_targets,
        "frontMatterPolicy": (front_matter_policy.get("policy") or {}) if isinstance(front_matter_policy, dict) else {},
        "repairMode": repair_mode,
        "numberingScheme": numbering_analysis.get("dominantFamily") or "science_decimal",
        "headerText": header_text,
        "documentRiskClass": document_risk_class,
        "recommendedMode": recommended_mode,
        "editableStartParagraph": editable_start,
        "sourceFrontMatterDropParagraphCount": source_prefix_drop_count,
        "templatePrefixParagraphCount": template_prefix_end,
        "useTemplateFrontMatter": bool(preserve_template_prefix and template_prefix_end),
        "preserveRegions": preserve_regions,
        "styleMapping": style_mapping,
        "paragraphRoles": role_map,
        "paragraphActions": paragraph_actions,
        "numberingActions": numbering_actions,
        "actions": gated_actions,
        "actionWhitelist": sorted(ACTION_TYPE_WHITELIST),
        "safeSubsetActionTypes": sorted(SAFE_SUBSET_ACTION_TYPES),
        "confirmationFirstActionTypes": sorted(CONFIRMATION_FIRST_ACTION_TYPES),
        "schemaGate": {
            "enabled": True,
            "effectiveMode": effective_mode,
            "recommendedMode": recommended_mode,
            "documentRiskClass": document_risk_class,
            "confidenceThresholds": {
                "lowRiskAutomatic": 0.85,
                "mediumRiskAutomatic": None,
                "highRiskAutomatic": None,
            },
            "summary": action_counts,
        },
        "blockedAutoRepairs": blocked_auto_repairs,
        "rebuildRegions": rebuild_regions,
        "manualReview": manual_review,
        "confirmationRequests": deduped_confirmation_requests,
        "issues": audit.get("issues") or [],
        "templateRuleDiff": audit.get("rulesComparison") or {},
        "explanations": [
            "AI judgement is limited to block classification and repair planning.",
            "Final Word changes should be written through deterministic OOXML edits and logged.",
            "Every action in actions[] has passed through a schema gate for whitelist, target, confidence, risk, and safe-subset checks.",
            "Low-confidence heading, numbering, equation, and front-matter blocks remain in manual review or confirmationRequests.",
        ],
        "recommendedStyleTargets": audit.get("roleStyleTargets") or {},
    }


def summarize_audit(audit: Dict[str, Any]) -> Dict[str, Any]:
    paragraphs = audit.get("paragraphs") or []
    role_counts = Counter((paragraph.get("role") or {}).get("role") for paragraph in paragraphs if paragraph.get("role"))
    heading_counts = Counter()
    for paragraph in paragraphs:
        role_name = (paragraph.get("role") or {}).get("role")
        if role_name and role_name.startswith("heading"):
            heading_counts[role_name] += 1
    return {
        "paragraphCount": len(paragraphs),
        "roleCounts": dict(role_counts),
        "headingCounts": dict(heading_counts),
        "issues": audit.get("issues") or [],
        "dominantNumberingFamily": ((audit.get("numberingAnalysis") or {}).get("dominantFamily")),
    }


def diff_audits(before: Dict[str, Any], after: Dict[str, Any], plan: Optional[Dict[str, Any]] = None, execution_log: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    before_paragraphs = before.get("paragraphs") or []
    after_paragraphs = after.get("paragraphs") or []
    changed_texts = []
    for before_item, after_item in zip(before_paragraphs, after_paragraphs):
        if before_item.get("text") != after_item.get("text"):
            changed_texts.append(
                {
                    "paragraphIndex": before_item["index"],
                    "before": before_item.get("text", "")[:160],
                    "after": after_item.get("text", "")[:160],
                }
            )
    return {
        "generatedAt": now_iso(),
        "sourceBefore": before.get("source"),
        "sourceAfter": after.get("source"),
        "textChangeSummary": {"changedParagraphCount": len(changed_texts), "changes": changed_texts[:100]},
        "paragraphCountDelta": len(after_paragraphs) - len(before_paragraphs),
        "headingCountBefore": summarize_audit(before).get("headingCounts"),
        "headingCountAfter": summarize_audit(after).get("headingCounts"),
        "numberingRepairSummary": execution_log or [],
        "headerFooterSummary": {"before": before.get("headerFooterParts"), "after": after.get("headerFooterParts")},
        "fontSpacingSummary": {
            "beforeSuspiciousRuns": len(((before.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or [])),
            "afterSuspiciousRuns": len(((after.get("fontSlotAnalysis") or {}).get("suspiciousRuns") or [])),
        },
        "manualReview": (plan or {}).get("manualReview") or [],
    }


def write_json(data: Dict[str, Any], path: Optional[str]) -> None:
    text = json.dumps(data, ensure_ascii=False, indent=2)
    if path:
        Path(path).write_text(text, encoding="utf-8")
    else:
        print(text)



