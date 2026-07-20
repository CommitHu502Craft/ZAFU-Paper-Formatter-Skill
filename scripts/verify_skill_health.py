"""Verify that the Paper-Formatter-Skill is correctly configured."""
from __future__ import annotations

import sys
from pathlib import Path


def check_file(path: Path, description: str) -> bool:
    if path.exists():
        print(f"  [OK] {description}: {path}")
        return True
    else:
        print(f"  [FAIL] {description}: {path}")
        return False


def check_script_content(path: Path, pattern: str, description: str) -> bool:
    if not path.exists():
        print(f"  [FAIL] {description}: {path} not found")
        return False
    content = path.read_text(encoding="utf-8")
    if pattern in content:
        print(f"  [OK] {description}")
        return True
    else:
        print(f"  [FAIL] {description}: pattern '{pattern}' not found")
        return False


def main():
    print("=== Paper-Formatter-Skill Health Check ===")
    print()

    skill_dir = Path(__file__).parent.parent
    all_ok = True

    # Check core scripts
    print("1. Core scripts:")
    all_ok &= check_file(skill_dir / "scripts" / "build_docx_from_markdown.py", "Build script")
    all_ok &= check_file(skill_dir / "scripts" / "apply_ooxml_fixes.py", "Repair script")
    all_ok &= check_file(skill_dir / "scripts" / "docx_ooxml.py", "OOXML engine")

    # Check template
    print()
    print("2. Template:")
    all_ok &= check_file(skill_dir / "assets" / "zafu_front_matter_template.docx", "Template DOCX")

    # Check key fixes in apply_ooxml_fixes.py
    print()
    print("3. Key fixes in apply_ooxml_fixes.py:")
    fix_script = skill_dir / "scripts" / "apply_ooxml_fixes.py"
    all_ok &= check_script_content(fix_script, "remove_vanish_attributes", "Vanish removal fix")
    all_ok &= check_script_content(fix_script, 'qn("pgMar")', "pgMar fix (not qn('w:pgMar'))")
    all_ok &= check_script_content(fix_script, "word/media/", "Media file preservation")

    # Check key fixes in build_docx_from_markdown.py
    print()
    print("4. Key fixes in build_docx_from_markdown.py:")
    build_script = skill_dir / "scripts" / "build_docx_from_markdown.py"
    all_ok &= check_script_content(build_script, "convert_svg_to_png", "SVG conversion")
    all_ok &= check_script_content(build_script, "Skip markdown header", "Markdown header skipping")
    all_ok &= check_script_content(build_script, "ABSTRACT_CN_HEADINGS", "Abstract heading detection")
    all_ok &= check_script_content(build_script, "found_keywords_heading", "Keywords extraction")

    # Check published contracts
    print()
    print("5. Published contracts:")
    all_ok &= check_file(skill_dir / "SKILL.md", "Skill documentation")
    all_ok &= check_file(skill_dir / "pyproject.toml", "uv dependency manifest")
    all_ok &= check_file(skill_dir / "scripts" / "thesis_format.py", "Unified dispatcher")
    all_ok &= check_file(skill_dir / "scripts" / "validate_thesis_ir.py", "ThesisIR validator")
    all_ok &= check_file(skill_dir / "profiles" / "zafu_2022" / "rules.yaml", "Default profile rules")

    # Summary
    print()
    if all_ok:
        print("=== All checks passed ===")
    else:
        print("=== Some checks failed ===")
        print("Please review the issues above.")

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
