from __future__ import annotations

from pathlib import Path
from typing import Protocol

from formatter_core.checks import CheckPolicy


class WordProfile(Protocol):
    template_docx: Path | None
    rules_yaml: Path
    style_map_yaml: Path | None
    front_matter_policy_yaml: Path | None
    validators_yaml: Path | None


def common_flags(command: list[str], profile: WordProfile, *, include_template: bool = True, include_validators: bool = True) -> list[str]:
    if include_template and profile.template_docx:
        command.extend(["--template-docx", str(profile.template_docx)])
    command.extend(["--rules-yaml", str(profile.rules_yaml)])
    if profile.style_map_yaml:
        command.extend(["--style-map-yaml", str(profile.style_map_yaml)])
    if profile.front_matter_policy_yaml:
        command.extend(["--front-matter-policy-yaml", str(profile.front_matter_policy_yaml)])
    if include_validators and profile.validators_yaml:
        command.extend(["--validators-yaml", str(profile.validators_yaml)])
    return command


def repair_commands(python: str, input_path: Path, profile: WordProfile, mode: str, output_dir: Path, policy: CheckPolicy, *, thesis_ir: Path | None = None, allow_structural_rebuild: bool = False) -> list[list[str]]:
    inspect = [python, "scripts/inspect_docx.py", str(input_path), "--output", str(output_dir / "audit_report.json")]
    validate_flags = ["--output", str(output_dir / "validation_report.json"), "--render-output-dir", str(output_dir / "render_validation"), *policy.flags()]
    commands = [common_flags(inspect, profile)]
    if mode == "audit-only":
        commands.append(common_flags([python, "scripts/validate_docx.py", str(input_path), *validate_flags], profile))
        return commands
    repaired = output_dir / "repaired.docx"
    plan = [python, "scripts/plan_docx_repairs.py", str(input_path), "--output", str(output_dir / "repair_plan.json")]
    apply = [python, "scripts/apply_ooxml_fixes.py", str(input_path), str(repaired), "--plan-json", str(output_dir / "repair_plan.json"), "--report-json", str(output_dir / "repair_execution.json")]
    if thesis_ir is not None:
        plan.extend(["--thesis-ir-json", str(thesis_ir)])
    if allow_structural_rebuild:
        plan.append("--allow-structural-rebuild")
        apply.append("--allow-structural-rebuild")
    validate = [python, "scripts/validate_docx.py", str(repaired), "--before-docx", str(input_path), "--plan-json", str(output_dir / "repair_plan.json"), *validate_flags]
    commands.extend([common_flags(plan, profile, include_validators=False), common_flags(apply, profile), common_flags(validate, profile)])
    return commands
