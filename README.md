# ZAFU Paper Formatter Skill

一个面向浙江农林大学毕业论文的智能排版 Codex Skill，支持 DOCX、Markdown 和 TXT 统一导入，提供格式审计、学校模板适配、风险分级、可追踪的 Word/OOXML 修复及结果验证。

An auditable Codex Skill for formatting graduation theses from DOCX, Markdown, and plain-text sources.

## Features

- 使用统一的 ThesisIR 语义模型处理 DOCX、Markdown 和 TXT
- 基于学校 profile 执行格式检查和低风险自动修复
- 保留正文措辞以及不可变的封面、诚信页内容
- 检查样式、编号、分页、页眉页脚、图表、公式和参考文献
- 输出预检、修复计划、执行记录和验证报告
- 对复杂或高风险操作要求人工确认

## Requirements

- Python 3.11 or newer
- [uv](https://docs.astral.sh/uv/)
- Optional: LibreOffice and Poppler for rendered PDF/page-preview validation

## Setup

```powershell
git clone https://github.com/CommitHu502Craft/ZAFU-Paper-Formatter-Skill.git
cd ZAFU-Paper-Formatter-Skill
uv sync
```

## Usage

```powershell
uv run python scripts\thesis_format.py input.docx `
  --profile zafu_2022 `
  --mode conservative-repair `
  --output-dir output
```

Markdown and plain-text inputs use the same dispatcher:

```powershell
uv run python scripts\thesis_format.py thesis.md --profile zafu_2022 --output-dir output
uv run python scripts\thesis_format.py thesis.txt --profile zafu_2022 --output-dir output
```

## Validation

```powershell
uv run python scripts\verify_skill_health.py
uv run python -m unittest tests\test_unified_thesis_ir.py
uv run python scripts\run_regression_fixtures.py --profile zafu_2022 --output-dir test_output
```

## Safety

The formatter is intentionally conservative. It automatically applies only high-confidence, low-risk changes. Complex numbering rebuilds, floating objects, embedded Office content, citation conversion, and major section/header/footer changes remain confirmation-first operations.

The bundled ZAFU profile and template-related assets are intended for thesis-formatting research and personal academic use. Confirm institutional template redistribution requirements before broader reuse.

## Status

This is an early public release. Review generated documents and validation reports before formal submission.
