---
name: thesis-docx-formatter
description: Create, audit, repair, and standardize graduation theses from DOCX, Markdown, or plain-text input through a unified semantic model, profile rules, risk classification, and deterministic OOXML writes. Use when Codex must preserve thesis wording and immutable front matter, deeply inspect styles, numbering, sections, fonts, tables, figures, equations, citations, or page numbering, safely fix high-confidence layout drift, and emit auditable reports and render checks.
---

# Thesis DOCX Formatter

## 1. Purpose and scope

This skill upgrades a thesis document into a more compliant and reviewable Word deliverable.

Primary goal:
- perform **high-confidence, low-risk, auditable, rollback-friendly** thesis formatting repair
- preserve thesis wording and Word structures whenever possible
- stop at audit or manual review when the document is structurally risky

This skill is **not**:
- a universal Word control layer
- a generic “beautify any DOCX” tool
- an AI thesis rewriter
- a pipeline that forces all inputs through Markdown

## 2. Core principle: safe subset, not universal Word control

The core design rule is:

> automatically repair only a safe subset of thesis formatting problems; audit and escalate everything else.

This skill does not aim to fully control Word behavior. It aims to fix the common and defensible subset of graduation-thesis problems while leaving complex or destructive operations behind explicit review gates.

## 3. Input routes: md/txt create vs docx repair

Two routes must coexist.

### Markdown / TXT creation route

Use this route when the source is text-first and the user wants to generate a standard thesis `.docx`.

Pipeline:

```text
md/txt -> text-first IR -> preflight -> build source docx -> inspect -> plan -> OOXML repair -> validate -> render preview
```

Why this route exists:
- Markdown and TXT are easy to generate, diff, and version
- they are useful for thesis drafting from scratch
- they are not the lossless truth layer for Word repair

### DOCX repair route

Use this route when the user already has a Word thesis and wants conservative repair.

Pipeline:

```text
docx -> OOXML-backed IR -> preflight -> inspect -> repair plan -> deterministic OOXML patch -> validate -> render preview
```

DOCX repair defaults:
- do not rewrite thesis body text
- do not silently rebuild numbering systems
- do not delete images, tables, equations, notes, comments, or tracked changes
- do not use Markdown as a fake universal intermediate format

## 4. Operating modes

The unified interface should support three operating modes.

### `audit-only`

Use when:
- the user wants inspection only
- the document is high risk
- the environment lacks confidence for automatic repair

Allowed:
- structured IR extraction
- preflight checks
- risk classification
- audit reports
- validation reports
- render preview planning or rendering when available

Not allowed:
- deterministic repair writes to the thesis package
- rebuild actions
- confirmation-first rewrites

### `conservative-repair`

Default mode.

Use when:
- the document is template-aligned or thesis-like
- requested changes stay inside the safe subset

Allowed:
- low-risk, high-confidence OOXML formatting repair
- profile-based style normalization
- low-risk header/footer text normalization under known template structure
- low-risk image/table alignment normalization

Not allowed by default:
- body text rewriting
- large-scale numbering reconstruction
- bibliography mode conversion
- cross-reference field rebuilding
- front-matter regeneration

### `rebuild`

Use when:
- the input is Markdown or TXT
- the user explicitly authorizes reconstruction
- the existing DOCX is severely polluted and rebuild is the safer path

Allowed:
- source DOCX build from text-first IR
- template-based front-matter injection when configured
- deterministic post-build OOXML repair

Still not allowed by default:
- AI-generated thesis wording replacement
- silent destructive edits to protected Word structures

## 5. Non-destructive principles

Default rules:
- do not rewrite thesis body text
- do not automatically delete images
- do not automatically delete tables
- do not automatically delete equations
- do not automatically delete footnotes, endnotes, comments, or tracked changes
- do not silently reorder bibliography entries
- do not silently rewrite in-text citations
- do not silently rewrite heading wording
- do not silently rewrite figure or table caption wording
- require explicit confirmation for all high-risk modifications

If any workflow must edit body text, the change must:
- appear in `confirmationRequests`
- be approved first
- be logged line-by-line in the execution report

## 6. Document risk classification

Preflight must classify each source document.

### Class A: `template-aligned`

Characteristics:
- close to the school template
- style drift exists but structure is stable

Default handling:
- allow `conservative-repair`
- auto-apply most low-risk safe-subset actions

### Class B: `thesis-like-generic`

Characteristics:
- looks like a thesis but is not strongly aligned with the school template
- body, headings, captions, and page settings are recognizable

Default handling:
- allow `conservative-repair` only for generic low-risk items
- treat front matter, numbering system, headers/footers, and template-specific regions cautiously

### Class C: `high-risk-polluted`

Characteristics may include:
- repeated WPS saves or broken OOXML structure
- many floating images or text boxes
- mixed automatic and manual numbering pollution
- many sections with complex header/footer inheritance
- heavy OLE / SmartArt / embedded Excel usage
- bibliography and in-text citation mismatch
- missing or mixed style systems

Default handling:
- recommend `audit-only`
- block high-risk auto repair
- require explicit rebuild or confirmation-first escalation

Preflight report contract should include:

```json
{
  "documentRiskClass": "A|B|C",
  "riskReasons": [],
  "recommendedMode": "audit-only|conservative-repair|rebuild",
  "blockedAutoRepairs": []
}
```

## 7. Profile system

School-specific rules belong in profiles, not scattered hard-coded assumptions.

Planned structure:

```text
profiles/
├─ zafu_2022/
│  ├─ template.docx
│  ├─ rules.yaml
│  ├─ style_map.yaml
│  ├─ front_matter_policy.yaml
│  ├─ validators.yaml
│  └─ profile.md
└─ generic_cn_bachelor/
   ├─ rules.yaml
   ├─ style_map.yaml
   └─ profile.md
```

Current default profile:
- [profiles/zafu_2022/profile.md](profiles/zafu_2022/profile.md)

Profile responsibilities:
- template document binding
- style IDs and style-name mapping
- page geometry and section defaults
- heading/body/caption/reference expectations
- front-matter policy
- validators and auto-repair blocks

## 8. Template fingerprint

Template matching should be explicit instead of guessed informally.

Fingerprint inputs should include:
- `styles.xml` style IDs and style names
- numbering definitions
- sections and page geometry
- header/footer bindings
- theme fonts
- `docDefaults`
- front-matter page count
- TOC field position
- known paragraph signatures

Preflight / audit should expose:

```json
{
  "templateSimilarity": 0.82,
  "templateFingerprintMatched": true,
  "styleDrift": [],
  "numberingDrift": [],
  "sectionDrift": []
}
```

This supports the A/B/C risk decision.

## 9. Unified ThesisIR contract

This skill does not treat `python-docx` as the formatting truth layer.

Word truth comes from OOXML parts such as:
- `word/document.xml`
- `word/styles.xml`
- `word/numbering.xml`
- `word/theme/theme1.xml`
- `word/settings.xml`
- `word/fontTable.xml`
- `word/header*.xml`
- `word/footer*.xml`
- when needed: footnotes, endnotes, comments, relationships, and other package parts

IR expectations:
- `.md`, `.txt`, and `.docx` converge into `paper-formatter.thesis-ir` v2
- `semanticBlocks` is the single ordered semantic stream for downstream code
- source adapters remain different: DOCX retains OOXML evidence and native asset
  anchors, Markdown retains external asset/table syntax, and TXT retains line evidence
- source evidence differences must be expressed as capabilities and anchors, not
  as competing semantic models
- compatibility views may remain during migration, but new code must prefer
  `semanticBlocks`

Detailed reference:
- [references/structured_ir_contract.md](references/structured_ir_contract.md)

## 10. AI planning boundaries

Role split:

```text
AI role = block recognition + repair planning
Program role = deterministic OOXML modification
```

AI may:
- classify likely paragraph roles
- infer likely heading levels
- identify candidate repair actions
- assign confidence and risk labels
- send ambiguous cases to manual review

AI may not:
- directly rewrite OOXML or Word XML
- silently rewrite thesis content
- bypass risk or schema gates
- force high-risk structural actions

## 11. Repair plan schema gate

Every repair plan must pass a schema gate before execution.

Target shape:

```json
{
  "actions": [
    {
      "type": "apply_paragraph_style",
      "target": "p_0120",
      "role": "heading_2",
      "style_id": "Heading2_ZAFU",
      "confidence": 0.94,
      "risk": "low",
      "reason": "Text pattern and neighborhood context match level-2 heading"
    },
    {
      "type": "manual_review",
      "target": "p_0318",
      "confidence": 0.58,
      "risk": "medium",
      "reason": "Could be heading_3 or numbered body list"
    }
  ]
}
```

Execution gate checks:
- action type is in the whitelist
- target exists in the audit / IR
- confidence meets the minimum threshold
- risk level is allowed by the active mode
- action belongs to the safe subset
- action does not modify thesis wording unless explicitly confirmed
- confirmation-first actions remain blocked without user approval

Failing actions must:
- not execute
- move to `manualReview` or `confirmationRequests`

Current implementation notes:
- existing plans still expose `paragraphActions`, `numberingActions`, `manualReview`, and related fields
- future dispatchers should normalize these into the gated `actions[]` view before execution

Detailed reference:
- [references/plan_schema.md](references/plan_schema.md)

## 12. Writer backends

Python remains the primary controller.

Planned abstraction:

```text
writer_backends/
├─ python_lxml_writer
├─ dotnet_openxml_validator
├─ dotnet_openxml_writer_optional
├─ libreoffice_renderer
└─ word_com_renderer_optional
```

Python responsibilities:
- orchestration
- OOXML parsing and IR extraction
- YAML/profile rules
- AI planning
- low-risk OOXML patching
- report generation

.NET Open XML SDK responsibilities, when introduced:
- optional schema validation
- package/relationship checks
- optional high-risk part edits
- optional complex numbering, section, header/footer, field, or bookmark repair

Open XML SDK is **not** an OOXML replacement. It is an optional safer backend for selected operations.

## 13. Validation layers

Validation must be layered.

### Structural validation

Check:
- DOCX package integrity
- relationships and part references
- missing media
- relationship target normalization for `word/document.xml.rels`
- image relationship validity after template/source merge
- header/footer parts
- numbering/style part presence
- optional OpenXML schema validation

### Rule validation

Check:
- margins and page settings
- body/heading/caption/reference formatting
- run-level direct formatting that conflicts with the resolved paragraph style
- heading hierarchy
- caption conventions
- table alignment
- image width vs printable region
- bibliography indent behavior
- header/footer template expectations

### Render validation

Check, when environment allows:
- DOCX to PDF conversion succeeds
- page count is available
- key pages can be exported as preview images
- cover, TOC, abstract, first body page, and references pages are easy to review manually

## 14. Render validation

Render validation is part of the pipeline, not a future afterthought.

Preferred outputs:
- `repaired.pdf`
- `before_after_page_preview/`
- `render_validation_report.json`

Recommended first-stage behavior:
- verify that the repaired DOCX can render to PDF
- record page count
- export a small set of review pages
- flag pages for human review instead of pretending full visual QA is solved

Backends:
- LibreOffice headless preferred for generic environments
- Word COM optional on Windows when available

## 15. Automatic vs manual boundaries

### Safe subset: default low-risk automatic repairs

- page size and page margins
- body font, size, line spacing, paragraph spacing, and first-line indent
- high-confidence heading style assignment
- high-confidence figure/table caption style assignment
- normal table centering
- table cell horizontal and vertical centering
- research-style table normalization with strict three-line defaults: top rule, one header separator, and bottom rule; repeated grouped-header dividers are opt-in
- default equation normalization for identifiable LaTeX/plain-text formulas into Word math objects via MathML/OMML conversion
- inline image down-scaling when it exceeds the text area
- centering the paragraph containing an inline image
- known-template header text replacement
- known-template section-role geometry repair for cover, TOC/abstract, and body regions
- known-template page-number format repair such as TOC/abstract upper-Roman numbering and body-page Arabic restart
- bibliography paragraph indent repair
- preserving or inserting a TOC field without forcing page-number refresh

TOC ownership rule:
- the pipeline must have a single TOC owner for each route
- Markdown / TXT rebuild route: the source-docx builder may insert the TOC once in the front matter; downstream OOXML repair may only preserve, detect, reposition, or restyle it
- DOCX repair route: OOXML repair may preserve or insert a TOC only when the source truly lacks one
- no pipeline stage may "retry" TOC insertion later in the body or at end-of-document as a fallback

### Default manual-review or confirmation-first items

- floating image to inline conversion
- text boxes, SmartArt, OLE, embedded Excel
- OMML equation structure or equation-number rebuild
- full numbering-system rebuild
- mixed manual and automatic numbering normalization at large scale
- bibliography citation-mode conversion
- in-text citation matching and reordering
- cross-reference conversion into Word fields or bookmarks
- front-matter rebuild
- complex multi-section header/footer inheritance repair
- aggressive WPS-pollution recovery

Every repair action must be checked against the safe subset before execution.

## 16. Critical OOXML guardrails

These guardrails are mandatory for template-based repair and high-risk patch chains.

- treat template front-matter media as protected assets; source `word/media/*` parts must not blindly overwrite them
- insert template-derived cover and integrity pages verbatim; never infer or
  populate their thesis title or personal fields during ordinary formatting,
  and never restyle or rewrite text inside this immutable prefix
- never reuse an internal image relationship by filename or normalized target path alone; only reuse when the target part path and binary content are both identical
- resolve targets from `word/_rels/document.xml.rels` relative to `word/document.xml`, not relative to the `.rels` file path
- TOC insertion must be single-owner and idempotent across the whole pipeline; later stages may not insert an additional TOC when one already exists anywhere in the document body
- after any repair pass that preserves existing runs, normalize run-level font and size against the resolved paragraph style so stale direct formatting cannot override the target style in Word
- for body paragraphs that simulate first-line indent with typed spaces or full-width spaces, strip the manual leading whitespace before applying true first-line indent
- renaming template cover images is not a root fix; it only reduces collision probability and must not replace relationship-aware merge logic
- role recognition for abstract, keywords, references, acknowledgements, and appendix headings must tolerate common suffix annotations such as `（示例）`, `（非完整）`, and similar bracketed notes
- body-style fallback must explicitly cover mixed prose-plus-math paragraphs so inline or adjacent equations do not prevent normal body font and line-spacing normalization

Detailed failure modes and repair rules:
- [references/ooxml_pitfalls.md](references/ooxml_pitfalls.md)

## 17. Recommended workflows

### Unified dispatcher

Future primary entrypoint:

```powershell
uv run python scripts\thesis_format.py input.docx --profile zafu_2022 --mode conservative-repair --output-dir output
```

Dispatcher responsibilities:
- detect source type
- load the active profile
- choose the pipeline
- run preflight first
- honor the selected mode
- stop at audit when risk classification blocks repair

### Existing script workflow

DOCX input:

```powershell
uv run python scripts\preflight_semantic_normalization.py input.docx --template-docx "浙江农林大学毕业论文模板参考.docx" --rules-yaml references\zafu_2022_rules.yaml --output preflight_report.json
uv run python scripts\inspect_docx.py input.docx --template-docx "浙江农林大学毕业论文模板参考.docx" --rules-yaml references\zafu_2022_rules.yaml --output audit.json
uv run python scripts\plan_docx_repairs.py input.docx --template-docx "浙江农林大学毕业论文模板参考.docx" --rules-yaml references\zafu_2022_rules.yaml --output repair_plan.json
uv run python scripts\apply_ooxml_fixes.py input.docx repaired.docx --plan-json repair_plan.json --template-docx "浙江农林大学毕业论文模板参考.docx" --rules-yaml references\zafu_2022_rules.yaml --report-json repair_execution.json
uv run python scripts\validate_docx.py repaired.docx --before-docx input.docx --template-docx "浙江农林大学毕业论文模板参考.docx" --rules-yaml references\zafu_2022_rules.yaml --plan-json repair_plan.json --output validation_report.json
```

Markdown / TXT input:

```powershell
uv run python scripts\preflight_semantic_normalization.py thesis.md --rules-yaml references\zafu_2022_rules.yaml --output source_preflight_report.json
uv run python scripts\build_docx_from_markdown.py thesis.md --template-docx assets\zafu_front_matter_template.docx --rules-yaml references\zafu_2022_rules.yaml --output-dir output --keep-source-docx
```

## 18. Scripts

Current scripts:
- [scripts/thesis_format.py](scripts/thesis_format.py)
  Unified dispatcher skeleton for `audit-only`, `conservative-repair`, and `rebuild`.
- [scripts/extract_structured_ir.py](scripts/extract_structured_ir.py)
  Source-aware IR extraction.
- [scripts/preflight_semantic_normalization.py](scripts/preflight_semantic_normalization.py)
  Semantic preflight and confirmation-first checks.
- [scripts/extract_template_baseline.py](scripts/extract_template_baseline.py)
  Template baseline extraction.
- [scripts/inspect_docx.py](scripts/inspect_docx.py)
  Deep OOXML audit.
- [scripts/plan_docx_repairs.py](scripts/plan_docx_repairs.py)
  Repair-plan builder; future dispatchers should treat it as the producer of gated action candidates.
- [scripts/apply_ooxml_fixes.py](scripts/apply_ooxml_fixes.py)
  Deterministic OOXML writer for allowed actions.
- [scripts/validate_docx.py](scripts/validate_docx.py)
  Structural and rule validation; render validation is the next expansion point.
- [scripts/run_regression_fixtures.py](scripts/run_regression_fixtures.py)
  Fixture regression runner scaffold.

Shared engine:
- [scripts/docx_ooxml.py](scripts/docx_ooxml.py)

## 19. Reports

Expected report families:
- `structured_ir.json`
- `source_preflight_report.json`
- `preflight_report.json`
- `template_baseline.json`
- `audit_report.json`
- `repair_plan.json`
- `repair_execution.json`
- `validation_report.json`
- `render_validation_report.json`

Key report expectations:
- risk classification appears in preflight
- template fingerprint appears in preflight or audit
- repair gating decisions appear in the plan
- confirmation-first actions appear in `confirmationRequests`
- unresolved ambiguity appears in `manualReview`

## 20. Extension points

Planned extensions:
- stronger template fingerprinting
- richer safe-subset validation
- LibreOffice render validation
- optional Word COM render validation
- optional Open XML SDK validator and writer backends
- complex numbering rebuild under explicit confirmation
- field/bookmark repair
- OMML-aware equation repair
- bibliography matching and citation reconciliation
- visual QA

## 21. Regression fixtures

Planned fixture layout:

```text
tests/fixtures/
├─ normal_template_docx/
├─ manual_numbering_docx/
├─ messy_styles_docx/
├─ many_images_docx/
├─ many_tables_docx/
├─ floating_images_docx/
├─ wps_saved_docx/
├─ broken_toc_docx/
├─ numeric_references_docx/
├─ markdown_source/
└─ severe_pollution_docx/
```

Each fixture should eventually carry:
- an input document
- expected risk class
- expected auto-repairable items
- expected manual-review items
- expected validation summary

Mandatory regression themes for this skill:
- template front matter and body both contain `word/media/image1.*` style names with different bytes; the repaired DOCX must preserve both sets of images without replacement
- body image relationships imported from source must resolve to real package parts; validators must catch broken targets that would cause `无法显示图片`
- source body paragraphs with direct run font size such as `12pt` must still end up at the profile body size after repair
- source body paragraphs using typed spaces for indent must be normalized to true first-line indent without doubled indentation
- patching a previously repaired DOCX must not reintroduce stale run-level body formatting from the earlier output
- Markdown / TXT rebuild must not insert a second TOC title or TOC field later in the document when front matter already contains one
- reference-section detection must recognize heading variants such as `参考文献（示例，非完整）` and split numbered entries into separate bibliography paragraphs
- mixed prose-plus-math body paragraphs under level-2 or level-3 headings must still resolve to the profile body style instead of staying at `Normal`
- Markdown horizontal rules such as `---` used as visual separators must not be converted into page breaks

Planned regression entrypoint:

```powershell
uv run python scripts\run_regression_fixtures.py --profile zafu_2022 --output-dir test_output
```

## References

- [references/repair_architecture.md](references/repair_architecture.md)
- [references/preflight_normalization_workflow.md](references/preflight_normalization_workflow.md)
- [references/structured_ir_contract.md](references/structured_ir_contract.md)
- [references/plan_schema.md](references/plan_schema.md)
- [references/ooxml_audit_guide.md](references/ooxml_audit_guide.md)
- [references/ooxml_pitfalls.md](references/ooxml_pitfalls.md)
- [references/numbering_systems.md](references/numbering_systems.md)
- [references/template_rule_coordination.md](references/template_rule_coordination.md)
- [references/local_codex_extension_points.md](references/local_codex_extension_points.md)
