# Repair Architecture

This project intentionally uses a narrow, conservative architecture.

Goal:
- repair the common thesis-formatting subset well
- surface risk explicitly
- avoid pretending Word is fully controllable

## 1. Source-aware ingest, source-independent semantics

DOCX, Markdown and TXT use separate evidence adapters but converge into one
validated ThesisIR v2 semantic model.

### Markdown and TXT evidence adapters

Use for `.md` / `.txt`:
- extract text-first IR
- preserve semantic block order
- build source DOCX
- continue through the standard repair and validation pipeline

### DOCX evidence adapter

Use for `.docx`:
- parse OOXML directly
- preserve Word-specific structures
- plan low-risk repair without collapsing the document into Markdown

### Unified output

All adapters emit the same ordered `semanticBlocks` contract. Source-specific
evidence remains attached through capabilities, anchors and sidecar artifacts.
Downstream components must not invent a second source-specific IR.

## 2. OOXML truth layer

Shared parser:
- [scripts/docx_ooxml.py](../scripts/docx_ooxml.py)

The truth layer should expose:
- styles and inheritance
- numbering definitions
- theme fonts and defaults
- section geometry
- headers and footers
- paragraph/run inventories
- package relationships

Higher layers should reason over the same OOXML-backed truth.

## 3. Profile and baseline layer

Two baselines must coexist:
- physical baseline from the real template DOCX
- rule baseline from the profile YAML

Profile system responsibilities:
- bind template and rules
- define style mapping
- define front-matter policy
- define validator expectations
- define blocked automatic repairs

See:
- [references/template_rule_coordination.md](template_rule_coordination.md)

## 4. Preflight and risk layer

Preflight comes before repair planning.

Responsibilities:
- semantic normalization checks
- template fingerprint comparison
- document risk classification
- manual review routing
- confirmation-first routing

Outputs should decide whether the file can proceed to `conservative-repair`, should stay in `audit-only`, or needs `rebuild`.

## 5. Recognition and planning layer

AI-facing planning should answer:
- what a block most likely is
- which style or structure action is appropriate
- whether the action is safe
- whether the action should be deferred

AI output must remain a plan, not a writer.

## 6. Schema-gated deterministic writer

Writers should only execute actions that pass:
- action-shape checks
- target-existence checks
- confidence thresholds
- risk-vs-mode checks
- safe-subset checks
- non-destructive guards
- confirmation-first gates

Main current writer:
- Python + OOXML (`zipfile` / XML / `lxml` style approach)

## 7. Writer backend abstraction

Planned backend families:

```text
writer_backends/
├─ python_lxml_writer
├─ dotnet_openxml_validator
├─ dotnet_openxml_writer_optional
├─ libreoffice_renderer
└─ word_com_renderer_optional
```

Role split:
- Python remains the primary controller and default writer
- .NET Open XML SDK may become an optional validation or high-risk write backend
- renderers are separate from repair writers

## 8. Validation stack

Validation should be layered:
- structural validation
- rule validation
- render validation

Render validation is used to reduce Word-side uncertainty, not to claim perfect visual understanding.

## 9. Unified dispatcher

Target entrypoint:
- [scripts/thesis_format.py](../scripts/thesis_format.py)

Dispatcher duties:
- detect input type
- load profile
- run preflight first
- choose operating mode
- stop if risk classification blocks repair
- chain inspect -> plan -> apply -> validate when allowed

## 10. Regression fixtures

Regression fixtures should cover:
- normal template-aligned documents
- numbering pollution
- style drift
- image-heavy and table-heavy cases
- floating-image and WPS-pollution risk cases
- bibliography-mode drift
- Markdown source rebuild

Regression is part of architecture because the safe subset must remain stable over time.
