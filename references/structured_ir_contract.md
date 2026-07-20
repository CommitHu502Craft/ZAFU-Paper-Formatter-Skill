# Unified ThesisIR v2 Contract

All supported inputs converge into one semantic model:

```text
DOCX ─┐
MD   ─┼─> source evidence adapter ─> ThesisIR v2 ─> preflight/planning/build/validation
TXT  ─┘
```

The semantic model is unified; the evidence layer is intentionally source-aware.
DOCX keeps OOXML structure as a sidecar because Word styles, relationships,
sections, fields, headers, footers and embedded assets cannot be represented
losslessly as plain text.

## Canonical identity

```json
{
  "schema": "paper-formatter.thesis-ir",
  "version": "2.0"
}
```

## Required source envelope

`source.format` is one of:

- `docx`
- `markdown`
- `text`

`source.capabilities` declares which evidence is available. For example, DOCX
must declare Word-native structure; Markdown may declare external image paths
and Markdown tables; TXT declares plain-text-only input.

## Canonical semantic stream

`semanticBlocks` is the single ordered stream consumed by new downstream code.
Every block contains:

- stable `id`
- `kind`
- semantic `role`
- optional text and heading level
- source anchor
- confidence
- source-specific attributes

Supported core kinds:

- heading
- paragraph
- caption
- reference
- image
- table
- equation
- page break

Markdown image paths and table rows live in `attributes`. DOCX tables and images
retain DOCX anchors and are reattached by the hybrid asset stage rather than
being flattened.

## Compatibility views

During migration, ThesisIR v2 also exposes the existing views:

- `frontMatter`
- `headingTree`
- `bodyBlocks`
- `captionBlocks`
- `references`
- `acknowledgements`
- `appendix`

They are compatibility projections, not alternative IRs. New components should
prefer `semanticBlocks`; removal of compatibility fields requires a later major
version.

## Commands

Default unified extraction:

```powershell
uv run python scripts\extract_structured_ir.py thesis.md --output thesis_ir.json
uv run python scripts\extract_structured_ir.py thesis.txt --output thesis_ir.json
uv run python scripts\extract_structured_ir.py thesis.docx --output thesis_ir.json
uv run python scripts\validate_thesis_ir.py thesis_ir.json
```

The deprecated source-specific structure is available only when explicitly
requested with `--legacy-source-ir`.

## Invariants

- Do not use Markdown as the universal lossless representation.
- Do use ThesisIR as the universal semantic representation.
- Preserve source evidence and anchors separately from normalized semantics.
- Do not discard DOCX-native assets or fields merely to make inputs look alike.
- Validate ThesisIR before strategy selection or document generation.
