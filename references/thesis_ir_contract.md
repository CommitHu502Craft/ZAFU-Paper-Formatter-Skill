# ThesisIR Contract

`ThesisIR v2` is the central, source-independent semantic contract between source evidence extraction, strategy selection, Word execution, and validation.

It exists so the formatter can treat:

- clean `.docx`
- polluted `.docx`
- `.txt`
- `.md`

as different evidence sources that converge into one thesis-structure model.

## Core rule

Structure recovery happens before Word generation or late OOXML patching.

The Word layer may preserve, rebuild, or normalize presentation, but it should not be the first place where the system guesses:

- what is the abstract
- what is the keyword line
- which blocks are headings
- which numbering family is dominant
- whether the source is safe to preserve

## Canonical top-level fields

```json
{
  "schema": "paper-formatter.thesis-ir",
  "version": "2.0",
  "source": {"path": "...", "format": "docx|markdown|text", "capabilities": {}},
  "semanticBlocks": [],
  "semanticBlockCount": 0,
  "sourceType": "legacy compatibility value",
  "strategyCandidate": "preserve_first|hybrid_rebuild|text_rebuild|audit_only",
  "evidenceSummary": {},
  "frontMatter": {},
  "headingTree": [],
  "acknowledgements": {},
  "appendix": {},
  "captionBlocks": [],
  "attachableAssetCandidates": [],
  "bodyBlocks": [],
  "numberingAnalysis": {},
  "preserveableAssets": {},
  "assetAnchorAmbiguities": [],
  "ambiguities": [],
  "confidence": {}
}
```

`semanticBlocks` is authoritative for block order and shared semantics. The
remaining structural fields are compatibility projections until all downstream
components complete migration.

## Required modeling expectations

### `evidenceSummary`

Must preserve how the system knows what it knows.

Minimum expectations:

- source line count or paragraph count
- blank-line evidence for text sources
- style / numbering evidence count for DOCX sources
- candidate label positions for abstract / keywords / references

### `frontMatter`

Must represent front-matter structure independently of input type.

Minimum fields:

- `title`
- `abstractCn`
- `abstractEn`
- `keywordsCn`
- `keywordsEn`
- `confidence`
- `sourceBlocks`

### `headingTree`

Must represent recovered headings, not just raw lines.

Minimum fields per heading:

- `text`
- `level`
- `numberingFamily`
- `sourceIndex`
- `confidence`

### `acknowledgements`

Acknowledgements should not stay mixed into generic body paragraphs once the section is recoverable.

Minimum fields:

- `heading`
- `headingSourceIndex`
- `blocks`
- `confidence`

### `appendix`

Appendix content must be represented as explicit appendix sections, not left as late-stage heading guesses.

Minimum fields:

- `sections`
- `confidence`

Each appendix section should include:

- `heading`
- `headingSourceIndex`
- `blocks`
- `confidence`

### `captionBlocks`

Caption recovery is the bridge between semantic rebuild and later hybrid asset reattachment.

Minimum fields per caption:

- `sourceIndex`
- `text`
- `captionType`
- `label`
- `anchorHint`
- `confidence`

### `attachableAssetCandidates`

This is the execution-facing bridge between semantic rebuild and future asset reattachment.

It is more concrete than `assetAnchorAmbiguities`, but still intentionally conservative.

Minimum fields per candidate:

- `assetKind`
- `sourceIndex`
- `captionSourceIndex`
- `captionLabel`
- `sectionIndex`
- `anchorMethod`
- `sourceRegion`
- `confidence`
- `recommendedAction`

Current action values:

- `reattach_candidate`
- `manual_review`
- `skip`

`skip` is for preserveable assets that are intentionally kept outside the current hybrid body-reattachment scope, such as front-matter template graphics.

### `numberingAnalysis`

Must explicitly model numbering ambiguity.

Minimum fields:

- `familiesPresent`
- `dominantFamily`
- `conflicts`
- `confidence`

Mixed systems such as `1` / `一、` / `①` must be recorded here instead of being silently flattened later.

### `preserveableAssets`

For DOCX sources, this is where Word-native strengths are retained.

Minimum fields:

- `tables`
- `images`
- `equations`
- `headers`
- `footers`
- `sections`

### `assetAnchorAmbiguities`

This is the placeholder layer for future hybrid rebuild.

It records where preserved Word-native assets are not yet safe to reattach blindly.

Examples:

- caption recovered but figure anchor unresolved
- image count and caption count mismatch
- cross-reference points to an unresolved caption target
- caption layout issue reported by DOCX evidence extraction

### `ambiguities`

Ambiguities are first-class output, not hidden failures.

Examples:

- heading vs body enumeration
- uncertain abstract boundary
- mixed numbering conflict region
- likely preserved asset with uncertain anchor

`assetAnchorAmbiguities` should be mirrored into `ambiguities` when they affect automatic strategy confidence, but it remains a distinct top-level field so later hybrid execution can consume it directly.

## Current implementation boundary

The first milestone does not fully replace all existing rebuild logic.

It does establish:

- `extract_source_evidence.py`
- `thesis_ir.py`
- `recover_numbering.py`
- `select_strategy.py`
- `validate_visual_contracts.py`

and wires them into the dispatcher so strategy selection and visual checks are explicit artifacts rather than implicit behavior.
