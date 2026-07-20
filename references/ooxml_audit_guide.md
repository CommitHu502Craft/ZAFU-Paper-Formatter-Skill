# OOXML Audit Guide

This skill audits Word documents from OOXML outward, not from Word UI assumptions inward.

## Why this matters

Word formatting problems often come from:
- inherited style properties that are not visible on the current paragraph
- builtin heading styles inheriting theme colors or fonts
- only `ascii` / `hAnsi` being set while `eastAsia` remains inherited
- numbering definitions living in `numbering.xml`, not in the visible text alone
- section properties being attached to paragraphs or trailing body `sectPr`

High-level wrappers often hide those details.

## Current audit coverage

The shared parser reads:
- `word/document.xml`
- `word/styles.xml`
- `word/numbering.xml`
- `word/theme/theme1.xml`
- `word/settings.xml`
- `word/fontTable.xml`
- `word/header*.xml`
- `word/footer*.xml`
- `word/footnotes.xml` and `word/endnotes.xml` when present
- `word/_rels/document.xml.rels`

## Effective formatting order

Current effective formatting resolution follows this order:

1. `docDefaults`
2. paragraph style inheritance via `basedOn`
3. numbering level properties when present
4. direct paragraph / run formatting
5. theme-font resolution for unresolved slots

The audit output includes:
- effective paragraph spacing, alignment, indent, outline level
- effective run fonts, size, bold/italic, color, language
- raw font-slot sources and notes when CJK text is inheriting an unresolved `eastAsia` slot

## Audit outputs to read first

- `issues`
- `styleAnalysis`
- `fontSlotAnalysis`
- `numberingAnalysis`
- `sections`
