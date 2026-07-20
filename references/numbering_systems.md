# Numbering System Handling

This skill treats numbering as a dedicated subsystem, not just a regex problem.

## Families currently recognized

### Science / engineering decimal
- level 1: `1`
- level 2: `1.1`
- level 3: `1.1.1`
- level 4: `1.1.1.1`

### Humanities Chinese
- level 1: `一、`
- level 2: `（一）`
- level 3: `1、`
- level 4: `（1）`

### Humanities chapter/section
- level 1: `第一章`
- level 2: `第一节`
- level 3: `一、`
- level 4: `（一）`

## What is distinguished

The numbering analyzer tries to separate:
- true headings
- body enumerations
- figure captions
- table captions
- equation-numbered blocks
- TOC-like entries

## Current conservative repair policy

The current implementation will only auto-rewrite heading prefixes when:
- the paragraph is already a high-confidence heading
- numbering is not obviously mixed
- the repair volume is not excessive
- the document does not look template-like or TOC-heavy

Automatic numbering repair is suppressed when:
- families are mixed
- too many repair candidates appear at once
- the document is likely an instructional scaffold rather than a thesis body

Those cases go to `manualReview`.

## Next correct direction

The next major improvement should be:
- rebuilding actual numbering definitions in `word/numbering.xml`
- binding heading styles to numbering levels
- keeping manual prefix rewriting only as fallback
