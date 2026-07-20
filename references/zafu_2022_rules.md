# Zhejiang A&F University 2022 undergraduate thesis rules summary

This file summarizes the most important formatting rules extracted from the uploaded school material.

## Core page geometry
- Paper: A4
- Margins: top 2.7 cm, bottom 2.7 cm, left 2.7 cm, right 2.7 cm
- Header distance: 1.8 cm
- Footer distance: 1.85 cm
- Body line spacing: exactly 20 pt
- Paragraph spacing for normal body text: before 0 pt, after 0 pt
- Header after the TOC/body section: centered, small fifth-size Songti, text extracted from the thesis title
- Preserve the template header style and bottom border line instead of rebuilding a fake header line
- Footer: centered page number

## Core document parts
The body package generally includes:
- cover
- integrity commitment
- table of contents
- chinese abstract and keywords
- english title
- english abstract and keywords
- main text
- references
- acknowledgements
- appendix when needed

The current Zhejiang A&F repair preset treats:
- the first two pages as immutable cover / integrity matter copied from the validated template package
- the TOC as appearing before the Chinese abstract block
- the running header as beginning only after the TOC/front-matter transition

## Table of contents
- TOC appears before the Chinese abstract block in the current school layout
- TOC title should be `目  录` with two ASCII spaces between the characters
- TOC title: Songti, size 2, centered
- TOC title spacing: 1.5 lines before, 1 line after
- TOC title line spacing: multiple 2.41

## Title hierarchy options
### Science / engineering style
- level 1: `1`, `2`, ...
- level 2: `1.1`, `1.2`, ...
- level 3: `1.1.1`, `1.1.2`, ...

### Humanities style option A
- level 1: `第一章`
- level 2: `第一节`
- level 3: `一、`

### Humanities style option B
- level 1: `一、`
- level 2: `（一）`
- level 3: `1、`

## Typography rules frequently needed by repair logic
- Chinese thesis title: Heiti, third size, centered
- Chinese abstract label: bold Heiti size 5 with two-character indent
- Chinese abstract content: Kaiti size 5
- Chinese keywords label: bold Heiti size 5 with two-character indent
- Chinese keywords content: Kaiti size 5, separated by semicolons
- English title: Times New Roman, third size, centered
- Abstract label: bold Times New Roman size 5 with two-character indent
- Abstract content: Times New Roman size 5 with two-character indent
- Key words label: bold Times New Roman size 5 with two-character indent
- Key words content: Times New Roman size 5 with two-character indent, separated by commas
- Body text: Songti size 5 for chinese, Times New Roman size 5 for latin text, exactly 20 pt line spacing
- Level 1 heading: bold Kaiti size 4, centered, 6 pt before and after
- Level 2 heading: bold Heiti small-4, 4 ASCII spaces before the numbering prefix, 2 ASCII spaces between numbering and title text, 3 pt before and after
- Level 3 heading: Heiti size 5, 4 ASCII spaces before the numbering prefix, 2 ASCII spaces between numbering and title text, 3 pt before and after
- Level 4+ heading: Songti size 5, two-character indent, 3 pt before and after

## Figures, tables, equations
- Table title: keep adjacent to the table block and prefer the line below the table in the current repair preset
- Table block: center the table itself, center cell text horizontally, and center cell content vertically
- Figure title: below figure, centered, Songti small-5
- Equation numbering: right-aligned in parentheses, continuous or chapter-based but consistent

## Repair audit emphasis
- Detect front-matter order problems such as TOC / abstract / english abstract inversion
- Detect caption placement problems such as caption-before-table or figure-caption without a figure block
- Detect plain-text references like `图5-1` / `表2-1` and compare them with known captions before attempting automated cross-reference rewriting

## References
- Follow GB/T 7714-2015
- References heading is a centered heading style
- Zhejiang A&F uses the author-year route shown in the template notes, not numeric bracket citations
- Body citations should use round-parenthesis author-year form such as `（郭宝林等，2000）`
- The reference list should not add `[1] [2] [3]` numbering in the current school mode
- Reference entries should not inherit the body paragraph two-character first-line indent
- Order the bibliography as:
  - Chinese references first, sorted by the first author's pinyin
  - Foreign references second, sorted by the first author's surname
- Chinese references should use Songti size 5
- Foreign references should use Times New Roman size 5

Template evidence used by this preset includes:
- the template guidance paragraphs on GB/T 7714-2015
- the template examples showing author-year in-text citations
- the template note stating `先中文，后外文`
- the template examples showing unnumbered bibliography entries

## Automation caution points
- Users may want to preserve wording while allowing numbering repair.
- Many failures are structural, not merely stylistic.
- Mixed direct formatting and mixed numbering families should be surfaced explicitly.
