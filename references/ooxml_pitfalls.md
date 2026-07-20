# OOXML Pitfalls

Use this reference when a DOCX repair touches template front matter, media relationships, or run-level formatting.

## 1. Template media overwrite

Trigger:
- template-based front-matter composition
- source and template both contain `word/media/*`

Failure signal:
- cover logo or integrity-page logo changes into a body figure
- a body figure changes into the cover logo

Wrong pattern:
- blindly copying all source `word/media/*` parts into the output package

Root cause:
- Word media filenames such as `image1.png` are package-local implementation details, not safe global identifiers

Required fix:
- preserve template front-matter media as protected assets
- import source media only through relationship-aware merge
- when contents differ, assign a distinct target part name instead of overwriting the template part

## 2. Relationship reuse by name or path only

Trigger:
- merge code attempts to reuse existing internal relationships
- template and source both contain targets like `media/image1.png`

Failure signal:
- body image points at the template logo
- some figures appear duplicated even though the source document had different images

Wrong pattern:
- reusing an existing relationship because the normalized target path matches

Root cause:
- path equality does not imply binary equality

Required fix:
- only reuse an internal relationship when both conditions hold:
  - normalized target part path matches
  - binary content matches exactly
- otherwise create a new relationship and a unique part name

## 3. Wrong base path for `document.xml.rels`

Trigger:
- target normalization for entries inside `word/_rels/document.xml.rels`

Failure signal:
- Word shows `无法显示图片`
- relationships resolve to paths like `_rels/media/image2.gif`

Wrong pattern:
- resolving `Target` relative to `word/_rels/document.xml.rels`

Root cause:
- package relationships for `document.xml.rels` are resolved relative to `word/document.xml`

Required fix:
- resolve all targets in `word/_rels/document.xml.rels` against `word/document.xml`
- validate that every internal image relationship points to an existing package part

## 4. Run-level direct formatting overrides paragraph style

Trigger:
- repair pass applies paragraph styles but preserves original runs
- source document contains direct `w:sz`, `w:szCs`, or `w:rFonts`

Failure signal:
- paragraph style says `zafu_body`, but Word still displays body text as small-four or another wrong size

Wrong pattern:
- assuming paragraph style alone is enough after patching an already formatted DOCX

Root cause:
- Word gives direct run formatting higher precedence than paragraph style defaults

Required fix:
- after repair passes that keep or rebuild runs, normalize run-level font and size from the resolved paragraph style
- do this as a final convergence step, not only during the first style assignment pass

## 5. Manual spaces used as indentation

Trigger:
- author created body indent with typed spaces or full-width spaces instead of paragraph indentation

Failure signal:
- body paragraphs show excessive left offset after the formatter applies first-line indent

Wrong pattern:
- applying `firstLineChars` while leaving typed leading spaces untouched

Root cause:
- manual whitespace and true paragraph indent stack together

Required fix:
- strip manual leading spaces before applying body or caption indentation rules
- keep this normalization limited to roles where leading manual whitespace is formatting noise rather than content

## 6. Regression expectations

Any change touching these areas should be checked against at least these cases:
- template and source share media filenames but not content
- repaired DOCX opened as a second-pass input does not reintroduce stale body run sizes
- body paragraphs with manual-space indentation collapse to clean text plus true first-line indent
- validators report `imageRelationshipsValid = true` for the final output
