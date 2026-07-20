# Preflight Normalization Workflow

Preflight is the first decision gate in both thesis pipelines.

## 1. Pipeline position

### Markdown / TXT route

```text
md/txt -> text-first IR -> preflight -> build source docx -> inspect -> plan -> repair -> validate
```

### DOCX route

```text
docx -> OOXML-backed IR -> preflight -> inspect -> plan -> repair -> validate
```

Preflight exists to stop risky repair before later layout code makes a wrong structure look cleaner.

## 2. Preflight responsibilities

Preflight should inspect:
- front-matter order and labels
- heading numbering families
- manual vs automatic numbering pollution
- bibliography mode drift
- reference-entry indentation drift
- caption placement
- unresolved cross-reference candidates
- template similarity and drift
- structural risk signals such as floating images, text boxes, many sections, or WPS-like pollution

## 3. Document risk classification

Every preflight result should classify the document.

### `A` / template-aligned

Use when:
- template fingerprint matches strongly
- style system is mostly intact
- only localized formatting drift is present

Expected recommendation:
- `recommendedMode = conservative-repair`

### `B` / thesis-like-generic

Use when:
- thesis structure is recognizable
- template match is partial or weak
- body and headings remain recoverable

Expected recommendation:
- `recommendedMode = conservative-repair`
- block template-specific or large-scale actions unless confirmed

### `C` / high-risk-polluted

Use when:
- structure is heavily polluted
- the safe subset is too small to meaningfully fix the file
- risk signals suggest auto-repair may damage the document

Expected recommendation:
- `recommendedMode = audit-only`
- list blocked automatic repairs

## 4. Risk classification output contract

Preflight report should expose:

```json
{
  "documentRiskClass": "A|B|C",
  "riskReasons": [],
  "recommendedMode": "audit-only|conservative-repair|rebuild",
  "blockedAutoRepairs": [],
  "templateSimilarity": 0.82,
  "templateFingerprintMatched": true,
  "styleDrift": [],
  "numberingDrift": [],
  "sectionDrift": []
}
```

Interpretation:
- `riskReasons` explains why the class was assigned
- `blockedAutoRepairs` lists actions that must not run automatically
- `recommendedMode` is the mode the dispatcher should use unless the user explicitly overrides it

## 5. Confirmation-first triggers

Preflight should not silently enable high-impact operations.

Use `confirmationRequests` for:
- bibliography mode conversion
- bulk heading-number rewrite
- numbering-system rebuild
- front-matter rebuild
- cross-reference field generation
- body-text rewrites

Each confirmation request should include:
- user-facing question
- reason
- impact scope
- fallback if the user declines

## 6. Manual review routing

Use `manualReview` for:
- heading vs body-list ambiguity
- suspicious but not fully broken numbering
- caption patterns that may be false positives
- generic thesis blocks that may not match the selected profile
- unresolved references that need human interpretation

## 7. Current implementation alignment

Existing outputs in this repo already include:
- `confirmationRequests`
- `manualReview`

Next-stage alignment:
- add formal risk-class fields to preflight output
- add template-fingerprint summary
- let the dispatcher consume `recommendedMode` and `blockedAutoRepairs`
