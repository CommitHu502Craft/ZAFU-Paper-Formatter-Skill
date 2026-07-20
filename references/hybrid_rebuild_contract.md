# Hybrid Rebuild Contract

`hybrid_rebuild` is the DOCX execution mode used when:

- Word-native assets are worth preserving
- semantic structure is not trustworthy enough to keep in-place
- full blind OOXML patching would be riskier than rebuilding text first

## Phase 2 boundary

Current Phase 2 behavior is:

1. extract evidence from the polluted DOCX
2. recover `ThesisIR`
3. build a clean intermediate DOCX from recovered text structure
4. reattach high-confidence image and simple table candidates when a matching caption block exists
5. run the existing conservative repair backend on the asset-augmented intermediate DOCX
6. emit explicit hybrid reports and ambiguity artifacts

Current Phase 2 does **not** yet:

- reuse section metadata selectively in a confidence-scored way
- guarantee full visual parity with the original DOCX
- safely resolve low-confidence or competing asset anchors without manual review

That limitation is deliberate. The system must prefer explicit deferment over wrong attachment.

## Required inputs

Minimum required artifacts:

- `source_evidence.json`
- `thesis_ir.json`
- `strategy_selection.json`
- `preflight_report.json`

## Required `ThesisIR` fields

Hybrid execution consumes:

- `headingTree`
- `bodyBlocks`
- `captionBlocks`
- `acknowledgements`
- `appendix`
- `attachableAssetCandidates`
- `assetAnchorAmbiguities`

## Execution contract

### Step 1: text-first rebuild

The executor must produce an intermediate DOCX whose text order is driven by `ThesisIR`, not by the original polluted paragraph order.

At minimum this rebuild must retain:

- front matter
- recovered headings
- body paragraphs
- captions as text blocks
- acknowledgements
- appendix
- references

### Step 2: asset decision layer

Each preserveable asset must be assigned one of:

- `reattach_candidate`
- `manual_review`
- `skip`

The decision must be based on explicit candidate objects, not hidden heuristics.

Current implemented subset:

- high-confidence image candidates can be reattached before their matched caption
- high-confidence simple table candidates can be reattached before their matched caption
- low-confidence candidates stay deferred

### Step 3: reporting

The execution report should eventually expose:

- `intermediateDocx`
- `assetAugmentedDocx`
- `attachedAssets`
- `manualReviewAssets`
- `skippedByPolicy`
- `skippedAssets`
- `assetAnchorAmbiguityCount`

## Confidence rule

The hybrid executor must not silently attach an asset when:

- no caption is recoverable
- multiple plausible anchors compete
- cross references remain unresolved
- section proximity conflicts with caption proximity

Those cases belong in `manual_review`.

## Acceptance direction

Phase 2 is good enough when:

1. the system no longer fakes hybrid mode while still patching the original DOCX directly
2. text-first rebuild is real
3. asset reattachment decisions are explicit
4. low-confidence cases are deferred rather than guessed
