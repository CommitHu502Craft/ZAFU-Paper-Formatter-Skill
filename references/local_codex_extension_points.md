# Local Codex Extension Points

This skill is now organized so local Codex can continue implementation without discarding the current pipeline.

## Stable core contracts

Keep these stable:
- [scripts/thesis_format.py](../scripts/thesis_format.py) is the unified dispatcher contract for mode selection and profile loading
- [scripts/docx_ooxml.py](../scripts/docx_ooxml.py) remains the shared OOXML truth layer
- [scripts/extract_structured_ir.py](../scripts/extract_structured_ir.py) is the preferred ingest contract for new work; do not collapse all sources into pure Markdown
- [scripts/preflight_semantic_normalization.py](../scripts/preflight_semantic_normalization.py) is the default semantic gate before typesetting and should stay lightweight enough to run on both `docx` and `md/txt`
- [scripts/audit_bibliography_mode.py](../scripts/audit_bibliography_mode.py) is the bibliography-mode guardrail for the Zhejiang A&F preset
- [scripts/inspect_docx.py](../scripts/inspect_docx.py) remains the common audit entrypoint
- [scripts/plan_docx_repairs.py](../scripts/plan_docx_repairs.py) emits the repair plan consumed by repair and validation
- [references/plan_schema.md](plan_schema.md) stays the canonical repair-plan contract
- [profiles/zafu_2022/profile.md](../profiles/zafu_2022/profile.md) is the default profile narrative and policy contract

## High-value next steps

1. Heading recognition
   - add paragraph-neighborhood features
   - use blank-line context, table/image proximity, and section transitions
   - distinguish instructional examples from real thesis headings more accurately

2. Numbering repair
   - rebuild Word numbering definitions in `word/numbering.xml`
   - bind heading styles to numbering levels where safe
   - keep manual-prefix rewrite only as conservative fallback

3. Template projection
   - rebuild cover / commitment / abstract / english abstract / references scaffolds from template skeleton blocks
   - keep original thesis wording, but move it into a clean shell

4. Caption and equation handling
   - tighten figure/table caption placement heuristics
   - detect chapter-based caption numbering
   - preserve OMML and repair spacing/alignment before attempting conversion work

5. Validation depth
   - add bibliography checks
   - add page-number style checks
   - add direct-format pollution scoring
   - add header/footer binding validation per section, not only text validation
   - add render validation outputs such as PDF conversion and key-page previews

6. Structured IR deepening
  - expand the `.docx` IR to include table cell matrices, image rel targets, bookmark inventories, and equation anchors
  - deepen semantic normalization for abstracts, keywords, captions, references, citations, and heading numbering once the preflight contract stabilizes
  - keep `txt` / `md` IR lightweight, but make it convergent with the planner-facing contract used by `.docx`

## Safe invariants

- never overwrite the original source document
- never let AI write arbitrary final OOXML without deterministic code review points
- never silently rewrite thesis prose
- prefer false negatives plus manual review over false positives in numbering repair
- keep template baseline and rules baseline separate, then compare and reconcile them
