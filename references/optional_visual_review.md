# Optional Visual Review

Read this reference only when the user explicitly requests page previews or visual review. Default formatting does not use it.

Generate previews with `--check visual`. Open the current run's manifest to locate `work/render_validation/` (Word) or `work/latex-preview/` (LaTeX). Do not infer a review result from generated images alone. Without image-understanding capability, do not pretend to have reviewed them; text/geometry analysis is a narrower check, not an equivalent guarantee.

If explicitly asked to refine Word layout, use a single whitelisted plan with `--check visual --visual-refinement-plan plan.json`. The first candidate remains in work as `repaired_pass1.docx`. A new CLI invocation creates a new isolated run; it does not overwrite a prior run. The plan must target the newly produced candidate with stable anchors. The whitelist does not apply to LaTeX.

```json
{
  "actions": [
    {"type": "set_keep_with_next", "target": {"paragraphIndex": 12, "textPrefix": "第三章"}, "reason": "标题落在页尾", "confidence": 0.92},
    {"type": "set_table_header_repeat", "target": {"tableIndex": 2}, "headerRows": 1, "reason": "跨页表头不重复"}
  ]
}
```

Supported actions: `set_keep_with_next`, `set_keep_lines`, `set_page_break_before`, `clear_page_break_before`, `remove_empty_paragraph`, `remove_duplicate_page_break`, `set_widow_control`, `set_spacing`, `center_paragraph`, `scale_image_to_width`, `set_table_header_repeat`, `set_table_rows_no_split`, `center_table`.

Prefer `paragraphIndex` plus `textPrefix` for dual verification; DOCX source indices can originate from IR `docxParagraphIndex`, but rebuilt documents must use the generated candidate's indices. `blockId` anchors are useful only where the executor can resolve them. Mismatched targets must be skipped and reported, not guessed.

After explicit refinement, structural/rule validation must run again on the refined candidate before publication. Do not loop over repeated visual repair passes. Tell the user what was actually reviewed, what was only heuristically checked, and what remains unresolved.
