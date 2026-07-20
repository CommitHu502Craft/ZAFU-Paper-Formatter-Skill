# Repair Plan Schema

This project keeps the repair planner explicit and auditable.

Core rule:
- AI can propose repair actions
- only deterministic code can execute them
- every action must pass a schema gate before writing OOXML

## 1. Current vs target plan shape

Current examples in this repository still expose fields such as:
- `paragraphActions`
- `numberingActions`
- `manualReview`
- `rebuildRegions`

Target dispatcher-facing shape is:

```json
{
  "actions": [
    {
      "type": "apply_paragraph_style",
      "target": "p_0120",
      "role": "heading_2",
      "style_id": "Heading2_ZAFU",
      "confidence": 0.94,
      "risk": "low",
      "reason": "Text pattern and neighborhood context match level-2 heading"
    }
  ],
  "manualReview": [],
  "confirmationRequests": []
}
```

Dispatcher compatibility rule:
- legacy `paragraphActions` and `numberingActions` may continue to be produced
- the unified entrypoint should normalize them into `actions[]`
- execution must use the gated normalized view, not raw legacy arrays

## 2. Allowed action types

Default low-risk whitelist:
- `apply_paragraph_style`
- `apply_caption_style`
- `normalize_page_geometry`
- `normalize_paragraph_spacing`
- `normalize_body_indent`
- `normalize_header_text`
- `center_table`
- `center_table_cells`
- `scale_inline_image_to_text_width`
- `center_inline_image_paragraph`
- `normalize_bibliography_indent`
- `preserve_or_insert_toc_field`
- `manual_review`

Confirmation-first or blocked-by-default action classes:
- `rebuild_numbering_system`
- `convert_bibliography_mode`
- `rewrite_in_text_citations`
- `rewrite_heading_text`
- `rebuild_front_matter`
- `convert_plain_text_cross_reference_to_field`
- `rewrite_equation_numbering`
- `convert_floating_image_to_inline`

## 3. Required action fields

Each action should provide:
- `type`
- `target`
- `confidence`
- `risk`
- `reason`

Optional fields depend on action type:
- `role`
- `style_id`
- `profile_rule`
- `target_region`
- `expected_before`
- `expected_after`
- `requires_confirmation`

Field expectations:
- `target` should identify a stable IR entity when possible
- `confidence` should be numeric from `0.0` to `1.0`
- `risk` should be one of `low`, `medium`, `high`
- `reason` should explain why the action is safe or why it is blocked

## 4. Schema gate

Every action must pass all applicable gates before execution.

### Gate A: shape and whitelist

Check:
- action type exists
- action type is recognized
- required fields are present

Fail result:
- move action to `manualReview`

### Gate B: target existence

Check:
- target exists in the current audit / IR
- target is still writable in the chosen mode
- target is not inside a preserved or immutable region

Fail result:
- block execution
- record as `manualReview`

### Gate C: confidence threshold

Recommended thresholds:
- `low-risk automatic action`: `>= 0.85`
- `medium-risk suggestion`: do not auto-execute
- `high-risk suggestion`: never auto-execute without confirmation

Fail result:
- block execution
- record as `manualReview`

### Gate D: risk vs mode

Mode policy:
- `audit-only`: execute nothing
- `conservative-repair`: only `low` risk safe-subset actions
- `rebuild`: low-risk safe-subset actions plus profile-authorized rebuild flow; still no silent high-risk edits

Fail result:
- block action
- list it under `blockedAutoRepairs`

### Gate E: safe subset

Check whether the action belongs to the automatic safe subset.

If not in safe subset:
- do not execute automatically
- route to `manualReview` or `confirmationRequests`

### Gate F: non-destructive text protection

Check:
- action does not rewrite thesis body text
- action does not silently rewrite heading/caption wording
- action does not silently change citation content

Fail result:
- require explicit confirmation

### Gate G: confirmation-first

If action belongs to any confirmation-first class:
- execution must stop until the user approves
- approval request must describe scope, risk, and expected side effects

## 5. Manual review and confirmation routing

Use `manualReview` when:
- classification is ambiguous
- target exists but confidence is insufficient
- the operation is outside the safe subset
- the document is risky but the change is still worth surfacing

Use `confirmationRequests` when:
- the action is high impact
- text, citations, numbering, front matter, or cross-references may be rewritten
- the user must choose whether to proceed

## 6. Report expectations

Repair-plan output should eventually expose:

```json
{
  "documentRiskClass": "A",
  "recommendedMode": "conservative-repair",
  "actions": [],
  "blockedAutoRepairs": [],
  "manualReview": [],
  "confirmationRequests": [],
  "explanations": [
    "AI judgement is limited to block classification and repair planning.",
    "Final Word changes must be written through deterministic OOXML edits and logged."
  ]
}
```

The execution layer should also record:
- which actions passed the gate
- which were blocked
- which were deferred to manual review
- which required user confirmation
