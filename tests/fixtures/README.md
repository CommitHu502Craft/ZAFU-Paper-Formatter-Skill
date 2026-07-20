# Regression Fixtures

This directory reserves fixture slots for stable thesis-formatting regression cases.

Suggested layout:

```text
tests/fixtures/
├─ normal_template_docx/
├─ template_media_name_collision_docx/
├─ reused_repaired_docx_run_override/
├─ manual_space_indent_docx/
├─ manual_numbering_docx/
├─ messy_styles_docx/
├─ many_images_docx/
├─ many_tables_docx/
├─ floating_images_docx/
├─ wps_saved_docx/
├─ broken_toc_docx/
├─ numeric_references_docx/
├─ plain_txt_clean/
├─ plain_txt_backmatter_sections/
├─ plain_txt_inline_abstract/
├─ plain_txt_numeric_reference_conflict/
├─ plain_txt_equation_conversion/
├─ docx_hybrid_rebuild_candidate/
├─ docx_hybrid_mixed_confidence_assets/
├─ docx_hybrid_table_mixed_confidence/
├─ docx_plain_text_equation_conversion/
├─ markdown_source/
├─ render_preview_enabled_normal_template_docx/
└─ severe_pollution_docx/
```

Each fixture should eventually include:
- the input document
- `expected.json`
- optional supporting notes or reduced samples

Recommended additional regression themes:
- `template_media_name_collision_docx/`
  source and template intentionally share media filenames such as `image1.png` while bytes differ; expected outcome is preserved cover assets plus preserved body assets
- `reused_repaired_docx_run_override/`
  input is a previously repaired DOCX with stale direct run sizes; expected outcome is body text converging back to profile size
- `manual_space_indent_docx/`
  body paragraphs use typed spaces or full-width spaces for indent; expected outcome is cleaned leading text plus true first-line indent
- `many_images_docx/`
  should also assert that final validation reports `structuralValidation.checks.imageRelationshipsValid = true`

Current minimal `expected.json` contract:

```json
{
  "input": "../../../some_input.docx",
  "profile": "zafu_2022",
  "mode": "audit-only",
  "expected": {
    "sourceKind": "docx|text",
    "dispatcherStatus": "completed|stopped_after_preflight|stopped_for_confirmation",
    "effectiveMode": "audit-only|conservative-repair|rebuild",
    "documentRiskClass": "A|B|C",
    "recommendedMode": "audit-only|conservative-repair|rebuild",
    "minConfirmationRequests": 1,
    "blockedAutoRepairsIncludes": [
      "rewrite_heading_numbering_tree"
    ],
    "requiresValidationReport": true,
    "requiresRenderValidationReport": true,
    "requiresRenderPreviewDir": true,
    "requiresPreviewManifest": true,
    "requiresRepairPlan": false,
    "requiresRepairExecution": false,
    "requiresRepairedDocx": false,
    "requiresRepairedPdf": false,
    "validationSummary": {
      "renderStatus": "ok|unavailable"
    },
    "renderValidation": {
      "status": "ok|unavailable"
    },
    "previewExport": {
      "available": true,
      "backend": "pdftoppm|null"
    },
    "previewImageCount": 0,
    "previewManifestMinPages": 0,
    "previewManifestPageCount": 0,
    "previewManifestPageLabelsIncludes": [
      "cover",
      "references"
    ],
    "previewManifestStatusesInclude": [
      "planned",
      "exported"
    ],
    "previewManifest": {
      "pages": []
    }
  }
}
```

Current environment example:
- This repository currently often runs without `LibreOffice/soffice` in `PATH`.
- In that case a render-capable DOCX fixture can assert:

```json
{
  "expected": {
    "requiresRenderValidationReport": true,
    "requiresRenderPreviewDir": true,
    "requiresPreviewManifest": true,
    "renderValidation": {
      "status": "unavailable"
    },
    "previewExport": {
      "available": false,
      "backend": null
    },
    "previewImageCount": 0,
    "previewManifestPageCount": 0
  }
}
```

Future PDF-preview-enabled example:
- If a future environment has both `LibreOffice/soffice` and `pdftoppm`, the same fixture contract can be tightened to assert exported preview pages.

```json
{
  "expected": {
    "requiresRenderValidationReport": true,
    "requiresRenderPreviewDir": true,
    "requiresPreviewManifest": true,
    "renderValidation": {
      "status": "ok"
    },
    "previewExport": {
      "available": true,
      "backend": "pdftoppm"
    },
    "previewImageCount": 5,
    "previewManifestMinPages": 5,
    "previewManifestPageLabelsIncludes": [
      "cover",
      "toc",
      "abstract",
      "firstBodyPage",
      "references"
    ],
    "previewManifestStatusesInclude": [
      "exported"
    ]
  }
}
```

Current runner behavior:
- executes `scripts/thesis_format.py`
- reads `dispatch_manifest.json`
- checks dispatcher and preflight decision fields
- can assert `strategy_selection.json` fields such as chosen strategy and strategy reasons
- can assert `thesis_ir.json` fields such as strategy candidate, overall confidence, reference normalization signals, and phase-2 structure fields like `acknowledgements / appendix / captionBlocks / attachableAssetCandidates / assetAnchorAmbiguities`
- can assert `attachableAssetCandidateSummary.actionCounts` to separate `reattach_candidate / manual_review / skip`
- can assert `assetAnchorAmbiguities` kinds explicitly
- when manifest artifacts declare validation outputs, verifies:
  - `validation_report.json` exists
  - `structuralValidation / ruleValidation / renderValidation / validationSummary` are present
  - `render_validation_report.json` exists and contains its minimal schema
- `preview_manifest.json` exists when `previewExport.manifestPath` is declared
- supports repair-mode artifact assertions:
  - `requiresRepairPlan`
  - `requiresRepairExecution`
  - `requiresRepairedDocx`
  - `requiresRepairedPdf`
- supports equation-specific repair assertions:
  - `repairExecutionLogActionIncludes`
  - `repairedDocxMinOmmlCount`
- supports optional hybrid-mode artifact assertions:
  - `requiresHybridAttachmentReport`
  - `hybridAttachmentReport.*`
- supports optional hybrid action-count assertions:
  - `hybridAttachmentReportActionCounts`
- supports optional hybrid execution report assertions:
  - `hybridRebuildReport.*`
  - `hybridRebuildReportActionCounts`
- supports optional fixture-level assertions for `validationSummary.*`, `renderValidation.*`, `previewExport.*`, preview image count, preview page count, preview labels, and preview statuses
- supports optional fixture gating fields in `expected.json`:
  - `profile`
  - `tags`
  - `requiresCommands`
- supports `--include-tag TAG` to run a subset such as the optional render-preview group
- marks unmet capability fixtures as `skipped` instead of `failed`
- does not yet compare full repaired DOCX content or visual render output

Optional render-preview group:
- `render_preview_enabled_normal_template_docx/expected.json` is a capability-gated fixture template for environments that have both `soffice` and `pdftoppm`
- default runs will skip it automatically when those commands are unavailable
- to run only that group:

```bash
uv run python scripts/run_regression_fixtures.py --profile zafu_2022 --include-tag render_preview_enabled --output-dir test_output/render_preview_enabled/
```
