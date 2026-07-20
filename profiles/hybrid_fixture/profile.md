# Hybrid Fixture Profile

This profile exists for regression fixtures that exercise the hybrid rebuild execution path.

It is intentionally lighter than the production ZAFU profile:

- template similarity is measured against a fixture-local template
- style-role requirements point to fixture-local paragraph styles
- render validation is disabled by default

The purpose is not school delivery quality.

The purpose is to provide a stable automated regression target for:

- `hybrid_rebuild` strategy selection
- `attachableAssetCandidates`
- hybrid asset reattachment reports
