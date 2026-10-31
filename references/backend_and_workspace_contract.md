# Backend and Workspace Contract

Public default is `--quality final`, with `--quality fast` for explicitly requested structural-only drafts; explicit `--check` wins. See [current architecture](formatting_first_architecture.md). Internal execution and review states support traceability; the completion claim to the user depends on the actual thesis meeting its requirements, with remaining issues listed. The Agent handles review records after inspecting the rendered pages; preview generation itself never attests review. Legacy audit artifacts remain available with `--pipeline legacy`.

This contract supersedes historical dispatcher defaults in the Word architecture notes.

## Pipeline boundaries

Ordinary runs: evidence/ThesisIR → independent writers or batch normalization → validation → bounded visual workflow → publication. The legacy expert route retains numbering/preflight/strategy artifacts.

The shared semantic stages run once for `--backend both`. The Word writer keeps native preservation/repair and hybrid rebuild. The LaTeX writer consumes `semanticBlocks` directly, not DOCX exported by the Word writer. `hybrid_rebuild` remains an internal Word strategy, not a synonym for two output backends.

Original DOCX repair forwards explicit IR role corrections to the in-place planner; it does not force a complete document rebuild. Unknown block IDs are skipped, not guessed from numeric substrings. Unsupported in-place roles and protected front matter remain unchanged and are reported.

The CLI's template-composition and whole-document rebuild route requires `--allow-structural-rebuild`; without it, ordinary DOCX normalization does not insert template front pages. Under the Agent workflow in SKILL.md, supplying missing required template pages verbatim and isolating headers/page numbers are within the formatting task: use the appropriate composition or targeted OOXML path and verify preservation. This does not authorize filling, editing or replacing existing cover/integrity pages, or unrelated whole-document rebuilding. Audit-only remains read-only. Newly generated Markdown/TXT documents may use template composition.

Word header policy is configured by `page.header_text_mode` (`fixed`, `thesis_title_after_toc`, `preserve`) and `page.header_text`. CLI `--header-mode thesis-title` maps to `thesis_title_after_toc`; an explicit `preserve` keeps headers unchanged. Otherwise, recognized body/abstract sections receive missing headers and empty headers are populated, including first/even variants when enabled. Pure TOC sections receive blank headers; shared parts must be isolated from immutable cover/integrity pages. Mixed or unrecognized scopes are reported for targeted follow-up rather than silently counted as compliant. The external LaTeX CLS is unchanged.

Modules:

- `formatter_core/checks.py`: independent check level and Word PDF export.
- `formatter_core/workspace.py`: safe names, unique runs, publication, relocations and relative artifact registry.
- `formatter_core/reports.py`: compact cross-backend summary and original-input hash check.
- `formatter_backends/word.py`: Word audit/repair/validation command planning.
- `formatter_backends/latex.py`: subset capability checks, TeX writer, asset isolation and bounded compilation.
- `thesis_format.py`: profile selection and orchestration; existing script CLIs remain usable.

## Python runtime

Python 3.11+ with the declared dependencies is sufficient to run the dispatcher. Users may install them with `python -m pip install -r requirements.txt` and invoke `python scripts/thesis_format.py ...`; uv is not required at runtime. The dispatcher checks for missing declared modules before importing its backends and offers an installation command using the same interpreter. It never installs dependencies automatically. Subprocesses use `sys.executable`, including when invoked from another working directory. Developers may retain `uv.lock` and use uv for reproducible development environments.

## Checks

`--quality final` defaults to `visual`; `--quality fast` defaults to `structural`. Explicit `--check` overrides either. Rendering does not imply an Agent has reviewed images.

`--export-pdf` independently requests Word PDF conversion. `layout` requests PDF text/geometry analysis without image export. `visual` additionally generates previews. `visualReviewed` stays false until actual candidate-bound image inspection is recorded with `record_visual_review.py`.

Profile validator selections cannot implicitly turn on default rendering. Requested PDF export is not cancelled by a profile's disabled render validator. LaTeX compilation is document generation and is independent of Word PDF export and review.

## Workspace and naming

Default base: `<input-parent>/thesis-output`. Explicit `--output-dir` replaces only the base, not the run isolation.

Layout: `<base>/<safe-project>/<timestamp-microseconds-random>/` with `manifest.json`, `deliverables/`, `reports/` and `work/`. Project defaults to input stem, never an inferred thesis title. File naming: `<project>__<profile>__<backend>.<extension>`.

The workspace refuses moves outside its own root and refuses overwrites. Publication moves a generated candidate rather than making legacy copies. Existing runs are not reused or automatically cleaned. The original input is never a publication candidate. Generated JSON references are remapped when artifacts move, and the manifest's artifact/deliverable paths are relative to the run root.

`manifest.json` records run ID, input SHA-256, requested configuration, status, commands, backend outcomes and artifact inventory. It is atomically replaced. A failed subprocess produces `status: failed` and an error record; any already-published partial backend artifacts must be read in conjunction with that status. Consumers must never assume mere file existence means success.

The CLI prints the exact manifest location. Library callers can use `Workspace.create` and its manifest path directly; do not scan another run's files as a fallback.

## Publication and statuses

- `dry-run`: shared commands recorded; no document generated or full backend plan promised.
- `completed`: requested processing completed; inspect warnings and check scope before claiming compliance.
- `completed_source_only`: LaTeX project generated, but compilation was skipped/unavailable.
- `completed_with_warnings`: e.g. Word delivered while the LaTeX engine was unavailable.
- `completed_with_hard_failures`: Word validation or LaTeX conversion/compilation has a hard failure; unsupported LaTeX content is preserved in an explicitly reported draft source package, not published as PDF.
- `failed`: pipeline exception or subprocess failure; consult the failure manifest.

Word candidates that fail the quality gate are retained in work, not published. LaTeX PDF publication requires conversion without errors, a successful compiler exit, stabilized XeLaTeX references (maximum three passes), and a readable nonempty PDF. A stale PDF after compiler failure is never used. CLI return code is nonzero for pipeline execution exceptions and is 2 for completed backend hard failures. Missing engines or explicit source-only generation are reported without an execution failure.

## LaTeX subset and trust

Inputs: Markdown/TXT only. Supported: paragraphs and inline emphasis, up to three heading levels, front-matter text, common allowlisted math, rectangular Markdown tables, local PNG/JPEG/PDF images, supplied reference text and labels. Every rendered block keeps its IR ID and source anchor in the conversion index.

Images must resolve inside the input directory; absolute paths, traversal, missing/corrupt images and unsupported types are reported. Assets use content-derived names. Formula commands use an allowlist; unknown or unsafe commands become escaped literal text and a conversion error. Plain text is TeX-escaped. Unknown block kinds and unsupported heading levels are reported, not silently omitted.

Compilation prefers XeLaTeX on PATH; auto may fall back to Tectonic on PATH. XeLaTeX uses `-no-shell-escape`, a finite timeout and at most three passes. Tectonic uses `--untrusted` and its internal rerun workflow. Neither mode installs a runtime; Tectonic can fetch resource bundles. The source ZIP excludes compile intermediates and logs. This is not a general sandbox for arbitrary external `.tex` files: the backend only compiles its own escaped/generated project and trusted repository template.

School-neutral semantics belong in IR. Word-specific preservation objects stay in source evidence/Word execution. Bibliographic metadata, cross-reference keys, complex TeX projects, OMML-to-TeX and DOCX native annotations are not inferred by this MVP.

## School rules and external reference

The generic template lives at `assets/latex/main.tex`; profiles without LaTeX settings receive a generic draft whose geometry and supported styles derive from effective profile rules. The ZAFU adapter uses `assets/latex/zafu-external.tex` with the original `ZafuThesis.cls`, not generic Word-derived style overrides. Its fixed revision and SHA256 are in `formatter_backends/external_class.py`. The first build downloads the class with a bounded timeout and integrity check into ignored `.local/dependencies/zafu-template/`; later builds work offline. Corrupt caches fail explicitly and are never silently replaced. Projects and source ZIPs contain the unchanged class and `UPSTREAM.json`.

The current ZAFU adapter has `compliance: draft`: Word's immutable cover/integrity pages, all caption/bibliography details and department-specific exceptions are not reproduced exactly. Font fallbacks must remain visible in compile warnings. Strict Word rules do not change this boundary.

External dependency: https://github.com/Stolorzs/ZafuTemplatePublic at `81c0d127d485aaaf00c5cbccc5d0d29affd446f2`. It is a nonofficial, department-specific template, not university-wide certification. No external class/sample thesis is committed. Missing required fonts/packages fail compilation, not permission to modify the class. Explicit `--latex-metadata` enables its cover/unsigned statement; separable bilingual abstracts use its macros. Personal data and signatures are never invented. Permission for redistributing the CLS has not been confirmed; see NOTICE before sharing generated ZIPs.

The adapter binds existing `zhsong`, `zhkai`, `zhhei`, `zhfs` families to installed SimSun, KaiTi, SimHei, FangSong when available; otherwise it retains upstream choices with warnings. No geometry, heading sizes or spacing are overwritten; the CLS remains byte-identical.

`latex_frontmatter.py` and `latex_bibliography.py` are independent external-class adapters. Explicit BibTeX uses pinned v2.1.4 2015 author-year BST resources with a separately named school bibliography variant, its license/source and provenance. It never changes the CLS or Word. Word produces only non-mutating `referenceFormatReport`; default citation-conversion-plan generation is retired. Detailed contracts are in `zafu_latex_adapter.md` and `zafu_reference_format.md`.

## Compatibility migration

Standalone low-level scripts keep their explicit output arguments. The public dispatcher intentionally replaces a flat `output/dispatch_manifest.json` plus duplicated `final/debug` layout with isolated `manifest.json` and registered paths. Regression fixtures resolve their own run manifest. For explicit structural-only runs, render assertions use `skipped` when conversion was not requested; Final requests rendering by default and must distinguish conversion failure from a skipped check.
