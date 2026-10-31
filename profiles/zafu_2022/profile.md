# ZAFU 2022 Profile

This profile packages the current Zhejiang A&F thesis defaults into a profile-oriented view.

## Bound resources

- Template baseline:
  [浙江农林大学毕业论文模板参考.docx](../../浙江农林大学毕业论文模板参考.docx)
- Front-matter template:
  [assets/zafu_front_matter_template.docx](../../assets/zafu_front_matter_template.docx)
- Rules:
  [references/zafu_2022_rules.yaml](../../references/zafu_2022_rules.yaml)

## Default mode

- Ordinary formatting: `--quality final`, with rendered-page inspection by the Agent; use `--quality fast` only for an explicitly requested quick draft.
- DOCX: normalize safe regions and preserve complex native objects; do not treat preserved or deferred regions as already compliant.
- Markdown or TXT: build a new document using the selected template and rules.
- `audit-only` is the explicit read-only expert route. Risk findings identify regions needing follow-up; they do not automatically cancel ordinary formatting of safe regions.

## Safe subset defaults

Allowed automatic repairs:
- page geometry normalization
- body and heading style normalization
- high-confidence caption styling
- table and inline-image alignment normalization
- bibliography indent repair
- creation of missing body/abstract headers and normalization of empty/existing headers, with shared parts isolated from protected front pages
- verbatim insertion of missing required template pages and necessary header/page-number section isolation, following SKILL.md's composition and preservation checks

Blocked by default:
- any text write inside the template-derived cover and integrity pages, including
  thesis-title inference or population; insert these two pages verbatim and keep
  all blank fields unchanged, without proposing form filling as a repair
- bibliography mode conversion
- numbering rebuild
- replacement or rewriting of existing cover/integrity pages; missing-page insertion is handled separately above
- field/bookmark reconstruction
- floating-image conversion

## LaTeX adapter

`latex.yaml` selects the original minimal template and font preferences. Geometry and supported body/heading rules derive from the effective profile rules. This adapter is explicitly draft-level: it does not reproduce the immutable Word cover/integrity pages, certify department-specific compliance or promise identical pagination. External ZafuTemplatePublic files are not bundled.

## Resource compatibility

This profile currently mirrors existing resources that still live under the repository root, `assets/`, and `references/`.
