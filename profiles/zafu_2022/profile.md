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

- Preferred mode: `conservative-repair`
- High-risk polluted input: downgrade to `audit-only`
- Markdown or TXT input: route to `rebuild`

## Safe subset defaults

Allowed automatic repairs:
- page geometry normalization
- body and heading style normalization
- high-confidence caption styling
- table and inline-image alignment normalization
- bibliography indent repair
- known-template header text normalization

Blocked by default:
- any text write inside the template-derived cover and integrity pages, including
  thesis-title inference or population; insert these two pages verbatim and leave
  all title/personal fields for manual completion
- bibliography mode conversion
- numbering rebuild
- front-matter rebuild
- field/bookmark reconstruction
- floating-image conversion

## Notes

This profile currently mirrors existing resources that still live under the repository root, `assets/`, and `references/`.
