# Template And Rule Coordination

Template data and profile rules complement each other. Neither replaces the other.

## 1. Current default profile resources

Physical template baseline:
- [浙江农林大学毕业论文模板参考.docx](../浙江农林大学毕业论文模板参考.docx)

Current front-matter template asset:
- [assets/zafu_front_matter_template.docx](../assets/zafu_front_matter_template.docx)

Current rules YAML:
- [references/zafu_2022_rules.yaml](zafu_2022_rules.yaml)

Profile view:
- [profiles/zafu_2022/profile.md](../profiles/zafu_2022/profile.md)

## 2. What the template contributes

- real style IDs and inheritance chains
- real numbering definitions and bindings
- real section and header/footer structures
- real theme fonts and document defaults
- known front-matter package structure

## 3. What the rules contribute

- target body and heading formatting
- target page geometry
- supported numbering families
- bibliography policy
- front-matter policy
- validator expectations
- blocked auto-repair categories

## 4. Coordination strategy

Preferred workflow:
1. extract the template baseline from the real DOCX
2. load profile rules
3. compare them and record mismatches
4. use the combined view for preflight, planning, repair, and validation

The rule system should not guess Word-side truth that the template can reveal directly.

The template should not silently override explicit policy choices that belong in the profile.

## 5. Template fingerprint and drift

Coordination must expose:
- `templateSimilarity`
- `templateFingerprintMatched`
- `styleDrift`
- `numberingDrift`
- `sectionDrift`

These fields are used by risk classification and mode recommendation.

## 6. Immutable and sensitive regions

For the current Zhejiang A&F workflow:
- cover and integrity pages must be inserted verbatim from the validated template
- no automatic step may infer, populate, merge, restyle, or otherwise rewrite text
  inside those two pages; this prohibition includes the thesis title as well as
  year, college, name, student ID, major/class, advisor/title, signature, and date
- the TOC position is template-sensitive
- the body header after the TOC is title-sensitive
- header border styling should be preserved from the template where possible

These regions should not be normalized like ordinary thesis body paragraphs.

If a future workflow supports front-page form filling, it must be a separate,
explicitly requested operation with a user-supplied field map and before/after
audit. It must never run as part of ordinary formatting or infer values from the
body manuscript.

## 7. Migration toward profile-managed resources

This repository still contains historical resources under `assets/` and `references/`.

Planned direction:
- move or mirror those resources under `profiles/`
- let the dispatcher load profile-local defaults first
- keep backward compatibility with existing paths during migration
