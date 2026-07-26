---
name: thesis-docx-formatter
description: 中文本科毕业论文自动排版(默认浙江农林大学 ZAFU 规范)。把 DOCX、Markdown 或 TXT 论文一键排成符合学校要求的 Word 并渲染 PDF 检查:标题层级与目录、页码分节、三线表、图题表题、LaTeX 公式转 Word 公式、摘要关键词、参考文献版式。触发词:毕业论文、论文排版、论文格式、浙江农林大学、Word 排版、DOCX 修复、三线表、Markdown 转论文。不用于普通代码任务或非论文类文档编辑。
---

# 毕业论文一键排版 Skill

把一份毕业论文(DOCX / Markdown / TXT)自动排版为符合学校规范的 Word,渲染 PDF 复查页面效果,最多做一次视觉二次修复后交付。

## 输入与输出

输入:一份 `.docx` / `.md` / `.txt` 论文文件(原文件永不修改)。

输出(`output/final/`,普通用户只需要看这里):

| 文件 | 说明 |
|---|---|
| `repaired.docx` | 排版完成、仍可编辑的 Word |
| `repaired.pdf` | 渲染预览(本机有 LibreOffice/Word/WPS 时) |
| `review_report.html` | 中文检查报告(含人工检查清单) |
| `contact_sheet.png` | 全文页面缩略图 |
| `critical_pages/` | 封面、目录、摘要、正文首页、参考文献等关键页高清图 |

中间 JSON 报告都在 `output/debug/`,普通用户无需查看。

## 核心原则

1. **不改写正文**——只调整格式,从不改动论文文字内容。
2. **不猜测个人信息**——姓名、学号、学院、题目等留给用户填写。
3. **尽量交付**——除非 DOCX 损坏、核心 XML 无法解析或正文严重丢失,警告不阻止生成结果;剩余风险写进报告。
4. **不覆盖原文件**——所有结果写入输出目录。
5. **不静默删除**图片、表格、公式、脚注、尾注、批注或修订。

## Agent 标准工作流(一次排版 + 最多一次视觉修复)

```text
1 自动排版      uv run python scripts/thesis_format.py 论文.docx --profile zafu_2022 --output-dir output
2 视觉检查      按下方 A/B 两条路线之一执行(先自问:我能不能看图?)
3 发现问题      编写 visual_refinement_plan.json(白名单动作,见下)
4 二次修复      uv run python scripts/thesis_format.py 论文.docx --profile zafu_2022 --output-dir output \
                 --visual-refinement-plan visual_refinement_plan.json
5 交付          把 output/final/ 的文件路径告诉用户,转述 review_report.html 的人工检查清单
```

第一次结果保留为 `output/repaired_pass1.docx`;最终结果始终是 `output/final/repaired.docx`。**最多执行一次视觉二次修复**,不要循环。

### 步骤 2A:有读图能力(多模态)的 Agent

- 查看 `output/final/contact_sheet.png`:有无空白页、异常留白、标题落在页尾、表格图片跨页断裂;
- 查看 `critical_pages/`:封面是否干净、目录是否只有一个、摘要标签是否加粗、正文页码是否从 1 开始、参考文献版式;
- 结合 `output/debug/visual_review_manifest.json` 的 `visualFindings` 交叉确认。

### 步骤 2B:没有读图能力的 Agent(重要:不要尝试读图)

不具备图像理解能力时,**跳过所有 PNG,不要假装看过图**。改用纯文本路线,检查质量同样有保障——`visualFindings` 本身就来自 PDF 逐页文本和页面几何分析,不依赖看图:

1. 读 `output/debug/visual_review_manifest.json`:`visualFindings`(空白页、标题落页尾、贴边、页数暴涨等)和 `unlocatedRegions`;
2. 读 `output/debug/semantic_pages.json` 的 `pageSummaries` 核对每页首行文字与页面顺序;
3. 需要修复时可先自动起草计划再人工复核:
   `uv run python scripts/suggest_visual_refinements.py output/repaired.docx --manifest-json output/debug/visual_review_manifest.json --output visual_refinement_plan.json`
4. 交付时把 `contact_sheet.png` 和 `critical_pages/` 的路径告诉用户,请用户自己翻看图片确认。

两条路线的后续步骤(3–5)完全相同。

### 步骤 3:visual_refinement_plan.json 白名单动作

```json
{
  "actions": [
    {"type": "set_keep_with_next", "target": {"paragraphIndex": 12, "textPrefix": "第三章"}, "reason": "一级标题落在页面末尾", "confidence": 0.92},
    {"type": "set_table_header_repeat", "target": {"tableIndex": 2}, "headerRows": 1, "reason": "表头跨页不重复"}
  ]
}
```

可用动作:`set_keep_with_next` / `set_keep_lines` / `set_page_break_before` / `clear_page_break_before` / `remove_empty_paragraph` / `remove_duplicate_page_break` / `set_widow_control` / `set_spacing`(beforePt/afterPt) / `center_paragraph` / `scale_image_to_width`(maxWidthCm) / `set_table_header_repeat`(headerRows) / `set_table_rows_no_split` / `center_table`。

段落目标锚点规则:
- 首选 `paragraphIndex`(document.xml 中第 N 个 `w:p`,0 起)+ `textPrefix` 双重校验;DOCX 源的 `paragraphIndex` 可直接取 `thesis_ir.json` semanticBlocks 里的 `docxParagraphIndex` 字段;
- `blockId`(block-NNNNN)是 IR 序号不是段落序号,单独使用会被拒绝执行——必须同时给 `textPrefix`;
- 不匹配的动作自动跳过并记录,绝不盲改。执行器:`scripts/apply_visual_refinements.py`。

## 语义覆盖(识别错误时)

自动识别把标题当正文、把关键词当标题时,不要改正文——写 `semantic_overrides.json` 重新标注语义角色:

```json
{"overrides": [
  {"blockId": "block-00142", "role": "heading_2", "headingLevel": 2, "reason": "无样式的二级标题"},
  {"blockId": "block-00361", "role": "figure_caption", "reason": "位于图片之后且以图3-2开头"}
]}
```

`blockId` 来自 `output/debug/thesis_ir.json` 的 `semanticBlocks`。角色支持 `heading_1..6`、`body`、`figure_caption`、`table_caption`、`equation`、`keywords`、`references_heading`、`reference_entry`、`acknowledgements_*`、`appendix_*`、`template_example`(忽略模板示例段)。无效 blockId 只报告不失败。运行时加 `--semantic-overrides semantic_overrides.json`。

## 统一命令与专家参数

```bash
uv run python scripts/thesis_format.py <输入文件> --profile zafu_2022 --output-dir output \
  [--semantic-overrides overrides.json] [--visual-refinement-plan plan.json] \
  [--mode audit-only|conservative-repair|rebuild] [--compliance default|strict-school] [--dry-run]
```

- Markdown/TXT 输入自动走重建路线(LaTeX 公式 `$...$` 转 Word 原生公式)。
- 风险等级 C 的 DOCX 也会尽量交付(策略层自动在 preserve_first / template_overlay / hybrid_rebuild 中选择);`audit-only` 仅当用户明确要求或文件无法解析时使用。
- 缺少 LibreOffice/Word/WPS 时跳过 PDF 渲染,DOCX 照常生成;缺 Poppler(pdftoppm)时跳过页面截图。

## 降级路径

| 情况 | 行为 |
|---|---|
| 无 PDF 渲染后端 | 生成 DOCX + 报告,提示用户自行打开检查 |
| DOCX 包损坏 / document.xml 无法解析 | 仅输出审计报告并如实告知 |
| 部分对象无法自动处理(SmartArt、OLE、复杂浮动图) | 保留原样 + 写入人工检查清单 |
| 分节 / 页码修复失败 | 仍交付文档并在报告中标记 |

## References 导航(按需阅读,不要默认全读)

- `references/thesis_ir_contract.md` — ThesisIR 契约与 semanticBlocks 结构
- `references/plan_schema.md` — repair plan schema
- `references/ooxml_pitfalls.md` — OOXML 修改陷阱(改 XML 前必读)
- `references/ooxml_audit_guide.md` — 审计字段说明
- `references/numbering_systems.md` — 中文论文编号体系
- `references/template_rule_coordination.md` — 模板与规则合并
- `references/hybrid_rebuild_contract.md` — 混合重建约束
- `references/architecture_deep_dive.md` — 完整架构说明(原 SKILL.md)
- `references/zafu_2022_rules.yaml` — ZAFU 规范的机器可读规则
