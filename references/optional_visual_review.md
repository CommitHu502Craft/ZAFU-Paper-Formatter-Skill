# Final 视觉验收

普通排版默认 `--quality final`，请求 PDF、几何检测、低清总览和高清关键/异常页。Fast 默认不渲染；显式 `--check` 优先于 quality。CLI 不调用视觉模型，实际看图由 Agent 完成。

读取当前运行 manifest 的 renderValidationReport，以报告的 contactSheets、criticalPages 和 visualFindings 路径为准，不猜固定目录。先看总览的整体分页和空页，再实际打开高清的封面、诚信页、中英文摘要、全部目录页、正文/章节首页、全部图表页、全部参考文献页及异常页。不能只抽第一张图或参考文献首页。缩略图不能证明细小文字正确；缺失区域继续列为待核查，必要时补导出高清页。

逐项核对实际版面：

- 封面/诚信页与原始模板一致，空白字段、下划线、图片及分页保留；禁止以“封面信息不全”为由填写或重排。
- 目录条目与正文标题、层级、页码一致；目录尾部仅剩少量条目单独占页时检查多余空段、分页与段距，不为挤页擅改学校规定的行距。
- 中英文题目、摘要、关键词标签的字号、加粗、标点和空格正确；特别检查 `Abstract:Text`、`Keywords text`。正文内容不改写，普通标签格式问题应修复并复验。
- 正文/摘要每节实际出现规定页眉，首页不同和奇偶页变体不能漏；页码位置、格式、接续与目录一致；封面不得被连带加上正文页眉页码。
- 图片完整、比例正常；按图像与题注的几何位置、编号、相邻正文共同确认归属。图题在图下、表题在表上，核对跨页拆分、重叠、过大间距、表头续页。页眉或页码不算正文首末行，不能因“第一行是图题”就断言图文跨页。
- 参考文献全部条目的字体、字号、缩进、行距及跨页表现一致，条目数量与原稿相符；不能把“未识别到条目”当作没有问题。

异常尽可能提供 pageNumber、boundingBox、blockId、docxParagraphIndex、tableIndex、objectId、issueType、confidence 和 mappingStatus。优先使用 Poppler 文字框；降级的 PDF 基线坐标不等于字形边界，未知右边界为 null，依赖准确坐标的判断停止。重复文字标为 ambiguous，禁止按截图猜段落；不能确认的题注归属只提示。

每次候选绑定 SHA256。一次集中计划合并重复动作、隔离冲突并排序，修复后重新检查；发现新的、能够可靠定位的格式问题时继续定向修复，不以“已经做过一次”为由留下已知问题。连续两次同类修复无改善时停止重复方法并列出剩余问题；新增硬失败须恢复最近通过完整性检查的版本。新候选必须重新渲染，不能用旧图片验收。原生复杂区域须保持对象完整；补齐缺失标准前置页及页眉按 SKILL.md 的授权范围执行。

程序使用 PDF 图片变换矩阵和准确文字框测量图题位置，结果只是待看图确认的提示。嵌套 Form 支持到8层；旋转/裁剪页、内联图片、矢量图、裁剪效果和无法确定归属的多图版面不作自动几何判定，必须实际看图。几何未告警不等于视觉无问题。

## 实际审阅的证据

仅在打开总览和所有导出关键/异常页后，建立记录：

```json
{
  "candidateSha256": "当前交付DOCX的SHA256",
  "reviewer": "实际审阅者或Agent名称",
  "reviewedImages": [{"path": "该运行内图片路径", "sha256": "该图片SHA256"}],
  "issueDispositions": {
    "异常类型:页码": {"status": "dismissed", "reason": "看图后排除此提示的具体依据"}
  },
  "additionalIssues": []
}
```

reviewedImages 包含每张总览和 criticalPages 图片，路径为运行内相对或绝对路径。使用 hashlib 计算文件哈希。未处理的程序提示和 additionalIssues 保持未解决。

```powershell
python scripts/record_visual_review.py "该次运行目录" --record review.json
```

命令验证候选和图片哈希、截图覆盖范围，更新 manifest、summary JSON 和 HTML。它不会看图，也不证明记录者实际使用过视觉能力，记录者负责真实性。这些是 Agent 在成品检查后的内部记录；用户只需拿到排版后的论文与具体未解决项，不需要操作状态标记。

分别记录 rendered、previewGenerated、visualReviewed、visualReviewStatus、unresolvedVisualIssues。渲染不可用为 render_unavailable；无视觉能力时保留 pending_agent_review/pending_manual_review，只能交付待复核候选 DOCX。Windows auto 优先 Word，其次 WPS、LibreOffice；显式后端不切换，其他编辑器分页可能不同。

## 兼容动作

Generate previews with `--check visual`. LaTeX previews remain under `work/latex-preview/`; the Word review-record command does not certify LaTeX review.

Ordinary formatting includes fixing reliably located layout defects. Use a whitelisted plan with `--check visual --visual-refinement-plan plan.json`. The first candidate remains in work as `repaired_pass1.docx`. A new CLI invocation creates a new isolated run; it does not overwrite a prior run. The plan must target the newly produced candidate with stable anchors. The whitelist does not apply to LaTeX.

```json
{
  "actions": [
    {"type": "set_keep_with_next", "target": {"paragraphIndex": 12, "textPrefix": "第三章"}, "reason": "标题落在页尾", "confidence": 0.92},
    {"type": "set_table_header_repeat", "target": {"tableIndex": 2}, "headerRows": 1, "reason": "跨页表头不重复"}
  ]
}
```

执行计划还必须具有当前候选的 `candidateSha256`；不得把旧版本计划应用到新内容。

Supported actions: `set_keep_with_next`, `set_keep_lines`, `set_page_break_before`, `clear_page_break_before`, `remove_empty_paragraph`, `remove_duplicate_page_break`, `set_widow_control`, `set_spacing`, `center_paragraph`, `scale_image_to_width`, `set_table_header_repeat`, `set_table_rows_no_split`, `center_table`.

Prefer `paragraphIndex` plus `textPrefix` for dual verification; DOCX source indices can originate from IR `docxParagraphIndex`, but rebuilt documents must use the generated candidate's indices. `blockId` anchors are useful only where the executor can resolve them. Mismatched targets must be skipped and reported, not guessed.

After refinement, rerun integrity/rule validation and render the changed candidate before delivery. Further repairs require new evidence and verified anchors; stop repeating an ineffective method after two attempts. Tell the user which concrete problems remain and which pages could not be checked.
