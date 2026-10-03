---
name: thesis-docx-formatter
description: 中文本科毕业论文排版，默认浙江农林大学 ZAFU 规范。保守修复 DOCX，或从 Markdown/TXT 生成 Word、LaTeX 工程与 PDF。处理标题目录、页码分节、图表、公式、摘要关键词和参考文献。适用于毕业论文、论文格式、Word 排版、Markdown 转论文、LaTeX 论文输出；不用于普通文档或代码任务。默认只做程序校验，不渲染、不转图片、不进行视觉审阅。
---

# 毕业论文排版 Skill

## 默认工作流

1. 确认输入路径与交付格式；未指定时使用 Word 后端。
2. 运行统一入口，只读取必要的配置与引用资料。
3. 阅读该次运行的 `reports/summary.json`，仅有问题时追查详细报告。
4. 交付 `deliverables/` 中的文件，说明剩余风险与未执行的检查。

```powershell
python scripts/thesis_format.py "论文.docx"
python scripts/thesis_format.py "论文.md" --backend latex
python scripts/thesis_format.py "论文.md" --backend both
```

需要 Python 3.11+ 和 `requirements.txt` 中的依赖，**不要求安装 uv**。优先使用已具备依赖的 Python 环境；缺依赖时说明原因并提供 `python -m pip install -r requirements.txt`，不要未经许可安装软件、修改全局环境或强制引入 uv。

从其他项目目录调用时，使用 Skill 脚本的绝对路径，并明确选择已安装依赖的 Python 解释器。安装依赖也使用该解释器的 `-m pip` 与 `requirements.txt` 绝对路径，避免安装到另一个环境。子任务沿用同一解释器。开发者可以继续使用 `uv sync --locked` 和 `uv run --no-sync python ...`；uv 是可选开发工具，不是用户运行前提。

**默认不要看图、导出全文页面、生成视觉修复计划或重新排版。** Word 默认不导出 PDF；LaTeX 的 PDF 是编译产物，不等于视觉审阅。只有用户明确要求时才启用额外检查。

## 安全与交付边界

- 不改写论文内容，不覆盖或移动原输入，不猜测姓名、学号等个人信息。
- 摘要局部整理、交叉引用转换、标题编号、隐藏属性处理、参考文献拆分和去数字标号属于已允许的排版处理，不逐项请求确认；不进行 AI 润色、补写或章节重排。
- 原始 DOCX 默认不整篇重建、不替换原有前置部分、不自动新增正文分节。需要这些结构操作时先询问操作者，获准后使用 `--allow-structural-rebuild`。Markdown/TXT 新建文档不受此限制。
- “重新从头生成”使用当前文本源及明确指定的学校原始模板，不从历史项目 Word 偷取封面、页眉或样式。新建文档的封面来源与哈希记录在 manifest；“封面不改”必须相对于明确的来源判断，不填入猜测的题目或个人信息。
- DOCX 默认优先保留 Word 原生对象；不静默删除图片、公式、脚注、批注或修订。
- 警告可以继续交付；Word 质量门的硬失败会阻止候选文件进入交付目录。
- LaTeX 首版只支持 Markdown/TXT；不承诺复杂 DOCX 无损转换。
- LaTeX 学校适配为草稿级：可显式传入已确认的学校前置页 PDF 原样插入；不自动重绘封面或填写诚信页，不证明完全符合学校或学院规范。参数见 `references/zafu_latex_adapter.md`。
- 缺编译环境时交付 TeX 工程，明确报告未编译；编译失败不交付 PDF。
- 结构检查、布局分析和图片生成均不能声称 Agent 已完成视觉审阅。

## 文件保存

默认输出在**输入文件旁**，不在 Skill 安装目录随意写文件：

```text
thesis-output/<输入文件名>/<唯一运行ID>/
  manifest.json
  deliverables/
    <项目名>__<profile>__word.docx
    <项目名>__<profile>__word.pdf          可选
    <项目名>__<profile>__latex-source.zip 可选
    <项目名>__<profile>__latex.pdf        编译成功时
  reports/summary.json
  reports/summary.html                   Word 路线
  work/                                 中间文件与可编辑 TeX 工程
```

`--output-dir` 指定输出**基目录**，仍为每次运行创建独立子目录。`--project-name` 指定项目名；默认取输入文件名而非正文推断的题目。交付文件路径以该次 `manifest.json` 为准，其产物路径相对运行目录。

LaTeX 提示编译器不可用时，可检查已经安装的运行环境（包括 Codex 插件自带的 Tectonic），仅为本次命令设置 PATH 后重试；不要因此自动安装大型 TeX 环境或修改全局环境。

ZAFU LaTeX 直接使用固定版本的外部 `ZafuThesis.cls`，首次联网下载到忽略缓存；源码 ZIP 自带原始 CLS 和来源记录。不将 Word 样式覆盖到 CLS。封面/声明宏只按显式元数据选项调用，双语摘要映射到原始摘要宏；不猜个人信息或伪造签名。缺少依赖时明确报告失败；分享含外部 CLS 的源码前确认授权。

Word 不做引文管理，不读取 BibTeX、不改写正文文献引用、不重排文献；保留既有条目拆分和去数字标号，统一字体并生成格式提示。LaTeX 可用显式 `.bib` 与 `[@key]` 管理 GB/T 7714-2015 著者—年份引用，中文排序需确认的拼音键。需要相关功能时按需读取以下专用参考。

ZAFU Word 页眉默认使用“浙江农林大学本科生毕业论文（设计）”。保守修复 DOCX 时仅修改可识别正文/摘要节已有的非空页眉，不新增缺失页眉；Markdown/TXT 新建时为明确的正文/摘要节补齐页眉，并创建独立的单一 PAGE 域页脚。目录节使用空页眉；新建目录与摘要分节，前置部分罗马页码接续，正文阿拉伯页码从1开始。共享页眉单独隔离，避免影响封面。目录与正文同节或节用途不明确时保留并提示。`--header-mode thesis-title` 使用论文题目方式，`--header-mode preserve` 保持页眉原样；外部 LaTeX CLS 不受这些选项影响。

## 显式选项

| 选项 | 行为 |
|---|---|
| `--backend word` | 默认 Word 修复或重建 |
| `--backend latex` | 从公共 IR 生成独立 TeX 工程与 PDF |
| `--backend both` | 共用一次语义提取，分别生成两种格式 |
| `--export-pdf` | Word 导出 PDF，不自动生成 PNG |
| `--allow-structural-rebuild` | 操作者明确批准 DOCX 整篇重建或模板前置部分替换后使用 |
| `--header-mode fixed\|thesis-title\|preserve` | Word 页眉固定文字、论文题目或保持原样 |
| `--header-text "文字"` | 自定义固定页眉，隐含 fixed 模式 |
| `--latin-font "Times New Roman"` | 显式覆盖 Word 的西文字体；不改受保护封面及 LaTeX 字体，中文字体和字号仍依学校规则 |
| `--check structural` | 默认结构、规则及内容/资产检查 |
| `--check layout` | 增加 PDF 文本与几何分析，不生成页面图片 |
| `--check visual` | 显式生成预览图片；不自动开展 Agent 审阅 |
| `--no-compile` | 仅生成 LaTeX 工程，不调用编译器 |
| `--latex-engine auto\|xelatex\|tectonic` | auto 优先 PATH 上的 XeLaTeX，再尝试 Tectonic；不自动安装 |
| `--latex-metadata metadata.json` | 仅 LaTeX 的封面字段、声明开关及摘要覆盖 |
| `--latex-bibliography references.bib` | 仅 LaTeX 的显式著者—年份引文管理；Word 不读取 |
| `--compile-timeout 120` | 单次编译超时秒数，范围 1–600 |
| `--profile zafu_2022` | 学校规则适配 |
| `--compliance strict-school` | Word 严格规范覆盖；不将 LaTeX 草稿升级为合规认证 |
| `--dry-run` | 创建独立运行记录，仅记录公共阶段命令，不生成文档 |

Word 专家模式仍兼容 `--mode audit-only|conservative-repair|rebuild`。LaTeX 与 Word 页码、字体回退和布局可能不同，不承诺逐页一致。

## 语义识别纠错

需要纠错时按 `work/thesis_ir.json` 中的稳定 blockId 编写覆盖文件：

```json
{"overrides": [{"blockId": "block-00042", "role": "heading_2", "reason": "这是二级标题"}]}
```

通过 `--semantic-overrides overrides.json` 应用。原始 DOCX 保守修复读取同一 IR 中的显式纠正，支持正文、标题、图表题注等可直接映射的角色；保持原始 OOXML 对象。找不到 blockId 时跳过并报告，不按数字猜测段落；不支持的角色或受保护前置区域同样保留。纠正文件仅用于对应的当前输入，不应跨修改后的文档复用。

先看 `reports/summary.json` 的 `semanticCorrections`、`deferredStructuralChanges` 和 `formattingReview`。不能把部分未处理的问题说成已修复；不增加逐块内容比对或额外视觉审阅。

Word 更新目录会重新生成结果文字，应在更新后核对目录实际字体及页码。仅检查样式或结构通过，不代表 Word 已完成目录更新、实际分页核对或视觉审阅。需要 Word 自动化时区分执行失败与关闭阶段 RPC 失败：后者只有在产物和关键结果确已核验后才可记录为清理警告；不要盲目重跑或保存文档，Word 保存可能重写封面 XML 和图片。

## 按需参考

- `references/backend_and_workspace_contract.md`：双后端、目录、状态和迁移规则。
- `references/zafu_latex_adapter.md`：外部 CLS 元数据、原始前置页宏和显式 BibTeX 用法。
- `references/zafu_reference_format.md`：Word 参考文献书写规则及不做引文管理的边界。
- `references/thesis_ir_contract.md`：公共 IR 与来源锚点。
- `references/optional_visual_review.md`：仅用户要求时读取的视觉审阅与修复。
- `references/ooxml_pitfalls.md`：修改 OOXML 前必读。
- `references/plan_schema.md`：Word 修复计划契约。
- `references/hybrid_rebuild_contract.md`：Word 内部混合重建，不是双后端。
- `references/architecture_deep_dive.md`：已有 Word 引擎实现细节；当前默认以本 Skill 和新契约为准。
