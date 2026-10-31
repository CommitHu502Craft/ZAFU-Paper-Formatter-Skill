---
name: thesis-docx-formatter
description: 中文本科毕业论文排版，默认浙江农林大学 ZAFU 规范。批量规范 DOCX，保留复杂原生对象，或从 Markdown/TXT 生成 Word、LaTeX 工程与 PDF。普通排版按正式交付流程生成预览并进行重点验收，快速草稿可跳过渲染。不用于普通文档或代码任务。
---

# 毕业论文排版 Skill

## 默认工作流

1. 确认输入路径、学校要求和交付格式；未指定时使用 Word 后端。以最终论文是否实际排版到位为完成依据，不能以命令结束、报告通过或某个状态标记代替成品核对。
2. DOCX 先运行 `scripts/inspect_requirements.py "论文.docx" --output "工作目录/缺项.json"`。逐项检查封面、诚信承诺书、目录、中英文题目/摘要/关键词、正文页眉、页码、参考文献；向用户简要列出实际缺项及准备如何处理。正文没有的信息不得编造；模板封面的空白字段不算缺项。
3. 普通请求使用默认 `--quality final`，明确要求快速草稿时用 `--quality fast`。原稿缺少学校要求的模板前置页时，按下方“前置页补齐”处理，不能直接交付缺封面的默认保守输出。缺失正文页眉应创建，空页眉应补上规定文字。
4. 核对该次成品与缺项清单，按 [视觉验收说明](references/optional_visual_review.md) 实际查看总览、全部图表页以及高清关键/异常页，核对实际字体、分页、目录、页眉页码、图题和参考文献。发现问题先定位再修复；每次修改后重新渲染受影响页面，不能沿用旧截图。修复产生新问题时继续有针对性地处理；连续两次同类修复无改善时停止该方法，说明具体剩余问题。
5. 交付最终论文文件，简要说明已完成内容与仍存在的具体缺项。内部状态和审阅记录用于追溯，不要求用户理解或操作“标为正式完成”。仍缺必需部分或存在已知格式问题时，必须如实列出，不能只说排版完成。

```powershell
python scripts/thesis_format.py "论文.docx"
python scripts/thesis_format.py "论文.docx" --quality fast
python scripts/thesis_format.py "论文.md" --backend latex
python scripts/thesis_format.py "论文.md" --backend both
```

需要 Python 3.11+ 和 `requirements.txt` 中的依赖，**不要求安装 uv**。优先使用已具备依赖的 Python 环境；缺依赖时说明原因并提供 `python -m pip install -r requirements.txt`，不要未经许可安装软件、修改全局环境或强制引入 uv。

从其他项目目录调用时，使用 Skill 脚本的绝对路径，并明确选择已安装依赖的 Python 解释器。安装依赖也使用该解释器的 `-m pip` 与 `requirements.txt` 绝对路径，避免安装到另一个环境。子任务沿用同一解释器。开发者可以继续使用 `uv sync --locked` 和 `uv run --no-sync python ...`；uv 是可选开发工具，不是用户运行前提。

Final 默认请求 PDF、低清总览和关键/异常页高清图。CLI 只完成程序检测与预览，不能声称 Agent 看过图；Agent 必须实际查看。审阅记录在成品核对后由 Agent 自行处理，不能把记录操作当成用户的额外任务。渲染或审阅未完成时说明哪些版面尚未核实。Fast 默认不渲染；LaTeX 编译产物同样不代表视觉审阅。

### 前置页补齐

- **封面和诚信页只能从明确的学校原始模板原样插入，禁止修改或填写。** 题目、年份、学院、姓名、学号、专业、导师、日期、签名、下划线、空格、样式、图片和分页均属于保护范围；即使能从正文推断也不得填写。空白占位按模板保留，不列为待补个人信息。
- 原稿两页均缺失时，普通“按该学校规范排版”包含补齐所需模板页的授权。使用已核验模板，经 `--allow-structural-rebuild --mode conservative-repair` 的关系感知组合路径补齐，并复核正文、公式、图片、表格和模板前两页均保留；这不是重写封面的授权。该参数会进入专家兼容链，不能只看参数名就假定安全。
- 原稿已有封面/诚信页时保护原页；只缺其中一页时仅插入缺页，不能以整套模板替换已有页。当前 CLI 不提供通用的单页插入接口，需要按模板结构做定向 OOXML 合并及关系/样式核验；无法保证原样时如实说明具体障碍，不擅自重画。
- 缺目录应建立并更新真实目录域，检查目录文字与正文标题、页码对应。目录域会标记为需更新并启用“打开时更新域”，Word 打开文档后自动刷新页码；仅用不更新域的查看器打开时看到的是缓存结果。模板封面页不得受到正文页眉、页脚或共享样式的连带修改。
- 无编号的“引言/前言/绪论”按一级标题处理并进入目录，正文页码从该页起算，其后的编号章仍从“1”开始；章标题默认为每章另起一页。
- 插入前记录模板来源和哈希，插入后比对前两页 XML 子树及依赖资产，并查看渲染结果。Word 保存可能重写包内容；若保护区域发生变化，应恢复原始模板内容后重新核验，不能仅凭“看起来一样”通过。

## 安全与交付边界

- 不改写论文内容，不覆盖或移动原输入，不猜测姓名、学号等个人信息。
- 普通 DOCX 批量调整安全区域的格式，保留正文文字、强调和上下标；复杂区域保持原样并提示。旧式公式文本转换、编号和参考文献拆分保留在专家兼容链中；不进行润色、补写或章节重排。
- 原始 DOCX 默认不整篇重建、不替换已有前置页。补齐缺失的学校标准前置页和为页眉页码建立必要的隔离分节属于普通排版范围，按上述规则处理；只有无法从既有要求确定的内容或会改动现有受保护页的操作才需澄清。
- “重新从头生成”使用当前文本源及明确指定的学校原始模板，不从历史项目 Word 偷取封面、页眉或样式。新建文档的封面来源与哈希记录在 manifest；“封面不改”必须相对于明确的来源判断，不填入猜测的题目或个人信息。
- DOCX 默认优先保留 Word 原生对象；不静默删除图片、公式、脚注、批注或修订。
- 警告可以继续交付；Word 质量门的硬失败会阻止候选文件进入交付目录。
- LaTeX 首版只支持 Markdown/TXT；不承诺复杂 DOCX 无损转换。
- LaTeX 学校适配为草稿级：可显式传入已确认的学校前置页 PDF 原样插入；不自动重绘封面或填写诚信页，不证明完全符合学校或学院规范。参数见 `references/zafu_latex_adapter.md`。
- 本工作流的 LaTeX 封面/诚信页同样使用明确来源的原页 PDF 原样插入，不调用宏重画或填写封面。已有宏与元数据选项不构成违反封面保护规则的理由。
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

ZAFU LaTeX 直接使用固定版本的外部 `ZafuThesis.cls`，首次联网下载到忽略缓存；源码 ZIP 自带原始 CLS 和来源记录。不将 Word 样式覆盖到 CLS。封面/声明遵守原页插入规则，双语摘要映射到原始摘要宏；不猜个人信息或伪造签名。缺少依赖时明确报告失败；分享含外部 CLS 的源码前确认授权。

Word 不做引文管理，不读取 BibTeX、不改写正文文献引用、不重排文献。先识别实际条目及引用制度，再统一标题、字体、字号和缩进；标题兼容“参考文献：”“参 考 文 献”和 References。正文有数字引用时保留对应文后编号；实际制度与学校规则不同要明确列出，禁止只删文后编号造成引用失联。疑似存在文献却提取到零条必须报识别失败；逐条检查类型标识、年份、编号连续性与可识别的数字引用目标，不猜补题录。详见 `references/zafu_reference_format.md`。

ZAFU Word 页眉默认使用“浙江农林大学本科生毕业论文（设计）”。DOCX 与新建文档均为可识别正文/摘要节创建缺失页眉、补齐空页眉；首页不同、奇偶页不同时检查相应变体。目录节使用空页眉，封面/诚信页保持模板原样；共享页眉必须隔离。缺失页脚时在明确的非封面节创建单一 PAGE 域，现有页码核对显示及接续。新建目录与摘要分节，前置部分罗马页码接续，正文阿拉伯页码从1开始。目录与正文同节或节用途不明确时，列出尚未补齐的页眉及分节原因并继续定向处理，不能把跳过当作符合要求。`--header-mode thesis-title` 使用论文题目方式，`--header-mode preserve` 按用户明确要求保持页眉；外部 LaTeX CLS 不受这些选项影响。

## 显式选项

| 选项 | 行为 |
|---|---|
| `--backend word` | 默认 Word 修复或重建 |
| `--backend latex` | 从公共 IR 生成独立 TeX 工程与 PDF |
| `--backend both` | 共用一次语义提取，分别生成两种格式 |
| `--quality final\|fast` | 默认 final 请求渲染与重点验收；fast 快速交付 |
| `--conservative` | 显式要求保守处理；不允许结构重建 |
| `--render-backend auto\|word\|libreoffice\|wps` | Windows auto 优先 Word，其次 WPS、LibreOffice；显式后端不自动切换 |
| `--export-pdf` | 在 fast/structural 中也请求 PDF；图片仍由 check 决定 |
| `--allow-structural-rebuild` | 进入专家结构组合路径；补齐缺失模板页时可按上述范围使用，禁止替换或改写已有受保护页 |
| `--header-mode fixed\|thesis-title\|preserve` | Word 页眉固定文字、论文题目或保持原样 |
| `--header-text "文字"` | 自定义固定页眉，隐含 fixed 模式 |
| `--latin-font "Times New Roman"` | 显式覆盖 Word 的西文字体；不改受保护封面及 LaTeX 字体，中文字体和字号仍依学校规则 |
| `--check structural` | 显式覆盖 quality，仅做结构、规则及内容/资产检查 |
| `--check layout` | 增加 PDF 文本与几何分析，不生成页面图片 |
| `--check visual` | Final 默认检查层级；生成图片，实际审阅另行记录 |
| `--pipeline legacy` | 专家兼容旧审计/预检报告链，普通排版不需要 |
| `--no-compile` | 仅生成 LaTeX 工程，不调用编译器 |
| `--latex-engine auto\|xelatex\|tectonic` | auto 优先 PATH 上的 XeLaTeX，再尝试 Tectonic；不自动安装 |
| `--latex-metadata metadata.json` | LaTeX 摘要覆盖及原始前置页 PDF 设置；本工作流禁止使用填写/重画封面的选项 |
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

先看 `reports/summary.json` 的 `semanticCorrections`、`preservedRegions`、`visualAcceptance` 和 `formattingReview`。不能把未处理的问题说成已修复。学校 AI 使用披露是独立要求，工具不会代写或伪造声明。

Word 更新目录会重新生成结果文字，应在更新后核对目录实际字体及页码。仅检查样式或结构通过，不代表 Word 已完成目录更新、实际分页核对或视觉审阅。需要 Word 自动化时区分执行失败与关闭阶段 RPC 失败：后者只有在产物和关键结果确已核验后才可记录为清理警告；不要盲目重跑或保存文档，Word 保存可能重写封面 XML 和图片。

## 按需参考

- `references/backend_and_workspace_contract.md`：双后端、目录、状态和迁移规则。
- `references/zafu_latex_adapter.md`：外部 CLS 元数据、原始前置页宏和显式 BibTeX 用法。
- `references/zafu_reference_format.md`：Word 参考文献书写规则及不做引文管理的边界。
- `references/thesis_ir_contract.md`：公共 IR 与来源锚点。
- `references/optional_visual_review.md`：Final 的重点视觉审阅与有界修复。
- `references/formatting_first_architecture.md`：内部路线、缓存与兼容边界。
- `references/ooxml_pitfalls.md`：修改 OOXML 前必读。
- `references/plan_schema.md`：Word 修复计划契约。
- `references/hybrid_rebuild_contract.md`：Word 内部混合重建，不是双后端。
- `references/architecture_deep_dive.md`：已有 Word 引擎实现细节；当前默认以本 Skill 和新契约为准。
