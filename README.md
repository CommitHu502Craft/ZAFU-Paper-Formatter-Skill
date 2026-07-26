<h1 align="center">浙江农林大学本科毕业论文智能排版 Skill</h1>

<h3 align="center">ZAFU Paper Formatter Skill</h3>

<p align="center">
  <strong>🎓 把论文文件交给 AI Agent,Skill 会自动排版、渲染检查并输出新的 Word 和 PDF,原文件不会被覆盖。</strong>
</p>

这是一个面向**浙江农林大学本科生**的毕业论文智能排版 Skill。它可以由 Claude Code、Codex、Cursor 等能够访问本地文件和运行命令的 AI Agent 使用。你只需要提供论文文件(Word `.docx`、Markdown `.md` 或纯文本 `.txt`),Agent 会:

1. 自动排版:标题层级、目录、页码分节、三线表、图题表题、公式、摘要关键词、参考文献、中英文字体字号;
2. 渲染 PDF 并生成全文缩略图和关键页预览,像人一样"翻一遍"检查页面效果;
3. 发现明显问题时自动做一次视觉修复(如标题落在页尾、表头跨页不重复);
4. 输出排版后的 Word、PDF 和一份看得懂的中文检查报告。

不会 Python、没用过命令行都没关系——把文件路径和下面的提示词发给 Agent 即可。

## 🚀 安装

打开一个能访问本地文件的 AI Agent,发送:

```text
请从下面的 GitHub 地址安装 ZAFU Paper Formatter Skill:
https://github.com/CommitHu502Craft/ZAFU-Paper-Formatter-Skill

请自行检查运行环境并完成依赖安装。优先使用 uv 管理 Python 环境,不要修改我的论文文件。安装完成后运行 Skill 自带的健康检查,并告诉我安装位置和检查结果。
```

如果本机装有 LibreOffice、Microsoft Word 或 WPS,会自动用于 PDF 渲染检查;都没有也不影响生成 Word 文件。

## 💬 使用(推荐提示词)

把论文文件的完整路径和下面这段话发给已安装 Skill 的 Agent:

```text
请使用 ZAFU Paper Formatter Skill 排版以下本科毕业论文:
D:\毕业论文\我的论文.docx

要求:
1. 按浙江农林大学本科毕业论文规范执行默认自动排版;
2. 排版完成后检查页面效果:如果你能读取图片,查看全文缩略图(contact sheet)
   和关键页面预览;如果你不能读取图片,不要读图,改读
   output/debug/visual_review_manifest.json 里的文字检查结果,
   并把图片路径告诉我由我自己翻看;
3. 发现明显排版问题时,生成视觉修复计划并执行一次二次修复;
4. 输出最终的 Word、PDF 和检查报告,并把报告里的人工检查清单转述给我;
5. 不要改写论文正文的任何文字;
6. 不要覆盖我的原文件;
7. 不要因为少量警告就停止交付,把剩余风险写进报告即可。
```

Markdown / TXT 初稿同样适用,只需替换文件路径。

### 输出在哪里

排版完成后,`output/final/` 里就是你需要的全部文件:

| 文件 | 说明 |
|---|---|
| `repaired.docx` | 排版完成、可继续编辑的 Word |
| `repaired.pdf` | 渲染预览 |
| `review_report.html` | 中文检查报告(双击用浏览器打开) |
| `contact_sheet.png` | 全文页面缩略图 |
| `critical_pages/` | 封面、目录、摘要、正文首页等关键页大图 |

## ⚠️ 使用前注意

- 原论文不会被覆盖,但排版前请自行备份论文和图片附件;
- 封面与诚信页中的姓名、学号等个人信息需要你自己核对填写,工具不会猜;
- 拿到 Word 后请在 Word/WPS 中右键目录 → "更新整个目录" 刷新页码;
- 浮动图片、文本框、SmartArt、嵌入式 Excel 等复杂对象会保留原样并列入人工检查清单;
- 最终提交前务必按 `review_report.html` 里的清单人工检查一遍。

## 🧑‍💻 开发者与高级用户

手动安装:

```powershell
git clone https://github.com/CommitHu502Craft/ZAFU-Paper-Formatter-Skill.git
cd ZAFU-Paper-Formatter-Skill
uv sync
```

统一入口命令:

```powershell
uv run python scripts\thesis_format.py "D:\毕业论文\我的论文.docx" --profile zafu_2022 --output-dir output
```

可选参数:

- `--semantic-overrides overrides.json` — 用稳定 blockId 覆盖语义识别结果(标题层级、图题、关键词等);
- `--visual-refinement-plan plan.json` — 执行一次白名单视觉修复(keepNext、另起一页、表头重复、图片缩放等)并重新渲染;
- `--mode audit-only|conservative-repair|rebuild`、`--compliance default|strict-school`、`--dry-run` — 专家参数。

健康检查与测试:

```powershell
uv run python scripts\verify_skill_health.py
uv run python -m unittest tests\test_unified_thesis_ir.py
```

技术细节(ThesisIR 契约、OOXML 陷阱、修复计划 schema、架构说明)见 `SKILL.md` 与 `references/` 目录。

## 🧭 项目状态

当前版本以"尽量一次交付可用结果"为目标:自动执行确定性的排版修复,渲染 PDF 做视觉复查,无法可靠处理的项目写入人工检查清单而不是中断流程。

仓库中的 ZAFU profile 和模板相关资源用于论文排版研究及个人学术用途。进行更广泛的转载或再分发前,请确认学校模板的相关要求。

## 📜 许可证

本项目的源代码、Skill 指令、配置和原创文档采用 [Apache License 2.0](LICENSE)。

浙江农林大学官方论文模板及其派生资源不属于 Apache-2.0 授权范围,其权利归相应权利人所有,详见 [NOTICE](NOTICE)。本项目是独立社区项目,不是浙江农林大学官方项目,也不代表学校的认可或维护。
