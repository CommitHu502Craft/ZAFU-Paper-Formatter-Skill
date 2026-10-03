# ZAFU 外部 CLS 适配

LaTeX 排版来源：[Stolorzs/ZafuTemplatePublic](https://github.com/Stolorzs/ZafuTemplatePublic)。文档类固定到上游提交 `81c0d127d485aaaf00c5cbccc5d0d29affd446f2`，不将其描述为本项目原创模板。下载资源通过 SHA256 校验，生成工程记录 `UPSTREAM.json`；外部 CLS 和图片的再分发许可尚未确认，标注来源不代表获得授权，详见 [NOTICE](../NOTICE)。

本路线直接使用固定版本原始 `ZafuThesis.cls`，不依赖 Word 封面和页眉页脚实现。上游 CLS 不修改；字体族兼容设置和字段映射在独立适配模板完成。

## 封面与前置页

```powershell
python scripts/thesis_format.py '论文.md' --backend latex --latex-metadata 'metadata.json'
```

`metadata.json` 是显式配置，不从姓名、学号或文件名猜造个人信息：

```json
{
  "title": "论文中文题目",
  "titleEn": "English Thesis Title",
  "author": "实际作者",
  "academy": "实际学院",
  "session": "2026",
  "studentNumber": "实际学号",
  "className": "实际专业班级",
  "teacher": "实际指导教师",
  "teacherTitle": "实际职称",
  "date": "实际提交日期",
  "cover": true,
  "statement": false
}
```

- `cover: true` 调用原始 `\customCover`；上述封面字段缺失时阻止 PDF 交付。校徽/校名图片固定版本下载、校验、复制进工程并记录来源。
- `statement: true` 才调用原始 `\makestatement`；必须有题目、作者、日期，签名位置为空。不自动勾选承诺，不伪造签名。
- 默认 `cover`、`statement` 均关闭；日期不自动使用当天日期。
- 需要保留学校原始未填写封面、诚信页时，可显式提供 `"frontMatterPdf": "front-matter.pdf"`，路径相对元数据文件解析。适配器校验非空 PDF、复制并记录 SHA256，以 `pdfpages` 原样插入，不填写个人信息、不重绘页面。此方式不得同时启用 `cover` 或 `statement` 宏。
- 源文档中独立的中文/英文摘要标题、摘要段落和关键词分别映射到 `\ZhAbstract`、`\EnAbstract`，随后调用类的 `\customContent` 和 `\mainmatter`。不重复输出已映射的源块。
- 支持标题 `摘要`、`中文摘要`、`Abstract`、`English Abstract`、`英文摘要`，以及独立段落 `摘要：内容`。关键词使用 `关键词：`、`Keywords:` 等标签。
- 在 Markdown/TXT 中用空行分隔摘要、关键词和正文段落。若公共 IR 已将摘要和正文合并为一个块，保持该块并提示人工拆分，不冒险删掉正文。
- 可用 `abstractCn`、`keywordsCn`、`abstractEn`、`keywordsEn` 显式覆盖摘要内容；摘要不自动翻译，英文题目不自动编造。正文语义仍以共享 IR 为准。
- 该适配与 Word 完全隔离；`--backend both` 时这些元数据只传入 LaTeX。

## 引文管理（仅 LaTeX）

```powershell
python scripts/thesis_format.py '论文.md' --backend latex --latex-metadata 'metadata.json' --latex-bibliography 'references.bib'
```

显式引用标记：`[@chen1989]`、`[@chen1989; @reich2006]`，转换为 `\citep{...}`。已有作者—年份文字或 `[1]` 不会被猜测为某个键。缺少 `.bib` 或键不存在时阻止编译，代码示例中的标记不执行。

```bibtex
@book{chen1989,
  author = {陈守常 and 曾大鹏},
  title = {油茶病害及其防治},
  address = {北京},
  publisher = {中国林业出版社},
  year = {1989},
  key = {chen shou chang zeng da peng},
  language = {chinese}
}
```

中文作者必须在 BibTeX 的 `key` 字段提供经过确认的拼音排序键。这个字段不同于记录标识 `chen1989`。不猜多音姓氏、不补作者/年份。外文作者推荐 `Surname, Given Names and ...`，由样式输出姓全大写及无点缩写名。

使用固定 `gbt7714` v2.1.4 的 **2015 著者—年份样式**，不跟随环境默认切换到其他年份标准。原始 BST、源文件、许可和校验信息随工程保存；生成独立命名的学校 BST 变体，仅修改“年份不紧跟作者”和“作者后用句号”，以匹配用户给出的文后书写范例。CLS 保持未修改。正文圆括号、分号、同作者压缩和同年 a/b 由 natbib/BibTeX 处理。

显式 `.bib` 替代 LaTeX 的原文字面参考文献段，原文仍保留在输入与 IR；只输出被引用的记录，不自动 `\nocite{*}`。无 `.bib` 时保留原文字面参考文献。

Tectonic 内置运行 BibTeX；XeLaTeX 路线检测 `bibtex`，执行 TeX → BibTeX → TeX 重复轮次，并检查未解析引用。缺少 BibTeX、失败或引用未解析时不交付 PDF。

**Word 不读取此 `.bib`，也不处理 `[@key]`。** `both` 模式中的 Word 保留这些标记并给出人工检查提示；需要 Word 成稿时请提供已经写好的作者—年份正文引用，不要把 LaTeX 专用标记当成 Word 引文管理接口。

默认没有视觉审阅。编译成功不等于学校合规认证；用户授权后用 `--check visual` 导出页面并实际查看。

## 有界排版适配

保持固定版本 CLS 不变。英文摘要及关键词在宏参数内部显式选择10.5磅/15磅行距，避免上游5磅设置在斜体切换时生效。此补偿仅用于该类的英文前置内容。

ZAFU `latex.yaml` 设置 `inline_code_font: roman` 与 `break_inline_code: true`，代码参数沿用 Times New Roman，并在标点和每6字符处允许无可见连字符的断行；命令开头的双连字符不拆开，不删改参数或命令。长的裸英文标识和阶段比较名同样允许断行。普通适配器默认仍为等宽字体。英文摘要使用适配层页眉页脚样式，保持与中文摘要一致，但不重置摘要页码。

表格使用三线表，`table_alignment: center` 表示单元格水平居中，垂直居中并适度增加行高；列宽按内容长度有限调整，不通过整体缩字挤入页面。`wide_table_columns: 8` 仅将达到阈值的宽表置于独立横向页。表前标题和图后图注与对应对象绑定，不重复输出。

`emergency_stretch_em: 2` 为含大量英文标识的中文段落保留有限的行内间距弹性，仅在正常断行失败时参与选择，避免文字越过版心；不改变字号、固定行距或可打印边距。斜线连接的亚家族标签允许在分隔处断行。

额外断行只针对带数字、下划线、分隔符或驼峰结构的长标识；普通英文单词保留 TeX 自身的语言连字断行。短的 `log2(TPM+1)` 一类表达式不在内部强行拆开。

代码参数中的双连字符按字面字符输出，避免 Roman 字体的 TeX 文本连字将 `--paired-end` 变成长横线；参数引号保留源字符，不改写指令。此处理仅限代码标记，不关闭普通论文文字的正常排印规则。

裸网址识别在中文括号和书名号前结束，不将后面的访问日期等中文说明吞入西文字体的 `\url`。检查编译日志中的 `Missing character`，不能以编译成功替代缺字检查。

图片以 `image_width_fraction`、`image_height_fraction` 限制可打印宽高，保留宽高比和原始资产字节，不裁切、不重绘科学图件。当前ZAFU配置为满正文宽度、最多正文高度的0.82；实际字号仍取决于原图，须查看页面决定是否需要拆分。
