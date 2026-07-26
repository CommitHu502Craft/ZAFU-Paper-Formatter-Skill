#!/usr/bin/env python3
"""Generate a static Chinese-language review report (review_report.html).

Framework-free HTML aimed at undergraduate students: what was done, what
remains, and which pages to eyeball. Reads whatever reports exist in the
output directory; every input is optional so partial runs still get a report.
"""
from __future__ import annotations

import argparse
import html
import json
from pathlib import Path
from typing import Any, Dict, List, Optional


def read_json(path: Path) -> Dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


ACTION_LABELS = {
    "apply_body_style": "统一正文样式",
    "apply_heading_styles": "统一各级标题样式",
    "apply_section_geometry": "修正页边距和版心",
    "apply_page_numbering": "整理页码和分节",
    "apply_header_footer": "整理页眉页脚",
    "normalize_tables": "表格整理(三线表等)",
    "normalize_captions": "图题表题整理",
    "normalize_references": "参考文献排版",
    "normalize_abstract": "摘要与关键词排版",
    "ensure_toc": "目录整理",
}

RISK_LABELS = {"A": "A(接近模板,低风险)", "B": "B(常规论文,中风险)", "C": "C(结构混乱,高风险)"}
STRATEGY_LABELS = {
    "preserve_first": "保留优先(在原文档上做保守修复)",
    "template_overlay": "模板覆盖(保留正文,套用模板样式和页面规则)",
    "hybrid_rebuild": "混合重建(以规范结构为底座重建正文并回挂图表)",
    "text_rebuild": "文本重建(从 Markdown/TXT 生成全新 Word)",
    "audit_only": "仅审计(文件问题过多,未自动修改)",
}


def _issue_text(issue: Any) -> str:
    if isinstance(issue, dict):
        kind = str(issue.get("type") or issue.get("kind") or "问题")
        message = issue.get("message") or issue.get("expected") or issue.get("reason") or ""
        return f"{kind} {message}".strip()
    return str(issue)


def collect(output_dir: Path) -> Dict[str, Any]:
    manifest = read_json(output_dir / "dispatch_manifest.json")
    validation = read_json(output_dir / "validation_report.json")
    review = read_json(output_dir / "review_summary.json")
    render = read_json(output_dir / "render_validation" / "render_validation_report.json")
    if not render:
        render = read_json(output_dir / "render_validation_report.json")
    semantic = read_json(output_dir / "render_validation" / "semantic_pages.json")
    repair_execution = read_json(output_dir / "repair_execution.json")
    refinement = read_json(output_dir / "visual_refinement_execution.json")
    asset_preservation = read_json(output_dir / "asset_preservation.json")

    quality_gate = validation.get("qualityGate") or {}
    audit_summary = validation.get("auditSummary") or {}
    diff = validation.get("diffReport") or {}
    preflight_decision = manifest.get("preflightDecision") or {}

    auto_fixed = sorted({
        str(item.get("action"))
        for item in (repair_execution.get("executionLog") or [])
        if isinstance(item, dict) and item.get("action")
    })

    warnings = [_issue_text(item) for item in (quality_gate.get("warnings") or [])][:15]
    hard_failures = [_issue_text(item) for item in (quality_gate.get("hardFailures") or [])][:10]
    visual_findings = render.get("visualFindings") or []

    manual_checks: List[str] = []
    if render.get("status") != "ok":
        manual_checks.append("本机没有可用的 PDF 渲染工具,请自行打开 Word 检查整体版面。")
    manual_checks.append("在 Word 中打开后按 Ctrl+A 再按 F9(或右键目录→更新域→更新整个目录)刷新目录页码。")
    if any(f.get("type") == "key_region_not_located" for f in visual_findings):
        missing = [f.get("region") for f in visual_findings if f.get("type") == "key_region_not_located"]
        manual_checks.append(f"以下关键区域未能自动定位,请人工确认是否存在:{', '.join(str(m) for m in missing)}。")
    if any(f.get("type") == "near_blank_page" for f in visual_findings):
        pages = [str(f.get("page")) for f in visual_findings if f.get("type") == "near_blank_page"]
        manual_checks.append(f"第 {', '.join(pages)} 页接近空白,请确认是否为多余分页。")
    if any(f.get("type") == "heading_orphan_at_page_bottom" for f in visual_findings):
        manual_checks.append("有标题落在页面底部,建议在 Word 中检查对应章节的分页。")
    if manifest.get("pendingConfirmationRequests"):
        manual_checks.append("预检发现需要人工确认的项目(见 debug/preflight_report.json),已按默认策略继续,请核对。")
    for warning in (asset_preservation.get("warnings") or [])[:3]:
        manual_checks.append(f"内容保留检查:{warning}")
    manual_checks.append("核对封面和诚信承诺书上的姓名、学号、学院、指导教师等个人信息(工具不会自动填写)。")
    manual_checks.append("最后通读一遍摘要、关键词和参考文献,确认文字内容没有被改动。")
    manual_checks = manual_checks[:10]

    return {
        "input": manifest.get("input"),
        "profile": manifest.get("profile"),
        "riskClass": preflight_decision.get("documentRiskClass"),
        "strategy": (manifest.get("strategySelection") or {}).get("chosenStrategy"),
        "qualityGatePassed": quality_gate.get("passed"),
        "hardFailures": hard_failures,
        "warnings": warnings,
        "autoFixed": auto_fixed,
        "refinementApplied": refinement.get("appliedCount"),
        "pageCount": render.get("pageCount"),
        "pdfPath": render.get("pdfPath"),
        "contactSheets": render.get("contactSheets") or [],
        "criticalPages": render.get("criticalPages") or [],
        "regions": (semantic.get("regions") or {}),
        "visualFindings": visual_findings,
        "manualChecks": manual_checks,
        "assetPreservation": asset_preservation,
        "assetCounts": {
            "before": (diff.get("before") or {}).get("summary") or {},
            "after": (diff.get("after") or {}).get("summary") or audit_summary,
        },
    }


def _esc(value: Any) -> str:
    return html.escape(str(value if value is not None else "—"))


def _rel(path_value: Optional[str], base: Path) -> Optional[str]:
    if not path_value:
        return None
    try:
        return Path(path_value).resolve().relative_to(base.resolve()).as_posix()
    except ValueError:
        return Path(path_value).as_posix()


def render_html(data: Dict[str, Any], final_dir: Path) -> str:
    def li(items: List[str]) -> str:
        if not items:
            return "<li>无</li>"
        return "".join(f"<li>{_esc(item)}</li>" for item in items)

    contact_imgs = ""
    for sheet in data["contactSheets"]:
        rel = _rel(sheet, final_dir) or sheet
        contact_imgs += f'<a href="{_esc(rel)}"><img src="{_esc(rel)}" alt="全文缩略图" style="max-width:100%;border:1px solid #ccc;margin:6px 0"></a>'

    critical_imgs = ""
    for entry in data["criticalPages"]:
        img = entry.get("imagePath")
        rel = _rel(img, final_dir)
        if not rel:
            continue
        labels = "、".join(str(l) for l in (entry.get("labels") or []) if not str(l).endswith("_ctx")) or "上下文页"
        critical_imgs += (
            f'<figure style="display:inline-block;margin:8px;text-align:center">'
            f'<a href="{_esc(rel)}"><img src="{_esc(rel)}" style="height:260px;border:1px solid #ccc"></a>'
            f'<figcaption style="font-size:12px">第 {_esc(entry.get("page"))} 页 · {_esc(labels)}</figcaption></figure>'
        )

    regions_rows = ""
    region_names = {
        "cover": "封面", "integrity_statement": "诚信承诺书", "toc": "目录",
        "cn_abstract": "中文摘要", "en_abstract": "英文摘要", "body_start": "正文首页",
        "references": "参考文献", "acknowledgements": "致谢", "appendix": "附录", "last_page": "末页",
    }
    for key, region in (data["regions"] or {}).items():
        method = {"text_marker": "按页面文字定位", "first_page": "首页", "last_page": "末页", "fallback_estimate": "估计(未找到标志文字)"}.get(str(region.get("method")), str(region.get("method")))
        regions_rows += f"<tr><td>{_esc(region_names.get(key, key))}</td><td>第 {_esc(region.get('page'))} 页</td><td>{_esc(method)}</td></tr>"

    findings_rows = ""
    for finding in (data["visualFindings"] or [])[:20]:
        findings_rows += f"<tr><td>{_esc(finding.get('type'))}</td><td>{_esc(finding.get('page') or finding.get('pages') or finding.get('region'))}</td></tr>"

    verdict = "✅ 可以交付,请完成下方人工检查项" if not data["hardFailures"] else "⚠️ 已生成结果,但存在需要处理的硬性问题"

    asset = data.get("assetPreservation") or {}
    asset_rows = ""
    if asset:
        before, after = asset.get("before") or {}, asset.get("after") or {}
        labels = {"mediaFiles": "图片/媒体文件", "tables": "表格", "equations": "公式", "drawings": "图形对象", "textChars": "正文字符数"}
        for key, label in labels.items():
            b, a = before.get(key), after.get(key)
            mark = "✅" if (b or 0) <= (a or 0) else "⚠️"
            asset_rows += f"<tr><td>{label}</td><td>{_esc(b)}</td><td>{_esc(a)}</td><td>{mark}</td></tr>"

    risk = RISK_LABELS.get(str(data.get("riskClass")), _esc(data.get("riskClass")))
    strategy = STRATEGY_LABELS.get(str(data.get("strategy")), _esc(data.get("strategy")))
    fixed_items = [ACTION_LABELS.get(item, item) for item in data["autoFixed"]]

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>论文排版检查报告</title>
<style>
body {{ font-family: "Microsoft YaHei", "PingFang SC", sans-serif; max-width: 960px; margin: 24px auto; padding: 0 16px; color: #222; line-height: 1.7; }}
h1 {{ font-size: 22px; border-bottom: 2px solid #2c6e49; padding-bottom: 8px; }}
h2 {{ font-size: 17px; margin-top: 28px; color: #2c6e49; }}
table {{ border-collapse: collapse; width: 100%; font-size: 14px; }}
td, th {{ border: 1px solid #ddd; padding: 6px 10px; text-align: left; }}
.verdict {{ font-size: 16px; padding: 10px 14px; background: #f0f7f2; border-left: 4px solid #2c6e49; margin: 16px 0; }}
.warn {{ background: #fff8ec; border-left-color: #d9822b; }}
ul {{ padding-left: 22px; }}
.meta td:first-child {{ width: 160px; color: #666; }}
</style>
</head>
<body>
<h1>论文排版检查报告</h1>
<div class="verdict {'warn' if data['hardFailures'] else ''}">{verdict}</div>

<h2>基本信息</h2>
<table class="meta">
<tr><td>输入文件</td><td>{_esc(data['input'])}</td></tr>
<tr><td>使用规范</td><td>{_esc(data['profile'])}</td></tr>
<tr><td>文档风险等级</td><td>{risk}</td></tr>
<tr><td>处理策略</td><td>{strategy}</td></tr>
<tr><td>渲染页数</td><td>{_esc(data['pageCount'])}</td></tr>
<tr><td>最终 Word</td><td>repaired.docx(本目录)</td></tr>
<tr><td>最终 PDF</td><td>{'repaired.pdf(本目录)' if data['pdfPath'] else '未生成(本机缺少渲染工具,不影响 Word 文件)'}</td></tr>
</table>

<h2>自动完成的排版修复</h2>
<ul>{li(fixed_items)}</ul>

{f'<h2>内容保留检查(排版前 → 排版后)</h2><table><tr><th>项目</th><th>排版前</th><th>排版后</th><th>状态</th></tr>{asset_rows}</table>' if asset_rows else ''}

<h2>关键页面定位(按页面文字识别)</h2>
<table><tr><th>区域</th><th>页码</th><th>定位方式</th></tr>{regions_rows or '<tr><td colspan=3>未渲染 PDF,无页面定位</td></tr>'}</table>

<h2>关键页面预览</h2>
{critical_imgs or '<p>无预览图(本机缺少 PDF 渲染或截图工具)。</p>'}

<h2>全文缩略图</h2>
{contact_imgs or '<p>未生成 contact sheet。</p>'}

<h2>视觉检查发现(均为提醒,不阻止交付)</h2>
<table><tr><th>类型</th><th>位置</th></tr>{findings_rows or '<tr><td colspan=2>未发现明显视觉问题</td></tr>'}</table>

<h2>仍存在的警告</h2>
<ul>{li(data['warnings'])}</ul>

{f"<h2>硬性问题</h2><ul>{li(data['hardFailures'])}</ul>" if data['hardFailures'] else ''}

<h2>请人工检查以下事项</h2>
<ol>{li(data['manualChecks'])}</ol>

<p style="color:#888;font-size:12px;margin-top:32px">
本报告由论文排版 Skill 自动生成。原始输入文件未被修改;详细技术报告见 output/debug/ 目录。
</p>
</body>
</html>
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a Chinese HTML review report from formatter outputs.")
    parser.add_argument("--output-dir", required=True, help="Formatter output directory containing the JSON reports")
    parser.add_argument("--report-html", required=True, help="Path to write review_report.html")
    args = parser.parse_args()

    output_dir = Path(args.output_dir).resolve()
    report_path = Path(args.report_html).resolve()
    data = collect(output_dir)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(render_html(data, report_path.parent), encoding="utf-8")
    print(json.dumps({"reportHtml": str(report_path)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
