"""Static HTML report rendering."""

from __future__ import annotations

from html import escape
from pathlib import Path

from doc_fix.model import CheckIssue, CheckReport


def render_html_report(report: CheckReport) -> str:
    status_class = "pass" if report.passed else "fail"
    status = "PASS" if report.passed else "FAIL"
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>Doc_Fix 检查报告</title>
  <style>
    body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 32px; color: #202124; }}
    h1, h2 {{ margin-bottom: 12px; }}
    .status {{ display: inline-block; padding: 4px 10px; border-radius: 4px; color: #fff; }}
    .pass {{ background: #188038; }}
    .fail {{ background: #d93025; }}
    code {{ background: #f1f3f4; padding: 2px 4px; border-radius: 3px; }}
    .issue {{ border-left: 4px solid #dadce0; background: #fff; padding: 10px 12px; margin: 10px 0; }}
    .issue.error {{ border-left-color: #d93025; }}
    .issue.warning {{ border-left-color: #b06000; }}
    .error {{ color: #d93025; font-weight: 600; }}
    .warning {{ color: #b06000; font-weight: 600; }}
    .panel {{ border: 1px solid #dadce0; border-radius: 6px; padding: 12px 16px; margin: 16px 0; }}
  </style>
</head>
<body>
  <h1>Doc_Fix 检查报告</h1>
  <p>检查结论：<span class="status {status_class}">{status}</span></p>
  {_render_paths(report)}
  {_render_annotation_warnings(report)}
  {_render_ai(report)}
  {_render_summary(report)}
  {_render_issues(report)}
</body>
</html>
"""


def write_html_report(report: CheckReport, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(render_html_report(report), encoding="utf-8")


def _render_paths(report: CheckReport) -> str:
    items = [
        ("模板文件", report.template_path),
        ("目标文件", report.input_path),
        ("模板工作副本", report.template_docx_path),
        ("目标工作副本", report.input_docx_path),
        ("标注审核副本", report.annotated_docx_path or "未生成"),
    ]
    return "<div class=\"panel\">" + "".join(
        f"<p><strong>{escape(label)}：</strong><code>{escape(value)}</code></p>"
        for label, value in items
    ) + "</div>"


def _render_annotation_warnings(report: CheckReport) -> str:
    if not report.annotation_warnings:
        return ""
    warnings = "".join(f"<li>{escape(item)}</li>" for item in report.annotation_warnings)
    return f"<h2>标注状态</h2><div class=\"panel\"><ul>{warnings}</ul></div>"


def _render_ai(report: CheckReport) -> str:
    parts: list[str] = []
    if report.ai_summary:
        parts.append(f"<h2>AI 辅助摘要</h2><div class=\"panel\"><p>{escape(report.ai_summary)}</p></div>")
    if report.ai_suggestions:
        suggestions = "".join(f"<li>{escape(item)}</li>" for item in report.ai_suggestions)
        parts.append(f"<h2>AI 处理建议</h2><div class=\"panel\"><ul>{suggestions}</ul></div>")
    if report.ai_review_findings:
        findings = "".join(_render_ai_review_finding(item) for item in report.ai_review_findings)
        parts.append(f"<h2>AI 收尾复核</h2><div class=\"panel\">{findings}</div>")
    if report.ai_error:
        parts.append(f"<h2>AI 辅助状态</h2><div class=\"panel\"><p>{escape(report.ai_error)}</p></div>")
    return "".join(parts)


def _render_ai_review_finding(finding) -> str:
    confidence = "" if finding.confidence is None else f"（置信度：{finding.confidence:.2f}）"
    rows = [
        f"<p><strong>{escape(finding.code)}</strong>{escape(confidence)}：{escape(finding.message)}</p>",
    ]
    if finding.locator:
        rows.append(f"<p><strong>怎么找：</strong>{escape(finding.locator)}</p>")
    if finding.chapter_path:
        rows.append(f"<p><strong>章节：</strong>{escape(finding.chapter_path)}</p>")
    if finding.evidence:
        rows.append(f"<p><strong>依据：</strong>{escape(finding.evidence)}</p>")
    if finding.suggested_action:
        rows.append(f"<p><strong>建议：</strong>{escape(finding.suggested_action)}</p>")
    return "<div class=\"issue warning\">" + "".join(rows) + "</div>"


def _render_summary(report: CheckReport) -> str:
    errors = sum(1 for issue in report.issues if issue.severity == "error")
    warnings = sum(1 for issue in report.issues if issue.severity == "warning")
    return f"<h2>问题汇总</h2><ul><li>错误：{errors}</li><li>警告：{warnings}</li></ul>"


def _render_issues(report: CheckReport) -> str:
    if not report.issues:
        return "<h2>检查明细</h2><p>未发现刚性规则问题。</p>"
    sections: list[str] = ["<h2>检查明细</h2>"]
    for chapter, issues in group_issues_by_chapter(report.issues).items():
        items = "".join(_render_issue(issue) for issue in issues)
        sections.append(f"<h3>{escape(chapter)}</h3>{items}")
    return "".join(sections)


def _render_issue(issue: CheckIssue) -> str:
    expected = "" if issue.expected is None else str(issue.expected)
    actual = "" if issue.actual is None else str(issue.actual)
    preview = f"<p><strong>内容摘录：</strong>{escape(issue.content_preview)}</p>" if issue.content_preview else ""
    caption = issue.caption or issue.nearby_heading or "未提供"
    return (
        f"<div class=\"issue {escape(issue.severity)}\">"
        f"<p><span class=\"{escape(issue.severity)}\">{escape(issue.severity.upper())}</span> "
        f"<code>{escape(issue.code)}</code> {escape(issue.message)}</p>"
        f"<p><strong>怎么找：</strong>{escape(issue.locator or '查看本章节附近内容')}</p>"
        f"<p><strong>板块/标题：</strong>{escape(issue.section_title or issue.nearby_heading or '未提供')}</p>"
        f"<p><strong>附近标题/表题：</strong>{escape(caption)}</p>"
        f"<p><strong>期望/实际：</strong><code>{escape(expected)}</code> / <code>{escape(actual)}</code></p>"
        f"{preview}"
        "</div>"
    )


def group_issues_by_chapter(issues: tuple[CheckIssue, ...]) -> dict[str, list[CheckIssue]]:
    grouped: dict[str, list[CheckIssue]] = {}
    for issue in issues:
        chapter = issue.chapter_path or "未能定位，需人工确认"
        grouped.setdefault(chapter, []).append(issue)
    return grouped
