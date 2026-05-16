"""Markdown report rendering."""

from __future__ import annotations

from pathlib import Path

from doc_fix.model import CheckIssue, CheckReport


def render_markdown_report(report: CheckReport) -> str:
    status = "PASS" if report.passed else "FAIL"
    lines = [
        "# Doc_Fix 检查报告",
        "",
        f"- 检查结论：**{status}**",
        f"- 模板文件：`{report.template_path}`",
        f"- 目标文件：`{report.input_path}`",
        f"- 模板工作副本：`{report.template_docx_path}`",
        f"- 目标工作副本：`{report.input_docx_path}`",
        "",
    ]
    lines.extend(_render_ai(report))
    lines.extend(_render_issue_summary(report.issues))
    lines.extend(_render_issue_table(report.issues))
    return "\n".join(lines).rstrip() + "\n"


def write_markdown_report(report: CheckReport, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(render_markdown_report(report), encoding="utf-8")


def _render_ai(report: CheckReport) -> list[str]:
    lines: list[str] = []
    if report.ai_summary:
        lines.extend(["## AI 辅助摘要", "", report.ai_summary, ""])
    if report.ai_suggestions:
        lines.extend(["## AI 处理建议", ""])
        for suggestion in report.ai_suggestions:
            lines.append(f"- {suggestion}")
        lines.append("")
    if report.ai_error:
        lines.extend(["## AI 辅助状态", "", f"> {report.ai_error}", ""])
    return lines


def _render_issue_summary(issues: tuple[CheckIssue, ...]) -> list[str]:
    errors = sum(1 for issue in issues if issue.severity == "error")
    warnings = sum(1 for issue in issues if issue.severity == "warning")
    return ["## 问题汇总", "", f"- 错误：{errors}", f"- 警告：{warnings}", ""]


def _render_issue_table(issues: tuple[CheckIssue, ...]) -> list[str]:
    if not issues:
        return ["## 检查明细", "", "未发现刚性规则问题。", ""]

    lines = ["## 检查明细", ""]
    for chapter, chapter_issues in group_issues_by_chapter(issues).items():
        lines.extend([f"### {chapter}", ""])
        for issue in chapter_issues:
            lines.extend(
                [
                    f"- **{issue.severity.upper()}** `{issue.code}`：{issue.message}",
                    f"  - 怎么找：{issue.locator or '查看本章节附近内容'}",
                    f"  - 板块/标题：{issue.section_title or issue.nearby_heading or '未提供'}",
                    f"  - 附近标题/表题：{issue.caption or issue.nearby_heading or '未提供'}",
                    f"  - 期望/实际：`{_md_cell(issue.expected if issue.expected is not None else '')}` / `{_md_cell(issue.actual if issue.actual is not None else '')}`",
                ]
            )
            if issue.content_preview:
                lines.append(f"  - 内容摘录：{_md_cell(issue.content_preview)}")
        lines.append("")
    return lines


def _md_cell(value) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ")


def group_issues_by_chapter(issues: tuple[CheckIssue, ...]) -> dict[str, list[CheckIssue]]:
    grouped: dict[str, list[CheckIssue]] = {}
    for issue in issues:
        chapter = issue.chapter_path or "未能定位，需人工确认"
        grouped.setdefault(chapter, []).append(issue)
    return grouped
