"""Report rendering APIs."""

from typing import Any

from doc_fix.model import CheckReport
from doc_fix.reporter.docx_annotator import write_annotated_docx
from doc_fix.reporter.json_reporter import report_to_dict, write_json_report
from doc_fix.reporter.markdown_reporter import render_markdown_report, write_markdown_report
from doc_fix.reporter.html_reporter import render_html_report, write_html_report


def render_report(report: CheckReport, console: Any = None) -> None:
    from doc_fix.reporter.rich_reporter import render_report as _render_report

    _render_report(report, console)

__all__ = [
    "render_html_report",
    "render_markdown_report",
    "render_report",
    "report_to_dict",
    "write_html_report",
    "write_annotated_docx",
    "write_json_report",
    "write_markdown_report",
]
