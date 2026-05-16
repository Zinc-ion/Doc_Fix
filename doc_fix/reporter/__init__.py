"""Report rendering APIs."""

from doc_fix.reporter.json_reporter import report_to_dict, write_json_report
from doc_fix.reporter.markdown_reporter import render_markdown_report, write_markdown_report
from doc_fix.reporter.html_reporter import render_html_report, write_html_report
from doc_fix.reporter.rich_reporter import render_report

__all__ = [
    "render_html_report",
    "render_markdown_report",
    "render_report",
    "report_to_dict",
    "write_html_report",
    "write_json_report",
    "write_markdown_report",
]
