"""Rich terminal report rendering."""

from __future__ import annotations

from rich.console import Console
from rich.table import Table

from doc_fix.model import CheckReport


def render_report(report: CheckReport, console: Console | None = None) -> None:
    console = console or Console()
    status = "[green]PASS[/green]" if report.passed else "[red]FAIL[/red]"
    console.print(f"Doc_Fix check result: {status}")
    console.print(f"Template copy: {report.template_docx_path}")
    console.print(f"Input copy: {report.input_docx_path}")

    if not report.issues:
        console.print("[green]未发现刚性规则问题。[/green]")
        return

    table = Table(show_header=True, header_style="bold")
    table.add_column("Severity")
    table.add_column("Code")
    table.add_column("Section")
    table.add_column("Expected")
    table.add_column("Actual")
    table.add_column("Message")

    for issue in report.issues:
        table.add_row(
            issue.severity,
            issue.code,
            issue.section_title or "",
            str(issue.expected) if issue.expected is not None else "",
            str(issue.actual) if issue.actual is not None else "",
            issue.message,
        )
    console.print(table)
