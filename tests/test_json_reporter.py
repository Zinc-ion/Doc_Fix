import json
from pathlib import Path

from doc_fix.model import CheckIssue, CheckReport
from doc_fix.reporter import report_to_dict, write_json_report


def test_report_to_dict_marks_error_report_as_failed() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
        issues=(CheckIssue(code="word_count.exceeded", severity="error", message="超限"),),
    )

    data = report_to_dict(report)

    assert data["passed"] is False
    assert data["issues"][0]["code"] == "word_count.exceeded"


def test_write_json_report(tmp_path: Path) -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
    )
    output_path = tmp_path / "report.json"

    write_json_report(report, output_path)

    assert json.loads(output_path.read_text(encoding="utf-8"))["passed"] is True
