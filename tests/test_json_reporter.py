import json
from pathlib import Path

from doc_fix.model import AiReviewFinding, CheckIssue, CheckReport
from doc_fix.reporter import report_to_dict, write_json_report


def test_report_to_dict_marks_error_report_as_failed() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
        annotated_docx_path="out/input.annotated.docx",
        annotation_warnings=("word_count.anchor_missing: 未能定位",),
        issues=(CheckIssue(code="word_count.exceeded", severity="error", message="超限"),),
    )

    data = report_to_dict(report)

    assert data["passed"] is False
    assert data["annotated_docx_path"] == "out/input.annotated.docx"
    assert data["annotation_warnings"] == ["word_count.anchor_missing: 未能定位"]
    assert data["issues"][0]["code"] == "word_count.exceeded"


def test_report_to_dict_includes_ai_review_without_affecting_passed() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
        ai_review_findings=(
            AiReviewFinding(
                code="boundary_uncertain",
                message="该段可能是模板说明文字。",
                confidence=0.8,
            ),
        ),
    )

    data = report_to_dict(report)

    assert data["passed"] is True
    assert data["ai_review_findings"][0]["code"] == "boundary_uncertain"


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
