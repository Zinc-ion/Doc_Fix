from pathlib import Path

from doc_fix.model import AiReviewFinding, CheckIssue, CheckReport
from doc_fix.reporter import render_html_report, render_markdown_report, write_html_report, write_markdown_report


def make_report() -> CheckReport:
    return CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
        issues=(
            CheckIssue(
                code="word_count.exceeded",
                severity="error",
                message="项目简介 <超限>",
                expected=1500,
                actual=1600,
                section_title="项目简介",
            ),
        ),
        ai_summary="存在字数超限。",
        ai_suggestions=("压缩项目简介。",),
        ai_review_findings=(
            AiReviewFinding(
                code="boundary_uncertain",
                message="该段可能是模板说明文字。",
                confidence=0.7,
                locator="目标文档段落 59",
                suggested_action="人工确认是否计入字数",
            ),
        ),
    )


def test_render_markdown_report_contains_human_sections() -> None:
    markdown = render_markdown_report(make_report())

    assert "# Doc_Fix 检查报告" in markdown
    assert "AI 辅助摘要" in markdown
    assert "AI 收尾复核" in markdown
    assert "boundary_uncertain" in markdown
    assert "word_count.exceeded" in markdown
    assert "项目简介" in markdown


def test_write_markdown_report(tmp_path: Path) -> None:
    output_path = tmp_path / "report.md"

    write_markdown_report(make_report(), output_path)

    assert output_path.read_text(encoding="utf-8").startswith("# Doc_Fix 检查报告")


def test_render_html_report_escapes_issue_content() -> None:
    html = render_html_report(make_report())

    assert "&lt;超限&gt;" in html
    assert "<超限>" not in html


def test_write_html_report(tmp_path: Path) -> None:
    output_path = tmp_path / "report.html"

    write_html_report(make_report(), output_path)

    assert "<!doctype html>" in output_path.read_text(encoding="utf-8")
