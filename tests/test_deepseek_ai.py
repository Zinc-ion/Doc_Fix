from doc_fix.ai.deepseek import build_prompt, parse_assistance
from doc_fix.model import CheckIssue, CheckReport


def test_parse_assistance_reads_json_object() -> None:
    assistance = parse_assistance('{"summary":"需要核对","suggestions":["先看字数","再看表格"]}')

    assert assistance.summary == "需要核对"
    assert assistance.suggestions == ("先看字数", "再看表格")


def test_parse_assistance_reads_ai_review_findings() -> None:
    assistance = parse_assistance(
        """
        {
          "summary": "需要核对",
          "suggestions": [],
          "review_findings": [
            {
              "code": "boundary_uncertain",
              "message": "该段可能是模板说明文字。",
              "confidence": 0.72,
              "chapter_path": "第二部分 > 一、项目目标",
              "locator": "目标文档段落 59",
              "evidence": "包含【研究内容】提示语",
              "suggested_action": "人工确认是否计入字数"
            }
          ]
        }
        """
    )

    assert len(assistance.review_findings) == 1
    finding = assistance.review_findings[0]
    assert finding.code == "boundary_uncertain"
    assert finding.confidence == 0.72
    assert finding.chapter_path == "第二部分 > 一、项目目标"


def test_build_prompt_uses_structured_report_without_secret() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
        issues=(CheckIssue(code="table.count_mismatch", severity="error", message="表格数量不一致"),),
    )

    prompt = build_prompt(report)

    assert "table.count_mismatch" in prompt
    assert "DEEPSEEK_API_KEY" not in prompt
    assert "sk-" not in prompt


def test_build_prompt_can_request_ai_review_findings() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="out/template.converted.docx",
        input_docx_path="out/input.converted.docx",
    )

    prompt = build_prompt(report, include_review=True)

    assert "review_findings" in prompt
    assert "人工确认项" in prompt
