from doc_fix.ai.deepseek import build_prompt, parse_assistance
from doc_fix.model import CheckIssue, CheckReport


def test_parse_assistance_reads_json_object() -> None:
    assistance = parse_assistance('{"summary":"需要核对","suggestions":["先看字数","再看表格"]}')

    assert assistance.summary == "需要核对"
    assert assistance.suggestions == ("先看字数", "再看表格")


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
