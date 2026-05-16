from doc_fix.extractor.docx import infer_chinese_heading_level
from doc_fix.model import CheckIssue, CheckReport
from doc_fix.reporter import render_html_report, render_markdown_report, report_to_dict


def test_infer_chinese_heading_levels() -> None:
    assert infer_chinese_heading_level("第一部分 国内外现状及趋势分析") == 1
    assert infer_chinese_heading_level("一、项目目标及考核指标") == 2
    assert infer_chinese_heading_level("（一）申报项目与所属指南方向的关联关系") == 3
    assert infer_chinese_heading_level("1、研究内容") == 4


def test_captions_are_not_headings() -> None:
    assert infer_chinese_heading_level("表1 国外从事相关研究的主要机构") is None
    assert infer_chinese_heading_level("图 1 技术路线图") is None


def test_report_json_contains_location_fields() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="template.docx",
        input_docx_path="target.docx",
        issues=(
            CheckIssue(
                code="word_count.exceeded",
                severity="error",
                message="超限",
                chapter_path="第二部分 > 一、项目目标",
                paragraph_index=12,
                content_preview="这是摘录",
                locator="目标文档段落 13",
            ),
        ),
    )

    issue = report_to_dict(report)["issues"][0]

    assert issue["chapter_path"] == "第二部分 > 一、项目目标"
    assert issue["paragraph_index"] == 12
    assert issue["content_preview"] == "这是摘录"


def test_human_reports_group_by_chapter_and_show_locator() -> None:
    report = CheckReport(
        template_path="template.doc",
        input_path="target.doc",
        template_docx_path="template.docx",
        input_docx_path="target.docx",
        issues=(
            CheckIssue(
                code="table.columns_mismatch",
                severity="error",
                message="列数不一致",
                chapter_path="第二部分 > 一、项目目标",
                caption="表4 项目目标、成果与考核指标表",
                locator="本章节第 1 个表格 / 全文第 4 个表格",
            ),
        ),
    )

    markdown = render_markdown_report(report)
    html = render_html_report(report)

    assert "### 第二部分 > 一、项目目标" in markdown
    assert "本章节第 1 个表格 / 全文第 4 个表格" in markdown
    assert "表4 项目目标、成果与考核指标表" in markdown
    assert "第二部分 &gt; 一、项目目标" in html
    assert "本章节第 1 个表格 / 全文第 4 个表格" in html
