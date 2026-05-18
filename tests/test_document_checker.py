from doc_fix.checker import DocumentChecker
from doc_fix.model import CountPolicy, DocumentSnapshot, FormatRule, ImageSpec, ParagraphBlock, TableSpec, WordCountRule


def test_checker_reports_exceeded_word_count() -> None:
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="项目简介 限3字以内"),
            ParagraphBlock(index=1, text="这是超过 limit"),
        ),
        tables=(),
        images=(),
    )
    rule = WordCountRule(
        section_title="项目简介",
        limit=3,
        anchor_text="项目简介 限3字以内",
        policy=CountPolicy(exclude_tables=False),
        paragraph_index=0,
    )

    issues = DocumentChecker().check(target, target, (rule,))

    assert len(issues) == 1
    assert issues[0].code == "word_count.exceeded"
    assert issues[0].expected == 3
    assert issues[0].actual > 3


def test_template_rule_paragraph_index_does_not_override_target_title_search() -> None:
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="无关段落", chapter_path="封面"),
            ParagraphBlock(index=1, text="项目简介 限3字以内", chapter_path="申报项目简介"),
            ParagraphBlock(index=2, text="这是超过 limit", chapter_path="申报项目简介"),
        ),
        tables=(),
        images=(),
    )
    rule = WordCountRule(
        section_title="项目简介",
        limit=3,
        anchor_text="项目简介 限3字以内",
        paragraph_index=0,
        source="template",
    )

    issues = DocumentChecker().check(target, target, (rule,))

    assert issues[0].paragraph_index == 1
    assert issues[0].chapter_path == "申报项目简介"


def test_template_rule_falls_back_to_target_chapter_path() -> None:
    chapter_path = "第二部分 > 一、项目目标 > （一）关联关系"
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="第二部分", heading_level=1, chapter_path="第二部分", is_heading=True),
            ParagraphBlock(index=1, text="一、项目目标", heading_level=2, chapter_path="第二部分 > 一、项目目标", is_heading=True),
            ParagraphBlock(index=2, text="（一）关联关系", heading_level=3, chapter_path=chapter_path, is_heading=True),
            ParagraphBlock(index=3, text="我我我我", chapter_path=chapter_path),
            ParagraphBlock(index=4, text="（二）后续章节", heading_level=3, chapter_path="第二部分 > 一、项目目标 > （二）后续章节", is_heading=True),
        ),
        tables=(),
        images=(),
    )
    rule = WordCountRule(
        section_title="包括项目与所属指南方向的匹配性",
        limit=3,
        anchor_text="包括项目与所属指南方向的匹配性。限3字以内。",
        source="template",
        heading_level=3,
        chapter_path=chapter_path,
    )

    issues = DocumentChecker().check(target, target, (rule,))

    assert len(issues) == 1
    assert issues[0].code == "word_count.exceeded"
    assert issues[0].actual == 4
    assert issues[0].paragraph_index == 2
    assert issues[0].chapter_path == chapter_path
    assert "按章节路径定位" in (issues[0].locator or "")


def test_chapter_path_fallback_does_not_count_heading_text() -> None:
    chapter_path = "第二部分 > 一、项目目标"
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="一、项目目标", heading_level=2, chapter_path=chapter_path, is_heading=True),
            ParagraphBlock(index=1, text="我", chapter_path=chapter_path),
        ),
        tables=(),
        images=(),
    )
    rule = WordCountRule(
        section_title="模板提示",
        limit=1,
        anchor_text="模板提示。限1字以内。",
        source="template",
        heading_level=2,
        chapter_path=chapter_path,
    )

    issues = DocumentChecker().check(target, target, (rule,))

    assert issues == ()


def test_template_rule_reports_anchor_missing_when_chapter_path_is_absent() -> None:
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="一、其他章节", heading_level=2, chapter_path="一、其他章节", is_heading=True),
            ParagraphBlock(index=1, text="我我我我", chapter_path="一、其他章节"),
        ),
        tables=(),
        images=(),
    )
    rule = WordCountRule(
        section_title="模板提示",
        limit=1,
        anchor_text="模板提示。限1字以内。",
        source="template",
        heading_level=2,
        chapter_path="一、项目目标",
    )

    issues = DocumentChecker().check(target, target, (rule,))

    assert len(issues) == 1
    assert issues[0].code == "word_count.anchor_missing"


def test_checker_warns_when_no_word_count_rules() -> None:
    snapshot = DocumentSnapshot(source_path="target.docx", paragraphs=(), tables=(), images=())

    issues = DocumentChecker().check(snapshot, snapshot, ())

    assert issues[0].code == "word_count.no_rules"
    assert issues[0].severity == "warning"


def test_checker_stops_word_count_at_next_same_level_heading() -> None:
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="第一部分 研究内容 限5字以内", heading_level=1),
            ParagraphBlock(index=1, text="研究内容"),
            ParagraphBlock(index=2, text="第二部分 研究基础", heading_level=1),
            ParagraphBlock(index=3, text="这段不应该计入上一节"),
        ),
        tables=(),
        images=(),
    )
    rule = WordCountRule(
        section_title="第一部分 研究内容",
        limit=5,
        anchor_text="第一部分 研究内容 限5字以内",
        paragraph_index=0,
        start_offset=len("第一部分 研究内容 限5字以内"),
        heading_level=1,
    )

    issues = DocumentChecker().check(target, target, (rule,))

    assert not [issue for issue in issues if issue.code == "word_count.exceeded"]


def test_checker_reports_table_detail_mismatches() -> None:
    template = DocumentSnapshot(
        source_path="template.docx",
        paragraphs=(),
        tables=(
            TableSpec(
                index=0,
                rows=2,
                columns=2,
                text="",
                alignment="center",
                border_signature="top:single:4:auto",
                column_widths_twips=(1000, 1000),
                font_names=("宋体",),
                font_sizes_pt=(12.0,),
            ),
        ),
        images=(),
    )
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(),
        tables=(
            TableSpec(
                index=0,
                rows=3,
                columns=2,
                text="",
                alignment="left",
                border_signature="top:single:8:auto",
                column_widths_twips=(1300, 1000),
                font_names=("黑体",),
                font_sizes_pt=(10.5,),
            ),
        ),
        images=(),
    )

    issues = DocumentChecker().check(template, target, ())
    codes = {issue.code for issue in issues}

    assert "table.rows_mismatch" in codes
    assert "table.alignment_mismatch" in codes
    assert "table.border_signature_mismatch" in codes
    assert "table.column_widths_mismatch" in codes
    assert "table.font_names_mismatch" in codes
    assert "table.font_sizes_mismatch" in codes


def test_checker_reports_image_detail_mismatches() -> None:
    template = DocumentSnapshot(
        source_path="template.docx",
        paragraphs=(),
        tables=(),
        images=(ImageSpec(index=0, width_emu=1000, height_emu=1000, paragraph_alignment="center", wrap_type="inline"),),
    )
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(),
        tables=(),
        images=(ImageSpec(index=0, width_emu=1200, height_emu=1300, paragraph_alignment="left", wrap_type="anchor"),),
    )

    issues = DocumentChecker().check(template, target, ())
    codes = {issue.code for issue in issues}

    assert "image.width_mismatch" in codes
    assert "image.height_mismatch" in codes
    assert "image.wrap_mismatch" in codes
    assert "image.alignment_mismatch" in codes


def test_checker_reports_body_paragraph_format_mismatches() -> None:
    template = DocumentSnapshot(
        source_path="template.docx",
        paragraphs=(
            ParagraphBlock(
                index=0,
                text="正文",
                chapter_path="第一章",
                font_names=("宋体",),
                font_sizes_pt=(12.0,),
                alignment="justify",
                first_line_indent_twips=480,
                line_spacing=1.5,
                space_before_twips=0,
                space_after_twips=120,
            ),
        ),
        tables=(),
        images=(),
    )
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(
                index=0,
                text="正文",
                chapter_path="第一章",
                font_names=("黑体",),
                font_sizes_pt=(10.5,),
                alignment="left",
                first_line_indent_twips=0,
                line_spacing=2.0,
                space_before_twips=120,
                space_after_twips=0,
            ),
        ),
        tables=(),
        images=(),
    )

    issues = DocumentChecker().check(template, target, ())
    codes = {issue.code for issue in issues}

    assert "paragraph.font_names_mismatch" in codes
    assert "paragraph.font_sizes_mismatch" in codes
    assert "paragraph.alignment_mismatch" in codes
    assert "paragraph.indent_mismatch" in codes
    assert "paragraph.line_spacing_mismatch" in codes
    assert "paragraph.spacing_mismatch" in codes
    assert all(issue.severity == "warning" for issue in issues if issue.code.startswith("paragraph."))


def test_format_rules_can_disable_checks_and_apply_tolerance() -> None:
    template = DocumentSnapshot(
        source_path="template.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="正文", chapter_path="第一章", font_names=("宋体",), font_sizes_pt=(12.0,)),
        ),
        tables=(),
        images=(),
    )
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(index=0, text="正文", chapter_path="第一章", font_names=("黑体",), font_sizes_pt=(12.4,)),
        ),
        tables=(),
        images=(),
    )
    format_rules = (
        FormatRule(
            scope="body",
            check_font_names=False,
            font_size_tolerance_pt=0.5,
            check_alignment=False,
            check_indent=False,
            check_line_spacing=False,
            check_spacing=False,
        ),
    )

    issues = DocumentChecker().check(template, target, (), format_rules=format_rules)

    assert not [issue for issue in issues if issue.code.startswith("paragraph.")]


def test_heading_format_compares_only_same_heading_level() -> None:
    template = DocumentSnapshot(
        source_path="template.docx",
        paragraphs=(
            ParagraphBlock(
                index=0,
                text="一、标题",
                heading_level=2,
                chapter_path="一、标题",
                is_heading=True,
                font_names=("黑体",),
            ),
            ParagraphBlock(
                index=1,
                text="正文",
                chapter_path="一、标题",
                font_names=("宋体",),
            ),
        ),
        tables=(),
        images=(),
    )
    target = DocumentSnapshot(
        source_path="target.docx",
        paragraphs=(
            ParagraphBlock(
                index=0,
                text="一、标题",
                heading_level=2,
                chapter_path="一、标题",
                is_heading=True,
                font_names=("宋体",),
            ),
            ParagraphBlock(
                index=1,
                text="正文",
                chapter_path="一、标题",
                font_names=("宋体",),
            ),
        ),
        tables=(),
        images=(),
    )

    paragraph_issues = [
        issue
        for issue in DocumentChecker().check(template, target, ())
        if issue.code.startswith("paragraph.")
    ]

    assert [issue.code for issue in paragraph_issues] == ["paragraph.font_names_mismatch"]
    assert paragraph_issues[0].paragraph_index == 0
