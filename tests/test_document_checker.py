from doc_fix.checker import DocumentChecker
from doc_fix.model import CountPolicy, DocumentSnapshot, ImageSpec, ParagraphBlock, TableSpec, WordCountRule


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
