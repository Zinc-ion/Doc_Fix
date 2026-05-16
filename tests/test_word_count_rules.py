from doc_fix.extractor.word_count_rules import WordCountRuleExtractor
from doc_fix.model import DocumentSnapshot, ParagraphBlock


def make_snapshot(*texts: str) -> DocumentSnapshot:
    return DocumentSnapshot(
        source_path="template.docx",
        paragraphs=tuple(
            ParagraphBlock(index=index, text=text)
            for index, text in enumerate(texts)
        ),
        tables=(),
        images=(),
    )


def test_extracts_basic_word_count_rule() -> None:
    snapshot = make_snapshot("第一部分 国内外现状及趋势分析。限2000字以内（不包括表格），并填写下表。")

    rules = WordCountRuleExtractor().extract(snapshot)

    assert len(rules) == 1
    assert rules[0].limit == 2000
    assert rules[0].policy.exclude_tables is True
    assert rules[0].source == "template"
    assert "国内外现状及趋势分析" in rules[0].section_title


def test_extracts_full_width_digits() -> None:
    snapshot = make_snapshot("申报项目简介 限１５００字以内")

    rules = WordCountRuleExtractor().extract(snapshot)

    assert len(rules) == 1
    assert rules[0].limit == 1500
