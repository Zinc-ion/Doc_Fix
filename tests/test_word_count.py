from doc_fix.checker import count_mixed_words


def test_count_mixed_words_counts_chinese_chars_and_english_tokens() -> None:
    assert count_mixed_words("多模态 network 2026 test") == 6


def test_count_mixed_words_ignores_punctuation_and_spaces() -> None:
    assert count_mixed_words("研究， 目标；A1。") == 5
