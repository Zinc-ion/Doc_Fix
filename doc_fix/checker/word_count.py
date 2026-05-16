"""Word-count helpers."""

from __future__ import annotations

import re


TOKEN_PATTERN = re.compile(r"[\u4e00-\u9fff]|[A-Za-z0-9]+")
FULL_WIDTH_HINT_PATTERN = re.compile(r"【[^】]*】")


def count_mixed_words(text: str) -> int:
    """Count Chinese chars plus English/number tokens."""

    return len(TOKEN_PATTERN.findall(text))


def strip_rule_hint(text: str) -> str:
    text = re.sub(r"(?:限|不超过)?\s*[0-9０-９]{2,6}\s*字\s*以内(?:（?不包括表格）?)?", "", text)
    return FULL_WIDTH_HINT_PATTERN.sub("", text)
