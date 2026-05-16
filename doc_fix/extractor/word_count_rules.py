"""Extract word-count rules from template text."""

from __future__ import annotations

import re

from doc_fix.model import CountPolicy, DocumentSnapshot, WordCountRule


LIMIT_PATTERN = re.compile(
    r"(?P<prefix>每个[^，。；;\n\r]{0,12}?)?"
    r"(?:限|不超过)?\s*"
    r"(?P<limit>[0-9０-９]{2,6})\s*字\s*以内"
    r"(?P<suffix>（?不包括表格）?)?"
)


class WordCountRuleExtractor:
    """Infer section word-count rules from template paragraphs."""

    def extract(self, snapshot: DocumentSnapshot) -> tuple[WordCountRule, ...]:
        rules: list[WordCountRule] = []
        current_heading: tuple[str, int] | None = None
        for paragraph in snapshot.paragraphs:
            text = normalize_text(paragraph.text)
            if not text:
                continue
            if paragraph.heading_level is not None and text:
                current_heading = (text, paragraph.heading_level)
            for match in LIMIT_PATTERN.finditer(text):
                limit = int(to_half_width_digits(match.group("limit")))
                title = infer_section_title(text, match.start(), current_heading)
                exclude_tables = "不包括表格" in (match.group("suffix") or "")
                rules.append(
                    WordCountRule(
                        section_title=title,
                        limit=limit,
                        anchor_text=text,
                        source="template",
                        policy=CountPolicy(exclude_tables=exclude_tables),
                        paragraph_index=paragraph.index,
                        start_offset=match.end(),
                        heading_level=paragraph.heading_level or (current_heading[1] if current_heading else None),
                        chapter_path=paragraph.chapter_path,
                    )
                )
        return tuple(dedupe_rules(rules))


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u0007", " ")).strip()


def to_half_width_digits(value: str) -> str:
    table = str.maketrans("０１２３４５６７８９", "0123456789")
    return value.translate(table)


def infer_section_title(text: str, limit_start: int, current_heading: tuple[str, int] | None = None) -> str:
    before = text[:limit_start].strip(" ，。；;:：")
    if not before:
        return current_heading[0] if current_heading else "未命名板块"
    sentence_parts = re.split(r"[。；;]", before)
    candidate = sentence_parts[-1].strip(" ，,:：")
    candidate = trim_leading_description(candidate)
    if len(candidate) > 48:
        candidate = candidate[-48:].strip(" ，,:：")
    if not candidate and current_heading:
        return current_heading[0]
    return candidate or "未命名板块"


def trim_leading_description(candidate: str) -> str:
    markers = ("包括", "围绕", "从", "针对", "拟", "项目")
    for marker in markers:
        index = candidate.rfind(marker)
        if index > 0:
            prefix = candidate[:index].strip()
            if len(prefix) >= 3:
                return prefix
    return candidate


def dedupe_rules(rules: list[WordCountRule]) -> list[WordCountRule]:
    seen: set[tuple[str, int, int | None]] = set()
    result: list[WordCountRule] = []
    for rule in rules:
        key = (rule.section_title, rule.limit, rule.paragraph_index)
        if key in seen:
            continue
        seen.add(key)
        result.append(rule)
    return result
