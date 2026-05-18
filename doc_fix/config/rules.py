"""Load and apply user rule configuration."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from doc_fix.model import CountPolicy, FormatRule, WordCountRule


@dataclass(frozen=True)
class RuleConfig:
    """Full optional rule configuration loaded from JSON."""

    word_count_rules: tuple[WordCountRule, ...] = ()
    format_rules: tuple[FormatRule, ...] = ()


def load_rule_config(path: Path) -> tuple[WordCountRule, ...]:
    """Load word-count rules from a JSON config file.

    Expected shape:
    {
      "word_count_rules": [
        {
          "section_title": "项目简介",
          "limit": 1500,
          "exclude_tables": true,
          "enabled": true
        }
      ]
    }
    """

    return load_config(path).word_count_rules


def load_config(path: Path) -> RuleConfig:
    """Load all supported optional rules from a JSON config file."""

    data = json.loads(path.read_text(encoding="utf-8"))
    return RuleConfig(
        word_count_rules=_parse_word_count_rules(data.get("word_count_rules", [])),
        format_rules=_parse_format_rules(data.get("format_rules", [])),
    )


def _parse_word_count_rules(rules: Any) -> tuple[WordCountRule, ...]:
    if not isinstance(rules, list):
        raise ValueError("Config field 'word_count_rules' must be a list.")

    parsed: list[WordCountRule] = []
    for index, raw_rule in enumerate(rules):
        if not isinstance(raw_rule, dict):
            raise ValueError(f"word_count_rules[{index}] must be an object.")
        section_title = _required_str(raw_rule, "section_title", index)
        limit = _required_int(raw_rule, "limit", index)
        enabled = bool(raw_rule.get("enabled", True))
        source = "config" if enabled else "config:disabled"
        parsed.append(
            WordCountRule(
                section_title=section_title,
                limit=limit,
                anchor_text=str(raw_rule.get("anchor_text") or section_title),
                source=source,
                policy=CountPolicy(exclude_tables=bool(raw_rule.get("exclude_tables", False))),
                paragraph_index=_optional_int(raw_rule.get("paragraph_index")),
                start_offset=_optional_int(raw_rule.get("start_offset")),
                heading_level=_optional_int(raw_rule.get("heading_level")),
            )
        )
    return tuple(parsed)


def _parse_format_rules(rules: Any) -> tuple[FormatRule, ...]:
    if not isinstance(rules, list):
        raise ValueError("Config field 'format_rules' must be a list.")

    parsed: list[FormatRule] = []
    for index, raw_rule in enumerate(rules):
        if not isinstance(raw_rule, dict):
            raise ValueError(f"format_rules[{index}] must be an object.")
        scope = _required_str(raw_rule, "scope", index, parent="format_rules")
        if scope not in ("body", "heading"):
            raise ValueError(f"format_rules[{index}].scope must be 'body' or 'heading'.")
        parsed.append(
            FormatRule(
                scope=scope,
                enabled=bool(raw_rule.get("enabled", True)),
                check_font_names=bool(raw_rule.get("check_font_names", True)),
                check_font_sizes=bool(raw_rule.get("check_font_sizes", True)),
                check_font_colors=bool(raw_rule.get("check_font_colors", True)),
                check_font_highlights=bool(raw_rule.get("check_font_highlights", True)),
                check_alignment=bool(raw_rule.get("check_alignment", True)),
                check_indent=bool(raw_rule.get("check_indent", True)),
                check_line_spacing=bool(raw_rule.get("check_line_spacing", True)),
                check_spacing=bool(raw_rule.get("check_spacing", True)),
                font_size_tolerance_pt=_optional_float(raw_rule.get("font_size_tolerance_pt"), default=0.1),
                indent_tolerance_twips=_optional_int(raw_rule.get("indent_tolerance_twips"), default=20),
                line_spacing_tolerance=_optional_float(raw_rule.get("line_spacing_tolerance"), default=0.05),
                spacing_tolerance_twips=_optional_int(raw_rule.get("spacing_tolerance_twips"), default=20),
            )
        )
    return tuple(parsed)


def apply_rule_config(
    template_rules: tuple[WordCountRule, ...],
    config_rules: tuple[WordCountRule, ...],
) -> tuple[WordCountRule, ...]:
    """Override/add/remove template rules with config rules."""

    merged: dict[str, WordCountRule] = {
        normalize_title(rule.section_title): rule
        for rule in template_rules
    }
    for rule in config_rules:
        key = normalize_title(rule.section_title)
        if rule.source == "config:disabled":
            merged.pop(key, None)
            continue
        merged[key] = rule
    return tuple(merged.values())


def normalize_title(value: str) -> str:
    return "".join(value.split())


def _required_str(raw_rule: dict[str, Any], field: str, index: int, parent: str = "word_count_rules") -> str:
    value = raw_rule.get(field)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{parent}[{index}].{field} must be a non-empty string.")
    return value.strip()


def _required_int(raw_rule: dict[str, Any], field: str, index: int) -> int:
    value = raw_rule.get(field)
    if not isinstance(value, int) or value <= 0:
        raise ValueError(f"word_count_rules[{index}].{field} must be a positive integer.")
    return value


def _optional_int(value: Any, default: int | None = None) -> int | None:
    if value is None:
        return default
    if not isinstance(value, int):
        raise ValueError("Optional integer fields must be integers when provided.")
    return value


def _optional_float(value: Any, default: float) -> float:
    if value is None:
        return default
    if not isinstance(value, (int, float)) or value < 0:
        raise ValueError("Optional numeric fields must be non-negative numbers when provided.")
    return float(value)
