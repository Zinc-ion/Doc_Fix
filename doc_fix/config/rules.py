"""Load and apply user rule configuration."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from doc_fix.model import CountPolicy, WordCountRule


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

    data = json.loads(path.read_text(encoding="utf-8"))
    rules = data.get("word_count_rules", [])
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


def _required_str(raw_rule: dict[str, Any], field: str, index: int) -> str:
    value = raw_rule.get(field)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"word_count_rules[{index}].{field} must be a non-empty string.")
    return value.strip()


def _required_int(raw_rule: dict[str, Any], field: str, index: int) -> int:
    value = raw_rule.get(field)
    if not isinstance(value, int) or value <= 0:
        raise ValueError(f"word_count_rules[{index}].{field} must be a positive integer.")
    return value


def _optional_int(value: Any) -> int | None:
    if value is None:
        return None
    if not isinstance(value, int):
        raise ValueError("Optional integer fields must be integers when provided.")
    return value
