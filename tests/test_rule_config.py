import json
from pathlib import Path

from doc_fix.config import apply_rule_config, load_config, load_rule_config
from doc_fix.model import CountPolicy, WordCountRule


def test_load_rule_config_reads_json_rules(tmp_path: Path) -> None:
    config_path = tmp_path / "rules.json"
    config_path.write_text(
        json.dumps(
            {
                "word_count_rules": [
                    {
                        "section_title": "项目简介",
                        "limit": 1200,
                        "exclude_tables": True,
                    }
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    rules = load_rule_config(config_path)

    assert rules[0].section_title == "项目简介"
    assert rules[0].limit == 1200
    assert rules[0].source == "config"
    assert rules[0].policy.exclude_tables is True


def test_load_config_reads_format_rules(tmp_path: Path) -> None:
    config_path = tmp_path / "rules.json"
    config_path.write_text(
        json.dumps(
            {
                "format_rules": [
                    {
                        "scope": "body",
                        "enabled": True,
                        "check_font_names": False,
                        "font_size_tolerance_pt": 0.5,
                        "indent_tolerance_twips": 40,
                    }
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    config = load_config(config_path)

    assert config.word_count_rules == ()
    assert config.format_rules[0].scope == "body"
    assert config.format_rules[0].check_font_names is False
    assert config.format_rules[0].font_size_tolerance_pt == 0.5
    assert config.format_rules[0].indent_tolerance_twips == 40


def test_apply_rule_config_overrides_and_disables_template_rules() -> None:
    template_rules = (
        WordCountRule("项目简介", 1500, "项目简介 限1500字以内"),
        WordCountRule("研究基础", 1000, "研究基础 限1000字以内"),
    )
    config_rules = (
        WordCountRule("项目简介", 1200, "项目简介", source="config", policy=CountPolicy(True)),
        WordCountRule("研究基础", 1, "研究基础", source="config:disabled"),
    )

    merged = apply_rule_config(template_rules, config_rules)

    assert len(merged) == 1
    assert merged[0].section_title == "项目简介"
    assert merged[0].limit == 1200
