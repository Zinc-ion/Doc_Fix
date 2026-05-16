"""JSON report serialization."""

from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from doc_fix.model import CheckReport


def report_to_dict(report: CheckReport) -> dict[str, Any]:
    return {
        "passed": report.passed,
        "template_path": report.template_path,
        "input_path": report.input_path,
        "template_docx_path": report.template_docx_path,
        "input_docx_path": report.input_docx_path,
        "issues": [asdict(issue) for issue in report.issues],
        "ai_summary": report.ai_summary,
        "ai_suggestions": list(report.ai_suggestions),
        "ai_error": report.ai_error,
    }


def write_json_report(report: CheckReport, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(report_to_dict(report), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
