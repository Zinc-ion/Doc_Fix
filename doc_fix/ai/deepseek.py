"""DeepSeek-powered optional report assistance."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from doc_fix.model import CheckReport
from doc_fix.reporter import report_to_dict


class AiConfigError(RuntimeError):
    """Raised when AI assistance cannot be configured."""


@dataclass(frozen=True)
class AiAssistance:
    """AI-generated report enrichment."""

    summary: str
    suggestions: tuple[str, ...]


@dataclass(frozen=True)
class DeepSeekSettings:
    api_key: str
    base_url: str
    model: str
    timeout_seconds: float


class DeepSeekAssistant:
    """Optional DeepSeek assistant for human-readable report guidance."""

    def __init__(self, settings: DeepSeekSettings | None = None) -> None:
        self._settings = settings or load_settings()

    def analyze(self, report: CheckReport) -> AiAssistance:
        try:
            from openai import OpenAI  # type: ignore[import-not-found]
        except ImportError as exc:
            raise AiConfigError("openai package is required for --ai. Please install requirements.txt.") from exc

        client = OpenAI(
            api_key=self._settings.api_key,
            base_url=self._settings.base_url,
            timeout=self._settings.timeout_seconds,
        )
        response = client.chat.completions.create(
            model=self._settings.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "你是申报书格式检查助手。只根据结构化检查结果生成摘要和人工核对建议。"
                        "不要改变检查通过/失败结论。必须输出 JSON。"
                    ),
                },
                {
                    "role": "user",
                    "content": build_prompt(report),
                },
            ],
            temperature=0.2,
            max_tokens=900,
        )
        content = response.choices[0].message.content or "{}"
        return parse_assistance(content)


def load_settings(env_path: Path | None = None) -> DeepSeekSettings:
    env_path = env_path or Path(".env")
    try:
        from dotenv import load_dotenv  # type: ignore[import-not-found]
    except ImportError as exc:
        raise AiConfigError("python-dotenv package is required for --ai. Please install requirements.txt.") from exc

    if env_path.exists():
        load_dotenv(env_path, override=True, encoding="utf-8-sig")
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise AiConfigError("DEEPSEEK_API_KEY is not configured in .env or environment variables.")

    timeout_raw = os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "30").strip()
    try:
        timeout_seconds = float(timeout_raw)
    except ValueError as exc:
        raise AiConfigError("DEEPSEEK_TIMEOUT_SECONDS must be a number.") from exc

    return DeepSeekSettings(
        api_key=api_key,
        base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip(),
        model=os.getenv("DEEPSEEK_MODEL", "deepseek-v4-pro").strip(),
        timeout_seconds=timeout_seconds,
    )


def build_prompt(report: CheckReport) -> str:
    data = report_to_dict(report)
    compact = {
        "passed": data["passed"],
        "template_path": data["template_path"],
        "input_path": data["input_path"],
        "issue_count": len(data["issues"]),
        "issues": data["issues"][:30],
    }
    return (
        "请基于下面的 Doc_Fix 结构化检查结果，输出 JSON："
        "{\"summary\":\"...\",\"suggestions\":[\"...\"]}。"
        "summary 用一段话概括，suggestions 给 3-8 条人工核对/修改建议。"
        "每条建议必须尽量引用 issue 中的 chapter_path、caption 或 locator，"
        "不要只说“第几个表格”。不要输出 Markdown，不要改变 passed 结论。\n"
        + json.dumps(compact, ensure_ascii=False)
    )


def parse_assistance(content: str) -> AiAssistance:
    data = _loads_json_object(content)
    summary = str(data.get("summary") or "").strip()
    suggestions_raw = data.get("suggestions") or []
    if not isinstance(suggestions_raw, list):
        suggestions_raw = []
    suggestions = tuple(str(item).strip() for item in suggestions_raw if str(item).strip())
    return AiAssistance(summary=summary, suggestions=suggestions)


def _loads_json_object(content: str) -> dict[str, Any]:
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        start = content.find("{")
        end = content.rfind("}")
        if start < 0 or end <= start:
            raise AiConfigError("DeepSeek returned non-JSON content.")
        data = json.loads(content[start : end + 1])
    if not isinstance(data, dict):
        raise AiConfigError("DeepSeek returned JSON that is not an object.")
    return data
