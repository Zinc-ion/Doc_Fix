"""Report data models."""

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class CheckIssue:
    """One check result item."""

    code: str
    severity: str
    message: str
    expected: Any = None
    actual: Any = None
    section_title: str | None = None
    source: str | None = None
    chapter_path: str | None = None
    paragraph_index: int | None = None
    table_index: int | None = None
    image_index: int | None = None
    nearby_heading: str | None = None
    caption: str | None = None
    content_preview: str | None = None
    locator: str | None = None


@dataclass(frozen=True)
class CheckReport:
    """Full check report."""

    template_path: str
    input_path: str
    template_docx_path: str
    input_docx_path: str
    issues: tuple[CheckIssue, ...] = field(default_factory=tuple)
    ai_summary: str | None = None
    ai_suggestions: tuple[str, ...] = field(default_factory=tuple)
    ai_error: str | None = None

    @property
    def passed(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)
