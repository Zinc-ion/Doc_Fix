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
    end_paragraph_index: int | None = None
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
    annotated_docx_path: str | None = None
    annotation_warnings: tuple[str, ...] = field(default_factory=tuple)
    ai_summary: str | None = None
    ai_suggestions: tuple[str, ...] = field(default_factory=tuple)
    ai_review_findings: tuple["AiReviewFinding", ...] = field(default_factory=tuple)
    ai_error: str | None = None

    @property
    def passed(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)


@dataclass(frozen=True)
class AiReviewFinding:
    """AI-generated review item that does not affect rigid pass/fail."""

    code: str
    message: str
    confidence: float | None = None
    chapter_path: str | None = None
    locator: str | None = None
    evidence: str | None = None
    suggested_action: str | None = None


@dataclass(frozen=True)
class CorrectionAction:
    """One correction applied to a generated copy."""

    code: str
    message: str
    count: int = 0
    paragraph_index: int | None = None
    table_index: int | None = None


@dataclass(frozen=True)
class CorrectionReport:
    """Full automatic correction report."""

    template_path: str
    input_path: str
    template_docx_path: str
    input_docx_path: str
    corrected_docx_path: str
    corrected_doc_path: str | None = None
    corrected_doc_export_error: str | None = None
    corrected_check_report_path: str | None = None
    corrected_doc_check_report_path: str | None = None
    actions: tuple[CorrectionAction, ...] = field(default_factory=tuple)
    warnings: tuple[str, ...] = field(default_factory=tuple)
