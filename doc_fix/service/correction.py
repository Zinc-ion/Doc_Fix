"""Reusable automatic correction workflow."""

from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path
from typing import Protocol

from doc_fix.checker import DocumentChecker
from doc_fix.config import apply_rule_config, load_config
from doc_fix.corrector import DocxCorrector, write_correction_report
from doc_fix.converter import ConversionResult
from doc_fix.extractor import DocxExtractor, WordCountRuleExtractor
from doc_fix.model import CheckReport, CorrectionReport
from doc_fix.reporter import write_json_report


class DocumentConverter(Protocol):
    """Converter interface needed by the correction workflow."""

    def normalize_to_docx(self, input_path: Path, out_dir: Path, label: str) -> ConversionResult:
        """Normalize a .doc/.docx input to a .docx work copy."""

    def export_to_doc(self, docx_path: Path, output_path: Path) -> Path:
        """Export a .docx file to .doc."""


@dataclass(frozen=True)
class CorrectionArtifacts:
    """Paths produced by an automatic correction run."""

    correction_report: CorrectionReport
    corrected_docx_path: Path
    corrected_doc_path: Path | None
    correction_report_path: Path
    corrected_check_report_path: Path
    corrected_doc_check_report_path: Path | None

    @property
    def preferred_download_path(self) -> Path:
        """Return the .doc export when available, otherwise the corrected .docx."""

        return self.corrected_doc_path or self.corrected_docx_path


def run_correction(
    template_path: Path,
    input_path: Path,
    out_dir: Path,
    converter: DocumentConverter,
    config_path: Path | None = None,
) -> CorrectionArtifacts:
    """Create an automatically corrected copy and verification reports."""

    extractor = DocxExtractor()
    rule_extractor = WordCountRuleExtractor()
    checker = DocumentChecker()
    corrector = DocxCorrector()

    template_copy = converter.normalize_to_docx(template_path, out_dir, "template")
    input_copy = converter.normalize_to_docx(input_path, out_dir, "input")

    template_snapshot = extractor.extract(template_copy.docx_path)
    input_snapshot = extractor.extract(input_copy.docx_path)
    rules = rule_extractor.extract(template_snapshot)
    format_rules = None
    if config_path is not None:
        config = load_config(config_path)
        rules = apply_rule_config(rules, config.word_count_rules)
        format_rules = config.format_rules or None

    corrected_docx_path = out_dir / "input.corrected.docx"
    correction_report = corrector.correct(
        template_copy.docx_path,
        input_copy.docx_path,
        corrected_docx_path,
        template_snapshot,
        input_snapshot,
        format_rules=format_rules,
    )

    corrected_snapshot = extractor.extract(corrected_docx_path)
    corrected_check_report = CheckReport(
        template_path=str(template_path),
        input_path=str(input_path),
        template_docx_path=str(template_copy.docx_path),
        input_docx_path=str(corrected_docx_path),
        issues=checker.check(template_snapshot, corrected_snapshot, rules, format_rules=format_rules),
    )
    corrected_check_report_path = out_dir / "corrected_check_report.json"
    write_json_report(corrected_check_report, corrected_check_report_path)

    corrected_doc_path = out_dir / "input.corrected.doc"
    corrected_doc_check_report_path = out_dir / "corrected_doc_check_report.json"
    corrected_doc_export_error = None
    corrected_doc_path_for_report = None
    corrected_doc_check_report_path_for_report = None
    try:
        converter.export_to_doc(corrected_docx_path, corrected_doc_path)
        corrected_doc_path_for_report = str(corrected_doc_path)
        corrected_doc_copy = converter.normalize_to_docx(corrected_doc_path, out_dir, "corrected_doc")
        corrected_doc_snapshot = extractor.extract(corrected_doc_copy.docx_path)
        corrected_doc_check_report = CheckReport(
            template_path=str(template_path),
            input_path=str(corrected_doc_path),
            template_docx_path=str(template_copy.docx_path),
            input_docx_path=str(corrected_doc_copy.docx_path),
            issues=checker.check(template_snapshot, corrected_doc_snapshot, rules, format_rules=format_rules),
        )
        write_json_report(corrected_doc_check_report, corrected_doc_check_report_path)
        corrected_doc_check_report_path_for_report = str(corrected_doc_check_report_path)
    except Exception as exc:
        corrected_doc_export_error = str(exc)

    correction_report = replace(
        correction_report,
        template_path=str(template_path),
        input_path=str(input_path),
        corrected_doc_path=corrected_doc_path_for_report,
        corrected_doc_export_error=corrected_doc_export_error,
        corrected_check_report_path=str(corrected_check_report_path),
        corrected_doc_check_report_path=corrected_doc_check_report_path_for_report,
    )
    correction_report_path = out_dir / "correction_report.json"
    write_correction_report(correction_report, correction_report_path)

    return CorrectionArtifacts(
        correction_report=correction_report,
        corrected_docx_path=corrected_docx_path,
        corrected_doc_path=Path(corrected_doc_path_for_report) if corrected_doc_path_for_report else None,
        correction_report_path=correction_report_path,
        corrected_check_report_path=corrected_check_report_path,
        corrected_doc_check_report_path=(
            Path(corrected_doc_check_report_path_for_report) if corrected_doc_check_report_path_for_report else None
        ),
    )
