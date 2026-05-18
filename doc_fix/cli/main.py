"""Command-line entrypoint."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import click

from doc_fix.ai import AiConfigError, DeepSeekAssistant
from doc_fix.checker import DocumentChecker
from doc_fix.config import apply_rule_config, load_config
from doc_fix.corrector import DocxCorrector, write_correction_report
from doc_fix.converter import ConversionError, WordConverter
from doc_fix.extractor import DocxExtractor, WordCountRuleExtractor
from doc_fix.model import CheckReport
from doc_fix.reporter import render_report, write_annotated_docx, write_html_report, write_json_report, write_markdown_report


@click.group()
def main() -> None:
    """Check WPS/Word declaration documents."""


@main.command()
@click.option("--template", "template_path", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--input", "input_path", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--out-dir", type=click.Path(file_okay=False, path_type=Path), default=Path("output"), show_default=True)
@click.option("--json", "json_path", type=click.Path(dir_okay=False, path_type=Path), help="Override JSON report path. Defaults to OUT_DIR/report.json.")
@click.option("--markdown", "markdown_path", type=click.Path(dir_okay=False, path_type=Path), help="Override Markdown report path. Defaults to OUT_DIR/report.md.")
@click.option("--html", "html_path", type=click.Path(dir_okay=False, path_type=Path), help="Override HTML report path. Defaults to OUT_DIR/report.html.")
@click.option("--config", "config_path", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--ai", "enable_ai", is_flag=True, help="Use DeepSeek to add human-readable assistance.")
@click.option("--ai-review", "enable_ai_review", is_flag=True, help="Use DeepSeek to add structured manual review findings.")
def check(
    template_path: Path,
    input_path: Path,
    out_dir: Path,
    json_path: Path | None,
    markdown_path: Path | None,
    html_path: Path | None,
    config_path: Path | None,
    enable_ai: bool,
    enable_ai_review: bool,
) -> None:
    """Check rigid v1 rules for one document."""

    converter = WordConverter()
    extractor = DocxExtractor()
    rule_extractor = WordCountRuleExtractor()
    checker = DocumentChecker()

    try:
        template_copy = converter.normalize_to_docx(template_path, out_dir, "template")
        input_copy = converter.normalize_to_docx(input_path, out_dir, "input")
    except ConversionError as exc:
        raise click.ClickException(str(exc)) from exc

    template_snapshot = extractor.extract(template_copy.docx_path)
    input_snapshot = extractor.extract(input_copy.docx_path)
    rules = rule_extractor.extract(template_snapshot)
    format_rules = None
    if config_path is not None:
        try:
            config = load_config(config_path)
            rules = apply_rule_config(rules, config.word_count_rules)
            format_rules = config.format_rules or None
        except ValueError as exc:
            raise click.ClickException(str(exc)) from exc
    issues = checker.check(template_snapshot, input_snapshot, rules, format_rules=format_rules)
    report = CheckReport(
        template_path=str(template_path),
        input_path=str(input_path),
        template_docx_path=str(template_copy.docx_path),
        input_docx_path=str(input_copy.docx_path),
        issues=issues,
    )

    if enable_ai or enable_ai_review:
        report = _with_ai_assistance(report, include_review=enable_ai_review)

    annotated_path, annotation_warnings = write_annotated_docx(report, out_dir / "input.annotated.docx")
    report = replace(
        report,
        annotated_docx_path=str(annotated_path),
        annotation_warnings=annotation_warnings,
    )

    json_path, markdown_path, html_path = _resolve_report_paths(out_dir, json_path, markdown_path, html_path)
    render_report(report)
    write_json_report(report, json_path)
    write_markdown_report(report, markdown_path)
    write_html_report(report, html_path)

    if not report.passed:
        raise click.exceptions.Exit(1)


@main.command()
@click.option("--template", "template_path", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--input", "input_path", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--out-dir", type=click.Path(file_okay=False, path_type=Path), default=Path("output"), show_default=True)
@click.option("--config", "config_path", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def correct(
    template_path: Path,
    input_path: Path,
    out_dir: Path,
    config_path: Path | None,
) -> None:
    """Create an automatically corrected .docx copy and verification reports."""

    converter = WordConverter()
    extractor = DocxExtractor()
    rule_extractor = WordCountRuleExtractor()
    checker = DocumentChecker()
    corrector = DocxCorrector()

    try:
        template_copy = converter.normalize_to_docx(template_path, out_dir, "template")
        input_copy = converter.normalize_to_docx(input_path, out_dir, "input")
    except ConversionError as exc:
        raise click.ClickException(str(exc)) from exc

    template_snapshot = extractor.extract(template_copy.docx_path)
    input_snapshot = extractor.extract(input_copy.docx_path)
    rules = rule_extractor.extract(template_snapshot)
    format_rules = None
    if config_path is not None:
        try:
            config = load_config(config_path)
            rules = apply_rule_config(rules, config.word_count_rules)
            format_rules = config.format_rules or None
        except ValueError as exc:
            raise click.ClickException(str(exc)) from exc

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
    except ConversionError as exc:
        corrected_doc_export_error = str(exc)

    correction_report = replace(
        correction_report,
        template_path=str(template_path),
        input_path=str(input_path),
        corrected_doc_path=corrected_doc_path_for_report,
        corrected_doc_export_error=corrected_doc_export_error,
        corrected_check_report_path=str(corrected_check_report_path),
        corrected_doc_check_report_path=(
            str(corrected_doc_check_report_path)
            if corrected_doc_export_error is None and corrected_doc_check_report_path.exists()
            else None
        ),
    )
    write_correction_report(correction_report, out_dir / "correction_report.json")

    click.echo(f"Corrected DOCX: {corrected_docx_path}")
    if correction_report.corrected_doc_path:
        click.echo(f"Corrected DOC: {correction_report.corrected_doc_path}")
    if correction_report.corrected_doc_export_error:
        click.echo(f"DOC export skipped/failed: {correction_report.corrected_doc_export_error}")


def _with_ai_assistance(report: CheckReport, include_review: bool = False) -> CheckReport:
    try:
        assistance = DeepSeekAssistant().analyze(report, include_review=include_review)
    except (AiConfigError, Exception) as exc:
        return replace(report, ai_error=f"AI 辅助不可用：{exc}")
    return replace(
        report,
        ai_summary=assistance.summary,
        ai_suggestions=assistance.suggestions,
        ai_review_findings=assistance.review_findings,
    )


def _resolve_report_paths(
    out_dir: Path,
    json_path: Path | None,
    markdown_path: Path | None,
    html_path: Path | None,
) -> tuple[Path, Path, Path]:
    return (
        json_path or out_dir / "report.json",
        markdown_path or out_dir / "report.md",
        html_path or out_dir / "report.html",
    )


if __name__ == "__main__":
    main()
