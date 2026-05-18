"""Write automatically corrected .docx copies."""

from __future__ import annotations

from collections import Counter
from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import re
import shutil
from typing import Iterable

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from doc_fix.model import CorrectionAction, CorrectionReport, DocumentSnapshot, FormatRule, ParagraphBlock


BRACKET_REMARK_PATTERN = re.compile(r"【[^】]*】")
DEFAULT_FORMAT_RULES = (
    FormatRule(scope="body"),
    FormatRule(scope="heading"),
)


class DocxCorrector:
    """Apply template-derived corrections to a target .docx copy."""

    def correct(
        self,
        template_docx_path: Path,
        input_docx_path: Path,
        output_path: Path,
        template_snapshot: DocumentSnapshot,
        input_snapshot: DocumentSnapshot,
        format_rules: tuple[FormatRule, ...] | None = None,
    ) -> CorrectionReport:
        """Create a corrected .docx and return a structured report."""

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_docx_path, output_path)

        template_document = Document(str(template_docx_path))
        target_document = Document(str(output_path))
        actions: list[CorrectionAction] = []
        warnings: list[str] = []

        rules = self._format_rule_map(format_rules)
        paragraph_count, paragraph_warnings = self._apply_paragraph_formats(
            template_document,
            target_document,
            template_snapshot,
            input_snapshot,
            rules,
        )
        table_count, table_warnings = self._apply_table_text_formats(
            template_document,
            target_document,
            rules,
        )
        removed_count, removed_paragraphs = self._remove_bracket_remarks(target_document)
        warnings.extend(paragraph_warnings)
        warnings.extend(table_warnings)

        if paragraph_count:
            actions.append(
                CorrectionAction(
                    code="paragraph.format_aligned",
                    message="已按模板对齐正文/标题段落格式。",
                    count=paragraph_count,
                )
            )
        if table_count:
            actions.append(
                CorrectionAction(
                    code="table.text_format_aligned",
                    message="已按模板对齐表格单元格文字格式。",
                    count=table_count,
                )
            )
        if removed_count:
            actions.append(
                CorrectionAction(
                    code="remarks.bracket_removed",
                    message="已删除【】包裹的备注内容。",
                    count=removed_count,
                )
            )
        if removed_paragraphs:
            actions.append(
                CorrectionAction(
                    code="remarks.empty_paragraph_removed",
                    message="已删除备注清空后的正文空段落。",
                    count=removed_paragraphs,
                )
            )

        target_document.save(output_path)
        return CorrectionReport(
            template_path=template_snapshot.source_path,
            input_path=input_snapshot.source_path,
            template_docx_path=str(template_docx_path),
            input_docx_path=str(input_docx_path),
            corrected_docx_path=str(output_path),
            actions=tuple(actions),
            warnings=tuple(warnings),
        )

    def _apply_paragraph_formats(
        self,
        template_document,
        target_document,
        template_snapshot: DocumentSnapshot,
        input_snapshot: DocumentSnapshot,
        rules: dict[str, FormatRule],
    ) -> tuple[int, list[str]]:
        references = self._paragraph_format_references(template_snapshot, rules)
        corrected = 0
        warnings: list[str] = []

        for paragraph in input_snapshot.paragraphs:
            scope = self._paragraph_scope(paragraph)
            if scope is None:
                continue
            rule = rules.get(scope)
            if rule is None or not rule.enabled:
                continue
            reference = self._matching_paragraph_reference(paragraph, references)
            if reference is None:
                continue
            if reference.index >= len(template_document.paragraphs) or paragraph.index >= len(target_document.paragraphs):
                warnings.append(f"paragraph {paragraph.index + 1}: 未能定位模板或目标段落，跳过格式修正。")
                continue
            self._copy_paragraph_format(
                template_document.paragraphs[reference.index],
                target_document.paragraphs[paragraph.index],
                rule,
            )
            corrected += 1
        return corrected, warnings

    def _apply_table_text_formats(self, template_document, target_document, rules: dict[str, FormatRule]) -> tuple[int, list[str]]:
        body_rule = rules.get("body")
        if body_rule is None or not body_rule.enabled:
            return 0, []

        corrected = 0
        warnings: list[str] = []
        for table_index, target_table in enumerate(target_document.tables):
            if table_index >= len(template_document.tables):
                warnings.append(f"table {table_index + 1}: 模板缺少对应表格，跳过表格文字格式修正。")
                continue
            template_table = template_document.tables[table_index]
            fallback = self._first_non_empty_table_paragraph(template_table)
            if fallback is None:
                continue
            for row_index, target_row in enumerate(target_table.rows):
                for cell_index, target_cell in enumerate(target_row.cells):
                    reference = self._matching_table_cell_paragraph(template_table, row_index, cell_index) or fallback
                    for paragraph in target_cell.paragraphs:
                        if not paragraph.text.strip():
                            continue
                        self._copy_paragraph_format(reference, paragraph, body_rule)
                        corrected += 1
        return corrected, warnings

    def _remove_bracket_remarks(self, document) -> tuple[int, int]:
        removed = 0
        removed_empty_body_paragraphs = 0

        for paragraph in list(document.paragraphs):
            paragraph_removed = _remove_bracket_remarks_from_paragraph(paragraph)
            removed += paragraph_removed
            if paragraph_removed and not paragraph.text.strip() and not _paragraph_has_drawing(paragraph):
                _delete_paragraph(paragraph)
                removed_empty_body_paragraphs += 1

        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        removed += _remove_bracket_remarks_from_paragraph(paragraph)
        return removed, removed_empty_body_paragraphs

    def _copy_paragraph_format(self, reference: Paragraph, target: Paragraph, rule: FormatRule) -> None:
        if rule.check_alignment:
            target.alignment = reference.alignment
        if rule.check_indent:
            target.paragraph_format.first_line_indent = reference.paragraph_format.first_line_indent
            target.paragraph_format.left_indent = reference.paragraph_format.left_indent
            target.paragraph_format.right_indent = reference.paragraph_format.right_indent
        if rule.check_line_spacing:
            target.paragraph_format.line_spacing = reference.paragraph_format.line_spacing
        if rule.check_spacing:
            target.paragraph_format.space_before = reference.paragraph_format.space_before
            target.paragraph_format.space_after = reference.paragraph_format.space_after

        reference_run = _first_text_run(reference) or (reference.runs[0] if reference.runs else None)
        if reference_run is None:
            return
        for run in _text_runs(target):
            self._copy_run_format(reference_run, run, rule)

    def _copy_run_format(self, reference: Run, target: Run, rule: FormatRule) -> None:
        reference_r_pr = reference._r.rPr
        target_r_pr = target._r.get_or_add_rPr()
        if rule.check_font_names:
            _replace_run_property(reference_r_pr, target_r_pr, "rFonts")
        if rule.check_font_sizes:
            _replace_run_property(reference_r_pr, target_r_pr, "sz")
            _replace_run_property(reference_r_pr, target_r_pr, "szCs")
        if rule.check_font_colors:
            _replace_run_property(reference_r_pr, target_r_pr, "color")
        if rule.check_font_highlights:
            _replace_run_property(reference_r_pr, target_r_pr, "highlight")
            _replace_run_property(reference_r_pr, target_r_pr, "shd")

    def _format_rule_map(self, configured_rules: tuple[FormatRule, ...] | None) -> dict[str, FormatRule]:
        rules = {rule.scope: rule for rule in DEFAULT_FORMAT_RULES}
        for rule in configured_rules or ():
            rules[rule.scope] = rule
        return rules

    def _paragraph_format_references(
        self,
        template: DocumentSnapshot,
        rules: dict[str, FormatRule],
    ) -> dict[tuple[str, int | None, str | None], ParagraphBlock]:
        references: dict[tuple[str, int | None, str | None], ParagraphBlock] = {}
        fallback_candidates: dict[tuple[str, int | None], list[ParagraphBlock]] = {}
        for paragraph in template.paragraphs:
            scope = self._paragraph_scope(paragraph)
            if scope is None:
                continue
            rule = rules.get(scope)
            if rule is None or not rule.enabled or not self._has_paragraph_format(paragraph):
                continue
            exact_key = self._paragraph_format_key(paragraph)
            references.setdefault(exact_key, paragraph)
            fallback_candidates.setdefault((exact_key[0], exact_key[1]), []).append(paragraph)
        for (scope, heading_level), candidates in fallback_candidates.items():
            representative = most_common_paragraph_format(candidates)
            if representative is not None:
                references[(scope, heading_level, None)] = representative
        return references

    def _matching_paragraph_reference(
        self,
        paragraph: ParagraphBlock,
        references: dict[tuple[str, int | None, str | None], ParagraphBlock],
    ) -> ParagraphBlock | None:
        exact_key = self._paragraph_format_key(paragraph)
        exact_reference = references.get(exact_key)
        if exact_reference is not None:
            return exact_reference
        if paragraph.chapter_path is None:
            return None
        return references.get((exact_key[0], exact_key[1], None))

    def _paragraph_format_key(self, paragraph: ParagraphBlock) -> tuple[str, int | None, str | None]:
        scope = self._paragraph_scope(paragraph) or "body"
        heading_level = paragraph.heading_level if scope == "heading" else None
        chapter_key = normalize_chapter_path_key(paragraph.chapter_path)
        if chapter_key is None:
            chapter_key = normalize_paragraph_text_key(paragraph.text)
        return (scope, heading_level, chapter_key)

    def _paragraph_scope(self, paragraph: ParagraphBlock) -> str | None:
        if not paragraph.text:
            return None
        return "heading" if paragraph.is_heading or paragraph.heading_level is not None else "body"

    def _has_paragraph_format(self, paragraph: ParagraphBlock) -> bool:
        return any(
            (
                paragraph.font_names,
                paragraph.font_sizes_pt,
                paragraph.alignment,
                paragraph.first_line_indent_twips is not None,
                paragraph.left_indent_twips is not None,
                paragraph.right_indent_twips is not None,
                paragraph.line_spacing is not None,
                paragraph.space_before_twips is not None,
                paragraph.space_after_twips is not None,
            )
        )

    def _first_non_empty_table_paragraph(self, table) -> Paragraph | None:
        for paragraph in _iter_table_paragraphs(table):
            if paragraph.text.strip() and paragraph.runs:
                return paragraph
        return None

    def _matching_table_cell_paragraph(self, table, row_index: int, cell_index: int) -> Paragraph | None:
        if row_index >= len(table.rows) or cell_index >= len(table.rows[row_index].cells):
            return None
        for paragraph in table.rows[row_index].cells[cell_index].paragraphs:
            if paragraph.text.strip() and paragraph.runs:
                return paragraph
        return None


def write_correction_report(report: CorrectionReport, output_path: Path) -> None:
    """Write correction_report.json."""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(asdict(report), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _remove_bracket_remarks_from_paragraph(paragraph: Paragraph) -> int:
    text_nodes = list(paragraph._p.xpath(".//w:t"))
    if not text_nodes:
        return 0
    full_text = "".join(node.text or "" for node in text_nodes)
    matches = list(BRACKET_REMARK_PATTERN.finditer(full_text))
    if not matches:
        return 0

    keep = [True] * len(full_text)
    for match in matches:
        for index in range(match.start(), match.end()):
            keep[index] = False

    offset = 0
    for node in text_nodes:
        value = node.text or ""
        length = len(value)
        node.text = "".join(char for char, should_keep in zip(value, keep[offset : offset + length], strict=True) if should_keep)
        offset += length
    return len(matches)


def _delete_paragraph(paragraph: Paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _paragraph_has_drawing(paragraph: Paragraph) -> bool:
    return bool(paragraph._p.xpath(".//w:drawing"))


def _first_text_run(paragraph: Paragraph) -> Run | None:
    for run in paragraph.runs:
        if run.text.strip():
            return run
    return None


def _text_runs(paragraph: Paragraph) -> Iterable[Run]:
    for run in paragraph.runs:
        if run.text:
            yield run


def _replace_run_property(reference_r_pr, target_r_pr, tag: str) -> None:
    for child in list(target_r_pr):
        if child.tag == qn(f"w:{tag}"):
            target_r_pr.remove(child)
    if reference_r_pr is None:
        return
    reference_children = reference_r_pr.xpath(f"./w:{tag}")
    if reference_children:
        target_r_pr.append(deepcopy(reference_children[0]))


def _iter_table_paragraphs(table) -> Iterable[Paragraph]:
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs


def normalize_chapter_path_key(chapter_path: str | None) -> str | None:
    if not chapter_path:
        return None
    return "".join(chapter_path.split())


def normalize_paragraph_text_key(text: str | None) -> str | None:
    if not text:
        return None
    return "text:" + "".join(text.split())


def most_common_paragraph_format(paragraphs: list[ParagraphBlock]) -> ParagraphBlock | None:
    if not paragraphs:
        return None
    signatures = Counter(paragraph_format_signature(paragraph) for paragraph in paragraphs)
    signature, _ = signatures.most_common(1)[0]
    for paragraph in paragraphs:
        if paragraph_format_signature(paragraph) == signature:
            return paragraph
    return None


def paragraph_format_signature(paragraph: ParagraphBlock) -> tuple:
    return (
        paragraph.font_names,
        paragraph.font_sizes_pt,
        paragraph.alignment,
        paragraph.first_line_indent_twips,
        paragraph.left_indent_twips,
        paragraph.right_indent_twips,
        paragraph.line_spacing,
        paragraph.space_before_twips,
        paragraph.space_after_twips,
    )
