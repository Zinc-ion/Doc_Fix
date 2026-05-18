"""Run v1 rigid checks against extracted snapshots."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass

from doc_fix.checker.word_count import count_mixed_words, strip_rule_hint
from doc_fix.model import CheckIssue, DocumentSnapshot, FormatRule, ImageSpec, ParagraphBlock, TableSpec, WordCountRule


TWIPS_TO_EMU = 635
INVALID_RULE_TITLE_TOKENS = ("年 月 日", "20 年", "年月日")
DEFAULT_FORMAT_RULES = (
    FormatRule(scope="body"),
    FormatRule(scope="heading"),
)


@dataclass(frozen=True)
class _AnchorMatch:
    index: int
    mode: str


class DocumentChecker:
    """Check word count plus basic table/image format constraints."""

    def check(
        self,
        template: DocumentSnapshot,
        target: DocumentSnapshot,
        word_count_rules: tuple[WordCountRule, ...],
        format_rules: tuple[FormatRule, ...] | None = None,
    ) -> tuple[CheckIssue, ...]:
        issues: list[CheckIssue] = []
        issues.extend(self._check_word_counts(target, word_count_rules))
        issues.extend(self._check_tables(template, target))
        issues.extend(self._check_images(template, target))
        issues.extend(self._check_paragraph_formats(template, target, format_rules))
        return tuple(issues)

    def _check_word_counts(
        self,
        target: DocumentSnapshot,
        rules: tuple[WordCountRule, ...],
    ) -> list[CheckIssue]:
        issues: list[CheckIssue] = []
        if not rules:
            return [
                CheckIssue(
                    code="word_count.no_rules",
                    severity="warning",
                    message="未能从模板中抽取字数限制规则，需人工确认。",
                )
            ]

        for index, rule in enumerate(rules):
            match = self._find_anchor(target, rule)
            if match is None:
                issues.append(
                    CheckIssue(
                        code="word_count.anchor_missing",
                        severity="warning",
                        message="目标文档中未能定位该字数规则锚点，需人工确认。",
                        expected=rule.limit,
                        section_title=rule.section_title,
                        source=rule.source,
                        chapter_path="未能定位，需人工确认",
                        content_preview=rule.anchor_text[:120],
                        locator="未能在目标文档中找到规则锚点",
                    )
                )
                continue
            start = match.index
            end = self._next_rule_boundary(target, rules, start, index)
            content_start = start + 1 if match.mode == "chapter_path" else start
            content = self._section_text(
                target,
                content_start,
                end,
                rule,
                use_start_offset=match.mode != "chapter_path",
            )
            actual = count_mixed_words(content)
            anchor = target.paragraphs[start]
            preview = compact_preview(content)
            invalid_title = self._looks_like_invalid_rule_title(rule.section_title)
            locator_prefix = "按章节路径定位：" if match.mode == "chapter_path" else ""
            if actual > rule.limit:
                issues.append(
                    CheckIssue(
                        code="word_count.exceeded",
                        severity="warning" if invalid_title else "error",
                        message=(
                            f"{rule.section_title} 疑似误识别为字数规则，需人工确认。"
                            if invalid_title
                            else f"{rule.section_title} 字数超限，超出 {actual - rule.limit} 字。"
                        ),
                        expected=rule.limit,
                        actual=actual,
                        section_title=rule.section_title,
                        source=rule.source,
                        chapter_path=anchor.chapter_path or "未能定位，需人工确认",
                        paragraph_index=start,
                        end_paragraph_index=max(start, end - 1),
                        nearby_heading=anchor.chapter_path,
                        content_preview=preview,
                        locator=f"{locator_prefix}目标文档段落 {start + 1}；统计到段落 {end}",
                    )
                )
        return issues

    def _check_tables(self, template: DocumentSnapshot, target: DocumentSnapshot) -> list[CheckIssue]:
        issues: list[CheckIssue] = []
        if len(template.tables) != len(target.tables):
            issues.append(
                CheckIssue(
                    code="table.count_mismatch",
                    severity="error",
                    message="目标文档表格数量与模板不一致。",
                    expected=len(template.tables),
                    actual=len(target.tables),
                    chapter_path="全文结构",
                    locator="模板与目标表格总数不同，后续表格差异可能按全局序号兜底匹配。",
                )
            )

        for template_table, target_table, fallback_match in self._iter_table_pairs(template, target):
            if template_table.rows != target_table.rows:
                issues.append(
                    CheckIssue(
                        code="table.rows_mismatch",
                        severity="warning",
                        message=f"第 {template_table.index + 1} 个表格行数与模板不一致。",
                        expected=template_table.rows,
                        actual=target_table.rows,
                        **self._table_location(target_table, fallback=fallback_match),
                    )
                )
            if template_table.columns != target_table.columns:
                issues.append(
                    CheckIssue(
                        code="table.columns_mismatch",
                        severity="error",
                        message=f"第 {template_table.index + 1} 个表格列数与模板不一致。",
                        expected=template_table.columns,
                        actual=target_table.columns,
                        **self._table_location(target_table, fallback=fallback_match),
                    )
                )
            if template_table.has_borders is not None and target_table.has_borders != template_table.has_borders:
                issues.append(
                    CheckIssue(
                        code="table.border_mismatch",
                        severity="warning",
                        message=f"第 {template_table.index + 1} 个表格边框状态与模板不一致。",
                        expected=template_table.has_borders,
                        actual=target_table.has_borders,
                        **self._table_location(target_table, fallback=fallback_match),
                    )
                )
            if (
                template_table.border_signature
                and target_table.border_signature
                and template_table.border_signature != target_table.border_signature
            ):
                issues.append(
                    CheckIssue(
                        code="table.border_signature_mismatch",
                        severity="warning",
                        message=f"第 {template_table.index + 1} 个表格边框线型/线宽与模板不一致。",
                        expected=template_table.border_signature,
                        actual=target_table.border_signature,
                        **self._table_location(target_table, fallback=fallback_match),
                    )
                )
            self._append_scalar_mismatch(
                issues,
                "table.style_mismatch",
                f"第 {template_table.index + 1} 个表格样式与模板不一致。",
                template_table.style_name,
                target_table.style_name,
                self._table_location(target_table, fallback=fallback_match),
            )
            self._append_scalar_mismatch(
                issues,
                "table.alignment_mismatch",
                f"第 {template_table.index + 1} 个表格对齐方式与模板不一致。",
                template_table.alignment,
                target_table.alignment,
                self._table_location(target_table, fallback=fallback_match),
            )
            self._append_sequence_mismatch(
                issues,
                "table.column_widths_mismatch",
                f"第 {template_table.index + 1} 个表格列宽与模板不一致。",
                template_table.column_widths_twips,
                target_table.column_widths_twips,
                tolerance_ratio=0.05,
                location=self._table_location(target_table, fallback=fallback_match),
            )
            self._append_sequence_mismatch(
                issues,
                "table.row_heights_mismatch",
                f"第 {template_table.index + 1} 个表格行高与模板不一致。",
                template_table.row_heights_twips,
                target_table.row_heights_twips,
                tolerance_ratio=0.05,
                location=self._table_location(target_table, fallback=fallback_match),
            )
            self._append_set_mismatch(
                issues,
                "table.font_names_mismatch",
                f"第 {template_table.index + 1} 个表格字体与模板不一致。",
                template_table.font_names,
                target_table.font_names,
                self._table_location(target_table, fallback=fallback_match),
            )
            self._append_set_mismatch(
                issues,
                "table.font_sizes_mismatch",
                f"第 {template_table.index + 1} 个表格字号与模板不一致。",
                template_table.font_sizes_pt,
                target_table.font_sizes_pt,
                self._table_location(target_table, fallback=fallback_match),
            )
            self._append_set_mismatch(
                issues,
                "table.cell_alignments_mismatch",
                f"第 {template_table.index + 1} 个表格单元格段落对齐与模板不一致。",
                template_table.cell_alignments,
                target_table.cell_alignments,
                self._table_location(target_table, fallback=fallback_match),
            )
            template_is_wide = self._table_exceeds_content_width(template, template_table)
            target_is_wide = self._table_exceeds_content_width(target, target_table)
            if target_is_wide and not template_is_wide:
                issues.append(
                    CheckIssue(
                        code="table.too_wide",
                        severity="error",
                        message=f"第 {target_table.index + 1} 个表格宽度超过页面宽度。",
                        expected=f"<= {target.page_content_width_twips}",
                        actual=target_table.width_twips,
                        **self._table_location(target_table, fallback=fallback_match),
                    )
                )
        return issues

    def _check_images(self, template: DocumentSnapshot, target: DocumentSnapshot) -> list[CheckIssue]:
        issues: list[CheckIssue] = []
        if len(template.images) != len(target.images):
            issues.append(
                CheckIssue(
                    code="image.count_mismatch",
                    severity="warning",
                    message="目标文档图片数量与模板不一致。",
                    expected=len(template.images),
                    actual=len(target.images),
                    chapter_path="全文结构",
                    locator="模板与目标图片总数不同，后续图片差异可能按全局序号兜底匹配。",
                )
            )

        for template_image, target_image in zip(template.images, target.images, strict=False):
            if template_image.width_emu and target_image.width_emu:
                tolerance = int(template_image.width_emu * 0.05)
                if abs(template_image.width_emu - target_image.width_emu) > tolerance:
                    issues.append(
                        CheckIssue(
                            code="image.width_mismatch",
                            severity="warning",
                            message=f"第 {template_image.index + 1} 张图片宽度与模板差异超过 5%。",
                            expected=template_image.width_emu,
                            actual=target_image.width_emu,
                            **self._image_location(target_image),
                        )
                    )
            if template_image.height_emu and target_image.height_emu:
                tolerance = int(template_image.height_emu * 0.05)
                if abs(template_image.height_emu - target_image.height_emu) > tolerance:
                    issues.append(
                        CheckIssue(
                            code="image.height_mismatch",
                            severity="warning",
                            message=f"第 {template_image.index + 1} 张图片高度与模板差异超过 5%。",
                            expected=template_image.height_emu,
                            actual=target_image.height_emu,
                            **self._image_location(target_image),
                        )
                    )
            self._append_scalar_mismatch(
                issues,
                "image.wrap_mismatch",
                f"第 {template_image.index + 1} 张图片环绕方式与模板不一致。",
                template_image.wrap_type,
                target_image.wrap_type,
                self._image_location(target_image),
            )
            self._append_scalar_mismatch(
                issues,
                "image.alignment_mismatch",
                f"第 {template_image.index + 1} 张图片所在段落对齐方式与模板不一致。",
                template_image.paragraph_alignment,
                target_image.paragraph_alignment,
                self._image_location(target_image),
            )
            template_is_wide = self._image_exceeds_content_width(template, template_image)
            target_is_wide = self._image_exceeds_content_width(target, target_image)
            if target_is_wide and not template_is_wide:
                issues.append(
                    CheckIssue(
                        code="image.too_wide",
                        severity="error",
                        message=f"第 {target_image.index + 1} 张图片宽度超过页面正文宽度。",
                        expected=f"<= {target.page_content_width_twips * TWIPS_TO_EMU if target.page_content_width_twips else None}",
                        actual=target_image.width_emu,
                        **self._image_location(target_image),
                    )
                )
        return issues

    def _check_paragraph_formats(
        self,
        template: DocumentSnapshot,
        target: DocumentSnapshot,
        configured_rules: tuple[FormatRule, ...] | None,
    ) -> list[CheckIssue]:
        rules = self._format_rule_map(configured_rules)
        references = self._paragraph_format_references(template, rules)
        missing_reference_keys: set[tuple[str, int | None, str | None]] = set()
        issues: list[CheckIssue] = []

        for paragraph in target.paragraphs:
            scope = self._paragraph_scope(paragraph)
            if scope is None:
                continue
            rule = rules.get(scope)
            if rule is None or not rule.enabled:
                continue
            reference = self._matching_paragraph_reference(paragraph, references)
            if reference is None:
                if not self._has_scope_reference(scope, references):
                    continue
                key = self._paragraph_format_key(paragraph)
                if key not in missing_reference_keys:
                    missing_reference_keys.add(key)
                    issues.append(
                        CheckIssue(
                            code="paragraph.format_reference_missing",
                            severity="warning",
                            message="未能在模板中定位可对比的段落格式参考，需人工确认。",
                            section_title=scope,
                            **self._paragraph_location(paragraph),
                        )
                    )
                continue
            self._append_paragraph_font_name_mismatch(issues, rule, reference, paragraph)
            self._append_paragraph_font_size_mismatch(issues, rule, reference, paragraph)
            self._append_paragraph_alignment_mismatch(issues, rule, reference, paragraph)
            self._append_paragraph_indent_mismatch(issues, rule, reference, paragraph)
            self._append_paragraph_line_spacing_mismatch(issues, rule, reference, paragraph)
            self._append_paragraph_spacing_mismatch(issues, rule, reference, paragraph)
        return issues

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

    def _has_scope_reference(
        self,
        scope: str,
        references: dict[tuple[str, int | None, str | None], ParagraphBlock],
    ) -> bool:
        return any(key[0] == scope for key in references)

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

    def _append_paragraph_font_name_mismatch(
        self,
        issues: list[CheckIssue],
        rule: FormatRule,
        expected: ParagraphBlock,
        actual: ParagraphBlock,
    ) -> None:
        if not rule.check_font_names or not expected.font_names or not actual.font_names:
            return
        if set(expected.font_names) == set(actual.font_names):
            return
        issues.append(
            CheckIssue(
                code="paragraph.font_names_mismatch",
                severity="warning",
                message="段落字体与模板不一致。",
                expected=expected.font_names,
                actual=actual.font_names,
                **self._paragraph_location(actual),
            )
        )

    def _append_paragraph_font_size_mismatch(
        self,
        issues: list[CheckIssue],
        rule: FormatRule,
        expected: ParagraphBlock,
        actual: ParagraphBlock,
    ) -> None:
        if not rule.check_font_sizes or not expected.font_sizes_pt or not actual.font_sizes_pt:
            return
        if same_float_set(expected.font_sizes_pt, actual.font_sizes_pt, rule.font_size_tolerance_pt):
            return
        issues.append(
            CheckIssue(
                code="paragraph.font_sizes_mismatch",
                severity="warning",
                message="段落字号与模板不一致。",
                expected=expected.font_sizes_pt,
                actual=actual.font_sizes_pt,
                **self._paragraph_location(actual),
            )
        )

    def _append_paragraph_alignment_mismatch(
        self,
        issues: list[CheckIssue],
        rule: FormatRule,
        expected: ParagraphBlock,
        actual: ParagraphBlock,
    ) -> None:
        if not rule.check_alignment:
            return
        self._append_scalar_mismatch(
            issues,
            "paragraph.alignment_mismatch",
            "段落对齐方式与模板不一致。",
            expected.alignment,
            actual.alignment,
            self._paragraph_location(actual),
        )

    def _append_paragraph_indent_mismatch(
        self,
        issues: list[CheckIssue],
        rule: FormatRule,
        expected: ParagraphBlock,
        actual: ParagraphBlock,
    ) -> None:
        if not rule.check_indent:
            return
        expected_values = (expected.first_line_indent_twips, expected.left_indent_twips, expected.right_indent_twips)
        actual_values = (actual.first_line_indent_twips, actual.left_indent_twips, actual.right_indent_twips)
        if same_optional_numbers(expected_values, actual_values, float(rule.indent_tolerance_twips)):
            return
        issues.append(
            CheckIssue(
                code="paragraph.indent_mismatch",
                severity="warning",
                message="段落缩进与模板不一致。",
                expected=expected_values,
                actual=actual_values,
                **self._paragraph_location(actual),
            )
        )

    def _append_paragraph_line_spacing_mismatch(
        self,
        issues: list[CheckIssue],
        rule: FormatRule,
        expected: ParagraphBlock,
        actual: ParagraphBlock,
    ) -> None:
        if not rule.check_line_spacing:
            return
        if expected.line_spacing is None or actual.line_spacing is None:
            return
        if abs(expected.line_spacing - actual.line_spacing) <= rule.line_spacing_tolerance:
            return
        issues.append(
            CheckIssue(
                code="paragraph.line_spacing_mismatch",
                severity="warning",
                message="段落行距与模板不一致。",
                expected=expected.line_spacing,
                actual=actual.line_spacing,
                **self._paragraph_location(actual),
            )
        )

    def _append_paragraph_spacing_mismatch(
        self,
        issues: list[CheckIssue],
        rule: FormatRule,
        expected: ParagraphBlock,
        actual: ParagraphBlock,
    ) -> None:
        if not rule.check_spacing:
            return
        expected_values = (expected.space_before_twips, expected.space_after_twips)
        actual_values = (actual.space_before_twips, actual.space_after_twips)
        if same_optional_numbers(expected_values, actual_values, float(rule.spacing_tolerance_twips)):
            return
        issues.append(
            CheckIssue(
                code="paragraph.spacing_mismatch",
                severity="warning",
                message="段前/段后间距与模板不一致。",
                expected=expected_values,
                actual=actual_values,
                **self._paragraph_location(actual),
            )
        )

    def _iter_table_pairs(self, template: DocumentSnapshot, target: DocumentSnapshot) -> list[tuple[TableSpec, TableSpec, bool]]:
        unused_targets = list(target.tables)
        pairs: list[tuple[TableSpec, TableSpec, bool]] = []
        for template_table in template.tables:
            matched_index = self._find_matching_table(template_table, unused_targets)
            fallback = matched_index is None
            if fallback:
                if template_table.index >= len(target.tables):
                    continue
                target_table = target.tables[template_table.index]
                if target_table in unused_targets:
                    unused_targets.remove(target_table)
            else:
                target_table = unused_targets.pop(matched_index)
            pairs.append((template_table, target_table, fallback))
        return pairs

    def _find_matching_table(self, template_table: TableSpec, candidates: list[TableSpec]) -> int | None:
        template_key = normalize_match_key(template_table.chapter_path, template_table.caption)
        if not template_key:
            return None
        for index, candidate in enumerate(candidates):
            if normalize_match_key(candidate.chapter_path, candidate.caption) == template_key:
                return index
        caption_key = normalize_match_key(None, template_table.caption)
        if not caption_key:
            return None
        for index, candidate in enumerate(candidates):
            if normalize_match_key(None, candidate.caption) == caption_key:
                return index
        return None

    def _append_scalar_mismatch(
        self,
        issues: list[CheckIssue],
        code: str,
        message: str,
        expected,
        actual,
        location: dict | None = None,
    ) -> None:
        if expected is None or actual is None or expected == actual:
            return
        issues.append(CheckIssue(code=code, severity="warning", message=message, expected=expected, actual=actual, **(location or {})))

    def _append_set_mismatch(
        self,
        issues: list[CheckIssue],
        code: str,
        message: str,
        expected: tuple,
        actual: tuple,
        location: dict | None = None,
    ) -> None:
        if not expected or not actual or set(expected) == set(actual):
            return
        issues.append(CheckIssue(code=code, severity="warning", message=message, expected=expected, actual=actual, **(location or {})))

    def _append_sequence_mismatch(
        self,
        issues: list[CheckIssue],
        code: str,
        message: str,
        expected: tuple[int | None, ...],
        actual: tuple[int | None, ...],
        tolerance_ratio: float,
        location: dict | None = None,
    ) -> None:
        if not expected or not actual or len(expected) != len(actual):
            return
        for expected_value, actual_value in zip(expected, actual, strict=True):
            if expected_value is None or actual_value is None:
                continue
            tolerance = max(1, int(expected_value * tolerance_ratio))
            if abs(expected_value - actual_value) > tolerance:
                issues.append(
                    CheckIssue(code=code, severity="warning", message=message, expected=expected, actual=actual, **(location or {}))
                )
                return

    def _table_exceeds_content_width(self, snapshot: DocumentSnapshot, table: TableSpec) -> bool:
        return (
            snapshot.page_content_width_twips is not None
            and table.width_twips is not None
            and table.width_twips > snapshot.page_content_width_twips
        )

    def _image_exceeds_content_width(self, snapshot: DocumentSnapshot, image: ImageSpec) -> bool:
        return (
            snapshot.page_content_width_twips is not None
            and image.width_emu is not None
            and image.width_emu > snapshot.page_content_width_twips * TWIPS_TO_EMU
        )

    def _find_anchor_index(self, snapshot: DocumentSnapshot, rule: WordCountRule) -> int | None:
        match = self._find_anchor(snapshot, rule)
        return match.index if match else None

    def _find_anchor(self, snapshot: DocumentSnapshot, rule: WordCountRule) -> _AnchorMatch | None:
        title = rule.section_title
        for paragraph in snapshot.paragraphs:
            if title and title in paragraph.text:
                return _AnchorMatch(paragraph.index, "title")
        chapter_key = normalize_chapter_path_key(rule.chapter_path)
        if rule.source == "template" and chapter_key:
            for paragraph in snapshot.paragraphs:
                if not (paragraph.is_heading or paragraph.heading_level is not None):
                    continue
                if normalize_chapter_path_key(paragraph.chapter_path) == chapter_key:
                    return _AnchorMatch(paragraph.index, "chapter_path")
        if rule.source == "config" and rule.paragraph_index is not None and rule.paragraph_index < len(snapshot.paragraphs):
            return _AnchorMatch(rule.paragraph_index, "config")
        return None

    def _next_rule_boundary(
        self,
        snapshot: DocumentSnapshot,
        rules: tuple[WordCountRule, ...],
        start: int,
        rule_index: int,
    ) -> int:
        later_rule_indexes = [
            match.index if match else None
            for later_rule in rules[rule_index + 1 :]
            for match in (self._find_anchor(snapshot, later_rule),)
        ]
        candidates = [index for index in later_rule_indexes if index is not None and index > start]
        rule = rules[rule_index]
        if rule.heading_level is not None:
            heading_candidates = [
                paragraph.index
                for paragraph in snapshot.paragraphs[start + 1 :]
                if paragraph.heading_level is not None and paragraph.heading_level <= rule.heading_level
            ]
            candidates.extend(heading_candidates)
        return min(candidates) if candidates else len(snapshot.paragraphs)

    def _section_text(
        self,
        snapshot: DocumentSnapshot,
        start: int,
        end: int,
        rule: WordCountRule,
        use_start_offset: bool = True,
    ) -> str:
        parts: list[str] = []
        for paragraph in snapshot.paragraphs[start:end]:
            text = paragraph.text
            if use_start_offset and paragraph.index == start and rule.start_offset is not None:
                text = text[rule.start_offset :]
            parts.append(strip_rule_hint(text))
        return "\n".join(parts)

    def _looks_like_invalid_rule_title(self, title: str) -> bool:
        normalized = "".join(title.split())
        return len(normalized) < 4 or any("".join(token.split()) in normalized for token in INVALID_RULE_TITLE_TOKENS)

    def _table_location(self, table: TableSpec, fallback: bool = False) -> dict:
        chapter_path = table.chapter_path or "未能定位，需人工确认"
        chapter_index = f"本章节第 {table.chapter_table_index} 个表格" if table.chapter_table_index else "本章节表格序号未知"
        locator = f"{chapter_index} / 全文第 {table.index + 1} 个表格"
        if fallback:
            locator += "；按全局序号兜底匹配"
        return {
            "chapter_path": chapter_path,
            "table_index": table.index,
            "nearby_heading": table.nearby_heading,
            "caption": table.caption,
            "content_preview": compact_preview(table.text),
            "locator": locator,
        }

    def _image_location(self, image: ImageSpec) -> dict:
        chapter_index = f"本章节第 {image.chapter_image_index} 张图片" if image.chapter_image_index else "本章节图片序号未知"
        return {
            "chapter_path": image.chapter_path or "未能定位，需人工确认",
            "image_index": image.index,
            "paragraph_index": image.paragraph_index,
            "nearby_heading": image.nearby_heading,
            "caption": image.caption,
            "locator": f"{chapter_index} / 全文第 {image.index + 1} 张图片",
        }

    def _paragraph_location(self, paragraph: ParagraphBlock) -> dict:
        return {
            "chapter_path": paragraph.chapter_path or "未能定位，需人工确认",
            "paragraph_index": paragraph.index,
            "nearby_heading": paragraph.chapter_path,
            "content_preview": compact_preview(paragraph.text),
            "locator": f"目标文档段落 {paragraph.index + 1}",
        }


def compact_preview(text: str | None, limit: int = 120) -> str | None:
    if not text:
        return None
    compact = " ".join(text.split())
    return compact[:limit] + ("..." if len(compact) > limit else "")


def normalize_match_key(chapter_path: str | None, caption: str | None) -> str | None:
    if not caption:
        return None
    parts = [chapter_path or "", caption]
    return "".join("".join(parts).split())


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


def same_float_set(expected: tuple[float, ...], actual: tuple[float, ...], tolerance: float) -> bool:
    if len(expected) != len(actual):
        return False
    unmatched = list(actual)
    for expected_value in expected:
        for index, actual_value in enumerate(unmatched):
            if abs(expected_value - actual_value) <= tolerance:
                unmatched.pop(index)
                break
        else:
            return False
    return True


def same_optional_numbers(
    expected: tuple[int | float | None, ...],
    actual: tuple[int | float | None, ...],
    tolerance: float,
) -> bool:
    comparable = [
        (expected_value, actual_value)
        for expected_value, actual_value in zip(expected, actual, strict=True)
        if expected_value is not None and actual_value is not None
    ]
    if not comparable:
        return True
    return all(abs(float(expected_value) - float(actual_value)) <= tolerance for expected_value, actual_value in comparable)
