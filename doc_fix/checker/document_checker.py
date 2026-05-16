"""Run v1 rigid checks against extracted snapshots."""

from __future__ import annotations

from doc_fix.checker.word_count import count_mixed_words, strip_rule_hint
from doc_fix.model import CheckIssue, DocumentSnapshot, ImageSpec, TableSpec, WordCountRule


TWIPS_TO_EMU = 635
INVALID_RULE_TITLE_TOKENS = ("年 月 日", "20 年", "年月日")


class DocumentChecker:
    """Check word count plus basic table/image format constraints."""

    def check(
        self,
        template: DocumentSnapshot,
        target: DocumentSnapshot,
        word_count_rules: tuple[WordCountRule, ...],
    ) -> tuple[CheckIssue, ...]:
        issues: list[CheckIssue] = []
        issues.extend(self._check_word_counts(target, word_count_rules))
        issues.extend(self._check_tables(template, target))
        issues.extend(self._check_images(template, target))
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

        paragraph_texts = [paragraph.text for paragraph in target.paragraphs]
        for index, rule in enumerate(rules):
            start = self._find_anchor_index(target, rule)
            if start is None:
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
            end = self._next_rule_boundary(target, rules, start, index)
            content = self._section_text(target, start, end, rule)
            actual = count_mixed_words(content)
            anchor = target.paragraphs[start]
            preview = compact_preview(content)
            invalid_title = self._looks_like_invalid_rule_title(rule.section_title)
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
                        nearby_heading=anchor.chapter_path,
                        content_preview=preview,
                        locator=f"目标文档段落 {start + 1}；统计到段落 {end}",
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
        title = rule.section_title
        for paragraph in snapshot.paragraphs:
            if title and title in paragraph.text:
                return paragraph.index
        if rule.source == "config" and rule.paragraph_index is not None and rule.paragraph_index < len(snapshot.paragraphs):
            return rule.paragraph_index
        return None

    def _next_rule_boundary(
        self,
        snapshot: DocumentSnapshot,
        rules: tuple[WordCountRule, ...],
        start: int,
        rule_index: int,
    ) -> int:
        later_rule_indexes = [
            self._find_anchor_index(snapshot, later_rule)
            for later_rule in rules[rule_index + 1 :]
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
    ) -> str:
        parts: list[str] = []
        for paragraph in snapshot.paragraphs[start:end]:
            text = paragraph.text
            if paragraph.index == start and rule.start_offset is not None:
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
