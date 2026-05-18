"""Extracted document structure models."""

from dataclasses import dataclass


@dataclass(frozen=True)
class ParagraphBlock:
    """Paragraph text and coarse structure metadata."""

    index: int
    text: str
    style_name: str | None = None
    heading_level: int | None = None
    chapter_path: str | None = None
    is_heading: bool = False
    font_names: tuple[str, ...] = ()
    font_sizes_pt: tuple[float, ...] = ()
    alignment: str | None = None
    first_line_indent_twips: int | None = None
    left_indent_twips: int | None = None
    right_indent_twips: int | None = None
    line_spacing: float | None = None
    space_before_twips: int | None = None
    space_after_twips: int | None = None


@dataclass(frozen=True)
class TableSpec:
    """Table shape and selected format properties."""

    index: int
    rows: int
    columns: int
    text: str
    chapter_path: str | None = None
    chapter_table_index: int | None = None
    caption: str | None = None
    nearby_heading: str | None = None
    paragraph_index: int | None = None
    width_twips: int | None = None
    has_borders: bool | None = None
    alignment: str | None = None
    style_name: str | None = None
    border_signature: str | None = None
    row_heights_twips: tuple[int | None, ...] = ()
    column_widths_twips: tuple[int | None, ...] = ()
    font_names: tuple[str, ...] = ()
    font_sizes_pt: tuple[float, ...] = ()
    cell_alignments: tuple[str, ...] = ()


@dataclass(frozen=True)
class ImageSpec:
    """Inline image dimensions and surrounding paragraph metadata."""

    index: int
    width_emu: int | None
    height_emu: int | None
    paragraph_index: int | None = None
    paragraph_alignment: str | None = None
    wrap_type: str | None = None
    chapter_path: str | None = None
    chapter_image_index: int | None = None
    caption: str | None = None
    nearby_heading: str | None = None


@dataclass(frozen=True)
class DocumentSnapshot:
    """A normalized view of a .docx document for checking."""

    source_path: str
    paragraphs: tuple[ParagraphBlock, ...]
    tables: tuple[TableSpec, ...]
    images: tuple[ImageSpec, ...]
    page_width_twips: int | None = None
    page_content_width_twips: int | None = None
