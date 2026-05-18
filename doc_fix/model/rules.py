"""Rule data models."""

from dataclasses import dataclass


@dataclass(frozen=True)
class CountPolicy:
    """Text counting policy for one section."""

    exclude_tables: bool = False


@dataclass(frozen=True)
class WordCountRule:
    """A section word-count limit inferred from a template or config."""

    section_title: str
    limit: int
    anchor_text: str
    source: str = "template"
    policy: CountPolicy = CountPolicy()
    paragraph_index: int | None = None
    start_offset: int | None = None
    heading_level: int | None = None
    chapter_path: str | None = None


@dataclass(frozen=True)
class FormatRule:
    """Paragraph format comparison options for one paragraph scope."""

    scope: str
    enabled: bool = True
    check_font_names: bool = True
    check_font_sizes: bool = True
    check_font_colors: bool = True
    check_font_highlights: bool = True
    check_alignment: bool = True
    check_indent: bool = True
    check_line_spacing: bool = True
    check_spacing: bool = True
    font_size_tolerance_pt: float = 0.1
    indent_tolerance_twips: int = 20
    line_spacing_tolerance: float = 0.05
    spacing_tolerance_twips: int = 20
