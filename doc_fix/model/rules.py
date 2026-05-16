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
