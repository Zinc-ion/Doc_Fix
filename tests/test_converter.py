from pathlib import Path

import pytest

from doc_fix.converter import ConversionError, WordConverter


def test_normalize_docx_copies_to_output(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"fake docx bytes")

    result = WordConverter().normalize_to_docx(source, tmp_path / "out", "template")

    assert result.converted is False
    assert result.docx_path.name == "template.converted.docx"
    assert result.docx_path.read_bytes() == b"fake docx bytes"


def test_normalize_rejects_unsupported_file_type(tmp_path: Path) -> None:
    source = tmp_path / "source.txt"
    source.write_text("not a document", encoding="utf-8")

    with pytest.raises(ConversionError):
        WordConverter().normalize_to_docx(source, tmp_path / "out", "input")
