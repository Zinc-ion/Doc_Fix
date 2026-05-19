"""Word-based document conversion."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path


WORD_FORMAT_DOC = 0
WORD_FORMAT_DOCX = 16


class ConversionError(RuntimeError):
    """Raised when a document cannot be converted."""


@dataclass(frozen=True)
class ConversionResult:
    """Input document and its normalized .docx copy."""

    original_path: Path
    docx_path: Path
    converted: bool


class WordConverter:
    """Normalize .doc/.docx inputs to .docx files.

    COM automation is intentionally isolated here so extraction/checking code
    can remain pure .docx logic.
    """

    def normalize_to_docx(self, input_path: Path, out_dir: Path, label: str) -> ConversionResult:
        input_path = input_path.resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        if not input_path.exists():
            raise ConversionError(f"File does not exist: {input_path}")

        suffix = input_path.suffix.lower()
        output_path = out_dir / f"{label}.converted.docx"

        if suffix == ".docx":
            if input_path.resolve() != output_path.resolve():
                shutil.copy2(input_path, output_path)
            return ConversionResult(input_path, output_path, converted=False)

        if suffix != ".doc":
            raise ConversionError(f"Unsupported file type '{input_path.suffix}'. Only .doc and .docx are supported.")

        working_source = out_dir / f"{label}.source.doc"
        shutil.copy2(input_path, working_source)
        self._convert_with_word(working_source, output_path, WORD_FORMAT_DOCX)
        return ConversionResult(input_path, output_path, converted=True)

    def export_to_doc(self, docx_path: Path, output_path: Path) -> Path:
        """Export .docx to .doc for future correction workflows."""

        docx_path = docx_path.resolve()
        output_path = output_path.resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self._convert_with_word(docx_path, output_path, WORD_FORMAT_DOC)
        return output_path

    def _convert_with_word(self, input_path: Path, output_path: Path, file_format: int) -> None:
        input_path = input_path.resolve()
        output_path = output_path.resolve()
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ConversionError("pywin32 is required for Word .doc conversion on Windows.") from exc

        com_initialized = False
        word = None
        document = None
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            document = word.Documents.Open(str(input_path), False, True)
            document.SaveAs2(str(output_path), FileFormat=file_format)
        except Exception as exc:  # pragma: no cover - depends on local Word installation
            raise ConversionError(f"Word failed to convert '{input_path}': {exc}") from exc
        finally:
            if document is not None:
                document.Close(False)
            if word is not None:
                word.Quit()
            if com_initialized:
                pythoncom.CoUninitialize()
