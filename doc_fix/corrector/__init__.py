"""Document correction APIs."""

from doc_fix.corrector.docx_corrector import DocxCorrector, write_correction_report

__all__ = [
    "DocxCorrector",
    "write_correction_report",
]
