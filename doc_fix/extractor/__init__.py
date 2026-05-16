"""Extraction APIs."""

from doc_fix.extractor.docx import DocxExtractor
from doc_fix.extractor.word_count_rules import WordCountRuleExtractor

__all__ = ["DocxExtractor", "WordCountRuleExtractor"]
