from __future__ import annotations

from docxpdf_native.fonts.bundled import bundled_font_paths, metric_compatible_substitution
from docxpdf_native.fonts.cjk import (
    contains_cjk_characters,
    detect_system_cjk_fonts,
    known_japanese_font_category,
)
from docxpdf_native.fonts.metrics import FontToolsTextMeasurer
from docxpdf_native.fonts.registry import FontRecord, FontRegistry
from docxpdf_native.fonts.resolver import DefaultFontResolver

__all__ = [
    "DefaultFontResolver",
    "FontRecord",
    "FontRegistry",
    "FontToolsTextMeasurer",
    "bundled_font_paths",
    "contains_cjk_characters",
    "detect_system_cjk_fonts",
    "known_japanese_font_category",
    "metric_compatible_substitution",
]
