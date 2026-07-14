"""Native, deterministic DOCX-to-PDF conversion."""

from __future__ import annotations

from docxpdf_native.converter import Converter
from docxpdf_native.models import ConversionOptions, ConversionResult, FontConfiguration

__version__ = "0.1.0"

__all__ = [
    "ConversionOptions",
    "ConversionResult",
    "Converter",
    "FontConfiguration",
    "__version__",
]
