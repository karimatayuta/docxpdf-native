from __future__ import annotations

import math
from pathlib import Path

from fontTools.ttLib import TTFont, TTLibError
from pydantic import Field
from reportlab.pdfbase import pdfmetrics

from docxpdf_native.abstractions import TextMeasurer
from docxpdf_native.exceptions import FontNotFoundError
from docxpdf_native.layout.japanese_breaking import UnicodeText
from docxpdf_native.models.base import FrozenModel
from docxpdf_native.models.fonts import ResolvedFont, TextMeasurement


class _LoadedFont(FrozenModel):
    units_per_em: int = Field(gt=0)
    ascent_units: int
    descent_units: int
    cmap: dict[int, str]
    advances: dict[str, int]
    missing_advance: int = Field(ge=0)


class FontToolsTextMeasurer(TextMeasurer):
    """Measure standard PDF fonts or concrete OpenType/TrueType glyph metrics."""

    def __init__(self) -> None:
        self._font_cache: dict[Path, _LoadedFont] = {}
        self._measurement_cache: dict[
            tuple[str, str, str | None, float, float], TextMeasurement
        ] = {}

    def measure(
        self,
        text: str,
        font: ResolvedFont,
        font_size: float,
        *,
        character_spacing: float = 0,
    ) -> TextMeasurement:
        if not math.isfinite(font_size) or font_size <= 0:
            raise ValueError("font_size must be a positive finite number")
        if not math.isfinite(character_spacing):
            raise ValueError("character_spacing must be finite")
        normalized = UnicodeText.normalize(text)
        path_key = str(font.path) if font.path is not None else None
        key = (normalized, font.family, path_key, font_size, character_spacing)
        cached = self._measurement_cache.get(key)
        if cached is not None:
            return cached

        if font.path is None:
            measured = self._measure_standard(normalized, font, font_size)
        else:
            measured = self._measure_file(normalized, font.path, font_size)

        cluster_count = len(UnicodeText.grapheme_clusters(normalized))
        spacing = max(cluster_count - 1, 0) * character_spacing
        result = measured.model_copy(update={"width": max(measured.width + spacing, 0.0)})
        self._measurement_cache[key] = result
        return result

    @staticmethod
    def _measure_standard(text: str, font: ResolvedFont, font_size: float) -> TextMeasurement:
        name = font.postscript_name or font.family
        try:
            width = float(pdfmetrics.stringWidth(text, name, font_size))
            ascent, descent = pdfmetrics.getAscentDescent(name, font_size)
        except (KeyError, ValueError) as error:
            raise FontNotFoundError(
                f"ReportLab standard font is not registered: {name}",
                requested_font=name,
                cause=error,
            ) from error
        return TextMeasurement(
            width=width,
            ascent=max(float(ascent), 0.0),
            descent=abs(float(descent)),
        )

    def _measure_file(self, text: str, path: Path, font_size: float) -> TextMeasurement:
        data = self._font_cache.get(path)
        if data is None:
            data = self._load_font(path)
            self._font_cache[path] = data
        advance_units = 0
        for character in text:
            codepoint = ord(character)
            glyph = data.cmap.get(codepoint)
            if glyph is None and self._zero_width_when_unmapped(character):
                continue
            advance_units += data.advances.get(glyph or ".notdef", data.missing_advance)
        scale = font_size / data.units_per_em
        return TextMeasurement(
            width=advance_units * scale,
            ascent=max(data.ascent_units * scale, 0.0),
            descent=abs(data.descent_units * scale),
        )

    @staticmethod
    def _load_font(path: Path) -> _LoadedFont:
        try:
            with TTFont(path, lazy=False) as font:
                units_per_em = int(font["head"].unitsPerEm)
                hhea = font["hhea"]
                raw_metrics = font["hmtx"].metrics
                advances = {name: int(values[0]) for name, values in raw_metrics.items()}
                missing_advance = advances.get(".notdef", units_per_em)
                return _LoadedFont(
                    units_per_em=units_per_em,
                    ascent_units=int(hhea.ascent),
                    descent_units=int(hhea.descent),
                    cmap=dict(font.getBestCmap() or {}),
                    advances=advances,
                    missing_advance=missing_advance,
                )
        except (OSError, KeyError, TTLibError) as error:
            raise FontNotFoundError(
                f"font metrics could not be read: {path}",
                requested_font=path.name,
                cause=error,
            ) from error

    @staticmethod
    def _zero_width_when_unmapped(character: str) -> bool:
        codepoint = ord(character)
        return (
            character == "\u200d"
            or 0xFE00 <= codepoint <= 0xFE0F
            or 0xE0100 <= codepoint <= 0xE01EF
        )
