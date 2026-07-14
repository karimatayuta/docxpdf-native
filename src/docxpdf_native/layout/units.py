from __future__ import annotations

import math
from decimal import Decimal

from docxpdf_native.ooxml.namespaces import OoxmlNamespaces


class LengthConverter:
    """Convert OOXML integer length units to PDF points."""

    _TWIPS_PER_POINT = Decimal(20)
    _EMU_PER_POINT = Decimal(12_700)
    _HALF_POINTS_PER_POINT = Decimal(2)

    @staticmethod
    def twip_to_point(value: int | float | Decimal) -> float:
        return LengthConverter._divide(value, LengthConverter._TWIPS_PER_POINT)

    @staticmethod
    def emu_to_point(value: int | float | Decimal) -> float:
        return LengthConverter._divide(value, LengthConverter._EMU_PER_POINT)

    @staticmethod
    def half_point_to_point(value: int | float | Decimal) -> float:
        return LengthConverter._divide(value, LengthConverter._HALF_POINTS_PER_POINT)

    @staticmethod
    def _divide(value: int | float | Decimal, divisor: Decimal) -> float:
        if not math.isfinite(float(value)):
            raise ValueError("OOXML length must be finite")
        return float(Decimal(str(value)) / divisor)


class OoxmlValueParser:
    """Parse small OOXML scalar types without scattering rules across parsers."""

    @staticmethod
    def boolean(value: str | None, *, default: bool = True) -> bool:
        return OoxmlNamespaces.parse_on_off(value, default=default)

    @staticmethod
    def color(value: str | None) -> str | None:
        return OoxmlNamespaces.normalize_color(value)
