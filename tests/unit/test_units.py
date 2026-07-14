from __future__ import annotations

import math

import pytest

from docxpdf_native.layout.units import LengthConverter, OoxmlValueParser


@pytest.mark.parametrize(
    ("value", "expected"),
    [(0, 0.0), (20, 1.0), (1440, 72.0), (-720, -36.0)],
)
def test_twip_to_point(value: int, expected: float) -> None:
    assert LengthConverter.twip_to_point(value) == expected


@pytest.mark.parametrize(
    ("value", "expected"),
    [(0, 0.0), (12_700, 1.0), (914_400, 72.0)],
)
def test_emu_to_point(value: int, expected: float) -> None:
    assert LengthConverter.emu_to_point(value) == expected


@pytest.mark.parametrize(
    ("value", "expected"),
    [(0, 0.0), (2, 1.0), (24, 12.0)],
)
def test_half_point_to_point(value: int, expected: float) -> None:
    assert LengthConverter.half_point_to_point(value) == expected


def test_length_converter_rejects_non_finite_values() -> None:
    with pytest.raises(ValueError, match="finite"):
        LengthConverter.twip_to_point(math.inf)


@pytest.mark.parametrize("value", ["true", "1", "on", "TRUE"])
def test_parse_boolean_accepts_ooxml_true_values(value: str) -> None:
    assert OoxmlValueParser.boolean(value) is True


@pytest.mark.parametrize("value", ["false", "0", "off", "FALSE"])
def test_parse_boolean_accepts_ooxml_false_values(value: str) -> None:
    assert OoxmlValueParser.boolean(value) is False


def test_parse_boolean_uses_default_for_missing_attribute() -> None:
    assert OoxmlValueParser.boolean(None, default=True) is True


@pytest.mark.parametrize("value", ["sometimes", "yes", "no"])
def test_parse_boolean_rejects_unknown_token(value: str) -> None:
    with pytest.raises(ValueError, match="OOXML boolean"):
        OoxmlValueParser.boolean(value)


@pytest.mark.parametrize(
    ("value", "expected"),
    [("ff00aa", "FF00AA"), ("#12abef", "12ABEF"), ("auto", None), (None, None)],
)
def test_normalize_color(value: str | None, expected: str | None) -> None:
    assert OoxmlValueParser.color(value) == expected


@pytest.mark.parametrize("value", ["12345", "GG0000", "#1234567", ""])
def test_normalize_color_rejects_invalid_hex(value: str) -> None:
    with pytest.raises(ValueError, match="color"):
        OoxmlValueParser.color(value)
