from __future__ import annotations

import pytest

from docxpdf_native.ooxml.namespaces import OoxmlNamespaces


def test_qn_builds_clark_name_for_registered_prefix() -> None:
    assert OoxmlNamespaces.qn("w", "document") == (
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document"
    )


def test_qn_rejects_unknown_prefix() -> None:
    with pytest.raises(ValueError, match="Unknown OOXML namespace prefix"):
        OoxmlNamespaces.qn("unknown", "value")


def test_qn_rejects_invalid_local_name() -> None:
    with pytest.raises(ValueError, match="local name"):
        OoxmlNamespaces.qn("w", "../document")


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, True),
        ("1", True),
        ("true", True),
        ("on", True),
        ("0", False),
        ("false", False),
        ("off", False),
    ],
)
def test_parse_on_off_accepts_ooxml_boolean_values(value: str | None, expected: bool) -> None:
    assert OoxmlNamespaces.parse_on_off(value) is expected


def test_parse_on_off_rejects_invalid_value() -> None:
    with pytest.raises(ValueError, match="Invalid OOXML boolean"):
        OoxmlNamespaces.parse_on_off("yes")


@pytest.mark.parametrize(
    ("value", "expected"),
    [("ff00aa", "FF00AA"), ("#ABCDEF", "ABCDEF"), ("auto", None), (None, None)],
)
def test_normalize_color(value: str | None, expected: str | None) -> None:
    assert OoxmlNamespaces.normalize_color(value) == expected


def test_normalize_color_rejects_non_rgb_value() -> None:
    with pytest.raises(ValueError, match="color"):
        OoxmlNamespaces.normalize_color("red")
