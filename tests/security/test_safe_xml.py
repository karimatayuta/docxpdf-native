from __future__ import annotations

import pytest

from docxpdf_native.exceptions import InvalidOoxmlError, ResourceLimitError
from docxpdf_native.ooxml.xml import SafeXmlParser


def test_safe_xml_parses_regular_document() -> None:
    parser = SafeXmlParser(max_part_size=1024, max_depth=8)

    root = parser.parse(b"<root><child>text</child></root>", part_name="word/document.xml")

    assert root.find("child") is not None


def test_safe_xml_rejects_xml_larger_than_limit() -> None:
    parser = SafeXmlParser(max_part_size=8, max_depth=8)

    with pytest.raises(ResourceLimitError, match=r"word/document\.xml"):
        parser.parse(b"<root />more", part_name="word/document.xml")


def test_safe_xml_rejects_excessive_depth() -> None:
    parser = SafeXmlParser(max_part_size=1024, max_depth=2)

    with pytest.raises(ResourceLimitError, match="XML depth"):
        parser.parse(b"<a><b><c /></b></a>", part_name="word/document.xml")


@pytest.mark.parametrize("declaration", [b"<!DOCTYPE root>", b"<!ENTITY entity 'value'>"])
def test_safe_xml_rejects_dtd_and_entity_declarations(declaration: bytes) -> None:
    parser = SafeXmlParser(max_part_size=1024, max_depth=8)

    with pytest.raises(InvalidOoxmlError, match="DTD or entity"):
        parser.parse(declaration + b"<root />", part_name="word/document.xml")


def test_safe_xml_wraps_malformed_xml_with_part_name() -> None:
    parser = SafeXmlParser(max_part_size=1024, max_depth=8)

    with pytest.raises(InvalidOoxmlError, match=r"word/document\.xml"):
        parser.parse(b"<root>", part_name="word/document.xml")


@pytest.mark.parametrize(
    ("max_part_size", "max_depth"),
    [(0, 8), (1024, 0)],
)
def test_safe_xml_rejects_non_positive_limits(max_part_size: int, max_depth: int) -> None:
    with pytest.raises(ValueError):
        SafeXmlParser(max_part_size=max_part_size, max_depth=max_depth)


def test_safe_xml_rejects_empty_part() -> None:
    parser = SafeXmlParser(max_part_size=1024, max_depth=8)

    with pytest.raises(InvalidOoxmlError, match="empty"):
        parser.parse(b"", part_name="word/settings.xml")
