from __future__ import annotations

import pytest

from docxpdf_native.exceptions import InvalidOoxmlError
from docxpdf_native.ooxml.content_types import ContentTypesParser
from docxpdf_native.ooxml.xml import SafeXmlParser

CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def parser() -> ContentTypesParser:
    return ContentTypesParser(SafeXmlParser(max_part_size=4096, max_depth=16))


def test_content_types_prefers_override_then_extension_default() -> None:
    xml = f"""<Types xmlns="{CT}">
      <Default Extension="XML" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/main+xml"/>
    </Types>""".encode()

    content_types = parser().parse(xml)

    assert content_types.for_part("word/document.xml") == "application/main+xml"
    assert content_types.for_part("word/styles.XML") == "application/xml"
    assert content_types.for_part("word/media/image") is None


@pytest.mark.parametrize(
    "xml",
    [
        b"<root />",
        f'<Types xmlns="{CT}"><Default ContentType="application/xml"/></Types>'.encode(),
        (
            f'<Types xmlns="{CT}"><Default Extension="xml" ContentType="a"/>'
            '<Default Extension="XML" ContentType="b"/></Types>'
        ).encode(),
        (
            f'<Types xmlns="{CT}"><Override PartName="word/document.xml" ContentType="a"/></Types>'
        ).encode(),
        (
            f'<Types xmlns="{CT}"><Override PartName="/word/document.xml" ContentType="a"/>'
            '<Override PartName="/word/document.xml" ContentType="b"/></Types>'
        ).encode(),
        f'<Types xmlns="{CT}"><Unexpected/></Types>'.encode(),
    ],
)
def test_content_types_rejects_invalid_manifest(xml: bytes) -> None:
    with pytest.raises(InvalidOoxmlError):
        parser().parse(xml)
