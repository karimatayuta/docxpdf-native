from __future__ import annotations

from io import BytesIO
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest

from docxpdf_native.exceptions import UnsupportedFeatureError
from docxpdf_native.ooxml.package import OoxmlPackage
from docxpdf_native.ooxml.unsupported import (
    LenientUnsupportedFeatureHandler,
    StrictUnsupportedFeatureHandler,
    UnsupportedFeatureDetector,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
V = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def build_package(
    *,
    extra_parts: dict[str, bytes] | None = None,
    document_relationships: str = "",
) -> OoxmlPackage:
    parts = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}"><Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            "</Types>"
        ).encode(),
        "_rels/.rels": (
            f'<Relationships xmlns="{PR}"><Relationship Id="rId1" Type="{R}/officeDocument" '
            'Target="word/document.xml"/></Relationships>'
        ).encode(),
        "word/document.xml": f'<w:document xmlns:w="{W}"><w:body/></w:document>'.encode(),
    }
    if document_relationships:
        parts["word/_rels/document.xml.rels"] = (
            f'<Relationships xmlns="{PR}">{document_relationships}</Relationships>'
        ).encode()
    parts.update(extra_parts or {})
    data = BytesIO()
    with ZipFile(data, "w") as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return OoxmlPackage.open(data.getvalue(), allow_external_relationships=True)


@pytest.mark.parametrize(
    ("xml", "feature_name"),
    [
        (f'<w:root xmlns:w="{W}"><w:txbxContent/></w:root>', "text_box"),
        (f'<w:root xmlns:w="{W}" xmlns:dgm="{DGM}"><dgm:relIds/></w:root>', "smart_art"),
        (f'<w:root xmlns:w="{W}" xmlns:c="{C}"><c:chart/></w:root>', "chart"),
        (f'<w:root xmlns:w="{W}"><w:object/></w:root>', "embedded_object"),
        (f'<w:root xmlns:w="{W}" xmlns:o="{O_NS}"><o:OLEObject/></w:root>', "ole"),
        (f'<w:root xmlns:w="{W}" xmlns:m="{M}"><m:oMath/></w:root>', "math"),
        (f'<w:root xmlns:w="{W}" xmlns:v="{V}"><v:textpath/></w:root>', "word_art"),
        (f'<w:root xmlns:w="{W}" xmlns:v="{V}"><v:shape/></w:root>', "floating_shape"),
        (f'<w:root xmlns:w="{W}" xmlns:wp="{WP}"><wp:anchor/></w:root>', "anchor_image"),
        (f'<w:root xmlns:w="{W}"><w:textDirection w:val="tbRl"/></w:root>', "vertical_writing"),
        (f'<w:root xmlns:w="{W}"><w:ruby/></w:root>', "ruby"),
        (f'<w:root xmlns:w="{W}"><w:bidi/></w:root>', "bidirectional_layout"),
        (f'<w:root xmlns:w="{W}"><w:cs/></w:root>', "complex_arabic_shaping"),
        (f'<w:root xmlns:w="{W}"><w:ins/></w:root>', "tracked_changes"),
        (f'<w:root xmlns:w="{W}"><w:commentReference/></w:root>', "comments"),
        (f'<w:root xmlns:w="{W}"><w:sdt/></w:root>', "content_control"),
        (f'<w:root xmlns:w="{W}"><w:cols w:num="2"/></w:root>', "multiple_columns"),
        (f'<w:root xmlns:w="{W}"><w:footnoteReference/></w:root>', "footnotes"),
        (f'<w:root xmlns:w="{W}"><w:endnoteReference/></w:root>', "endnotes"),
        (f'<w:root xmlns:w="{W}"><w:fldChar w:fldCharType="begin"/></w:root>', "complex_field"),
    ],
)
def test_lenient_detector_records_each_named_unsupported_feature(
    xml: str,
    feature_name: str,
) -> None:
    handler = LenientUnsupportedFeatureHandler()
    detector = UnsupportedFeatureDetector(handler)

    detector.scan_element(ElementTree.fromstring(xml), part_name="word/document.xml")

    assert handler.features[0].name == feature_name


def test_strict_handler_raises_contextual_error() -> None:
    root = ElementTree.fromstring(f'<w:root xmlns:w="{W}"><w:txbxContent/></w:root>')

    with pytest.raises(UnsupportedFeatureError) as caught:
        UnsupportedFeatureDetector(StrictUnsupportedFeatureHandler()).scan_element(
            root,
            part_name="word/document.xml",
        )

    assert caught.value.feature.name == "text_box"
    assert caught.value.feature.part == "word/document.xml"
    assert caught.value.feature.location == "/root[1]/txbxContent[1]"


def test_lenient_handler_creates_warning_for_every_feature() -> None:
    root = ElementTree.fromstring(f'<w:root xmlns:w="{W}"><w:txbxContent/><w:ruby/></w:root>')
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_element(root, part_name="word/document.xml")

    assert [warning.code for warning in handler.warnings] == [
        "unsupported_feature",
        "unsupported_feature",
    ]
    assert [feature.name for feature in handler.features] == ["text_box", "ruby"]


def test_detector_allows_single_column_and_horizontal_text() -> None:
    root = ElementTree.fromstring(
        f'<w:root xmlns:w="{W}"><w:cols w:num="1"/><w:textDirection w:val="lrTb"/></w:root>'
    )
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_element(root, part_name="word/document.xml")

    assert handler.features == ()


def test_detector_allows_simple_page_fields() -> None:
    root = ElementTree.fromstring(f'<w:root xmlns:w="{W}"><w:fldSimple w:instr=" PAGE "/></w:root>')
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_element(root, part_name="word/header1.xml")

    assert handler.features == ()


def test_detector_rejects_other_simple_fields() -> None:
    root = ElementTree.fromstring(f'<w:root xmlns:w="{W}"><w:fldSimple w:instr=" DATE "/></w:root>')
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_element(root, part_name="word/header1.xml")

    assert handler.features[0].name == "complex_field"


@pytest.mark.parametrize(
    ("part_name", "feature_name"),
    [
        ("word/vbaProject.bin", "macro"),
        ("word/embeddings/object1.bin", "embedded_object"),
        ("word/charts/chart1.xml", "chart"),
        ("word/diagrams/data1.xml", "smart_art"),
    ],
)
def test_detector_reports_unsupported_package_parts(
    part_name: str,
    feature_name: str,
) -> None:
    package = build_package(extra_parts={part_name: b"content"})
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_package(package)

    assert handler.features[0].name == feature_name
    assert handler.features[0].part == part_name


def test_detector_ignores_note_and_comment_parts_without_references() -> None:
    package = build_package(
        extra_parts={
            "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"/>'.encode(),
            "word/endnotes.xml": f'<w:endnotes xmlns:w="{W}"/>'.encode(),
            "word/comments.xml": f'<w:comments xmlns:w="{W}"/>'.encode(),
        }
    )
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_package(package)

    assert handler.features == ()


def test_detector_reports_external_link_image_without_fetching() -> None:
    relationship = (
        f'<Relationship Id="rId9" Type="{R}/image" '
        'Target="https://example.invalid/image.png" TargetMode="External"/>'
    )
    package = build_package(document_relationships=relationship)
    handler = LenientUnsupportedFeatureHandler()

    UnsupportedFeatureDetector(handler).scan_package(package)

    assert handler.features[0].name == "external_image"
    assert handler.features[0].location == "relationship:rId9"
