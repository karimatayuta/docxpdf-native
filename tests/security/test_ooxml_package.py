from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from docxpdf_native.exceptions import InvalidDocxError, MissingPartError, ResourceLimitError
from docxpdf_native.ooxml.package import OoxmlPackage
from docxpdf_native.ooxml.relationships import Relationship, RelationshipSet
from docxpdf_native.security.limits import ResourceLimits

CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""
DOCUMENT = b"""<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>"""


def build_docx(extra_parts: dict[str, bytes] | None = None) -> bytes:
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/document.xml": DOCUMENT,
    }
    parts.update(extra_parts or {})
    target = BytesIO()
    with ZipFile(target, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return target.getvalue()


def test_package_reads_required_parts_from_bytes() -> None:
    package = OoxmlPackage.open(build_docx())

    assert package.read_part("word/document.xml") == DOCUMENT


def test_package_accepts_path_and_binary_stream(tmp_path: Path) -> None:
    source = tmp_path / "input.docx"
    source.write_bytes(build_docx())

    from_path = OoxmlPackage.open(source)
    from_stream = OoxmlPackage.open(BytesIO(source.read_bytes()))

    assert from_path.part_names == from_stream.part_names


def test_package_rejects_invalid_zip() -> None:
    with pytest.raises(InvalidDocxError, match="ZIP"):
        OoxmlPackage.open(b"not a zip")


def test_package_rejects_missing_required_part() -> None:
    target = BytesIO()
    with ZipFile(target, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)

    with pytest.raises(MissingPartError, match=r"word/document\.xml"):
        OoxmlPackage.open(target.getvalue())


@pytest.mark.parametrize(
    "unsafe_name", ["../outside.xml", "/absolute.xml", "word\\document.xml", "C:/evil.xml"]
)
def test_package_rejects_unsafe_part_names(unsafe_name: str) -> None:
    target = BytesIO()
    with ZipFile(target, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr(unsafe_name, b"unsafe")

    with pytest.raises(InvalidDocxError, match="Unsafe ZIP part name"):
        OoxmlPackage.open(target.getvalue())


def test_package_rejects_duplicate_part_names() -> None:
    target = BytesIO()
    with ZipFile(target, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        with pytest.warns(UserWarning, match="Duplicate name"):
            archive.writestr("word/document.xml", DOCUMENT)

    with pytest.raises(InvalidDocxError, match="Duplicate ZIP part"):
        OoxmlPackage.open(target.getvalue())


def test_package_enforces_source_file_size_limit() -> None:
    data = build_docx()
    limits = ResourceLimits(max_docx_file_size=len(data) - 1)

    with pytest.raises(ResourceLimitError, match="DOCX file size"):
        OoxmlPackage.open(data, limits=limits)


def test_package_enforces_expanded_size_limit() -> None:
    limits = ResourceLimits(max_expanded_size=16)

    with pytest.raises(ResourceLimitError, match="expanded size"):
        OoxmlPackage.open(build_docx(), limits=limits)


def test_package_enforces_xml_part_size_limit() -> None:
    limits = ResourceLimits(max_xml_part_size=16)

    with pytest.raises(ResourceLimitError, match="XML part"):
        OoxmlPackage.open(build_docx(), limits=limits)


def test_package_enforces_image_size_limit() -> None:
    limits = ResourceLimits(max_image_size=4)

    with pytest.raises(ResourceLimitError, match="image part"):
        OoxmlPackage.open(build_docx({"word/media/image1.png": b"12345"}), limits=limits)


def test_read_part_reports_missing_name() -> None:
    package = OoxmlPackage.open(build_docx())

    with pytest.raises(MissingPartError, match=r"word/styles\.xml"):
        package.read_part("word/styles.xml")


def test_relationship_graph_validation_does_not_depend_on_python_recursion_depth() -> None:
    relationship_sets: dict[str, RelationshipSet] = {}
    for index in range(1500):
        source = f"word/part{index}.xml"
        target = f"word/part{index + 1}.xml"
        relationship_sets[source] = RelationshipSet(
            source_part=source,
            relationships=(
                Relationship(
                    relationship_id=f"rId{index}",
                    relationship_type="urn:test",
                    target=target,
                    source_part=source,
                    resolved_target=target,
                ),
            ),
        )

    OoxmlPackage._validate_relationship_cycles(relationship_sets)
