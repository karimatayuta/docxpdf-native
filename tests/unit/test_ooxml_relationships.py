from __future__ import annotations

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from docxpdf_native.exceptions import RelationshipError
from docxpdf_native.ooxml.package import OoxmlPackage
from docxpdf_native.ooxml.relationships import RelationshipParser, relationship_owner
from docxpdf_native.ooxml.xml import SafeXmlParser

CONTENT_TYPES = b"""<Types
 xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
</Types>"""
DOCUMENT = b"""<w:document
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"""
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def relationship_xml(*relationships: str) -> bytes:
    body = "".join(relationships)
    return f'<Relationships xmlns="{PACKAGE_REL_NS}">{body}</Relationships>'.encode()


def relationship(
    relationship_id: str,
    target: str,
    *,
    relationship_type: str = "urn:test",
    target_mode: str | None = None,
) -> str:
    mode = "" if target_mode is None else f' TargetMode="{target_mode}"'
    return (
        f'<Relationship Id="{relationship_id}" Type="{relationship_type}" Target="{target}"{mode}/>'
    )


def build_docx(parts: dict[str, bytes]) -> bytes:
    defaults = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": relationship_xml(
            relationship(
                "rId1",
                "word/document.xml",
                relationship_type=(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                    "officeDocument"
                ),
            )
        ),
        "word/document.xml": DOCUMENT,
    }
    defaults.update(parts)
    target = BytesIO()
    with ZipFile(target, "w", compression=ZIP_DEFLATED) as archive:
        for name, value in defaults.items():
            archive.writestr(name, value)
    return target.getvalue()


def test_relationships_resolve_relative_target_from_owner_part() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "media/image.png")
            ),
            "word/media/image.png": b"PNG",
        }
    )

    package = OoxmlPackage.open(data)
    relation = package.relationships_for("word/document.xml").by_id("rId5")

    assert relation.resolved_target == "word/media/image.png"


def test_relationships_resolve_parent_segments_without_leaving_package() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "../customXml/item.xml")
            ),
            "customXml/item.xml": b"<item />",
        }
    )

    relation = OoxmlPackage.open(data).relationships_for("word/document.xml").by_id("rId5")

    assert relation.resolved_target == "customXml/item.xml"


def test_relationships_reject_external_target_by_default() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "https://example.invalid/image.png", target_mode="External")
            )
        }
    )

    with pytest.raises(RelationshipError, match="External relationship"):
        OoxmlPackage.open(data)


def test_relationships_can_retain_external_target_without_fetching_it() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "https://example.invalid/", target_mode="External")
            )
        }
    )

    relation = (
        OoxmlPackage.open(data, allow_external_relationships=True)
        .relationships_for("word/document.xml")
        .by_id("rId5")
    )

    assert relation.resolved_target is None


def test_relationships_reject_target_that_escapes_package_root() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "../../outside.xml")
            )
        }
    )

    with pytest.raises(RelationshipError, match="escapes package root"):
        OoxmlPackage.open(data)


def test_relationships_reject_missing_internal_target() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "media/missing.png")
            )
        }
    )

    with pytest.raises(RelationshipError, match="missing part"):
        OoxmlPackage.open(data)


def test_relationships_reject_duplicate_ids() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(
                relationship("rId5", "media/a.png"),
                relationship("rId5", "media/b.png"),
            ),
            "word/media/a.png": b"A",
            "word/media/b.png": b"B",
        }
    )

    with pytest.raises(RelationshipError, match="Duplicate relationship Id"):
        OoxmlPackage.open(data)


def test_relationships_reject_cycle() -> None:
    data = build_docx(
        {
            "word/_rels/document.xml.rels": relationship_xml(relationship("rId2", "header1.xml")),
            "word/header1.xml": b"<header />",
            "word/_rels/header1.xml.rels": relationship_xml(relationship("rId3", "document.xml")),
        }
    )

    with pytest.raises(RelationshipError, match="cycle"):
        OoxmlPackage.open(data)


def test_relationship_lookup_reports_unknown_id() -> None:
    package = OoxmlPackage.open(build_docx({}))

    with pytest.raises(RelationshipError, match="rId404"):
        package.relationships_for(None).by_id("rId404")


def test_root_level_part_relationship_location_is_supported() -> None:
    assert relationship_owner("_rels/custom.xml.rels") == "custom.xml"


@pytest.mark.parametrize(
    "xml",
    [
        b"<root />",
        f'<Relationships xmlns="{PACKAGE_REL_NS}"><Unexpected/></Relationships>'.encode(),
        (
            f'<Relationships xmlns="{PACKAGE_REL_NS}">'
            '<Relationship Id="rId1" Type="urn:test"/></Relationships>'
        ).encode(),
        (
            f'<Relationships xmlns="{PACKAGE_REL_NS}">'
            '<Relationship Id="rId1" Type="urn:test" Target="part.xml" '
            'TargetMode="Unknown"/></Relationships>'
        ).encode(),
    ],
)
def test_relationship_parser_rejects_invalid_manifest(xml: bytes) -> None:
    parser = RelationshipParser(SafeXmlParser(max_part_size=4096, max_depth=16))

    with pytest.raises(RelationshipError):
        parser.parse(
            xml, relationship_part="word/_rels/document.xml.rels", source_part="word/document.xml"
        )


@pytest.mark.parametrize(
    "target",
    [
        "https://example.invalid/part.xml",
        "part.xml?query=1",
        "part.xml#fragment",
        "",
        "media\\image.png",
        "folder:/part.xml",
    ],
)
def test_internal_relationship_target_rejects_non_part_uri(target: str) -> None:
    with pytest.raises(RelationshipError):
        RelationshipParser.resolve_internal_target(target, source_part="word/document.xml")


@pytest.mark.parametrize("part_name", ["word/document.xml.rels", "word/_rels/nested/a.rels"])
def test_relationship_owner_rejects_invalid_location(part_name: str) -> None:
    with pytest.raises(RelationshipError, match="invalid package location"):
        relationship_owner(part_name)
