"""Parsing for the required OPC ``[Content_Types].xml`` part."""

from __future__ import annotations

from pathlib import PurePosixPath

from pydantic import BaseModel, ConfigDict

from docxpdf_native.exceptions import InvalidOoxmlError
from docxpdf_native.ooxml.namespaces import OoxmlNamespaces
from docxpdf_native.ooxml.xml import SafeXmlParser


class ContentTypes(BaseModel):
    """Immutable default and override content-type mappings."""

    model_config = ConfigDict(frozen=True)

    defaults: dict[str, str]
    overrides: dict[str, str]

    def for_part(self, part_name: str) -> str | None:
        """Return the declared content type for a normalized package part."""

        override = self.overrides.get(part_name)
        if override is not None:
            return override
        extension = PurePosixPath(part_name).suffix.removeprefix(".").casefold()
        return self.defaults.get(extension)


class ContentTypesParser:
    """Validate and parse an OPC content-types manifest."""

    def __init__(self, xml_parser: SafeXmlParser) -> None:
        self._xml_parser = xml_parser

    def parse(self, data: bytes) -> ContentTypes:
        """Parse ``[Content_Types].xml`` into deterministic mappings."""

        part_name = "[Content_Types].xml"
        root = self._xml_parser.parse(data, part_name=part_name)
        if root.tag != OoxmlNamespaces.qn("ct", "Types"):
            raise InvalidOoxmlError(f"Content types part has unexpected root element {root.tag!r}")
        default_tag = OoxmlNamespaces.qn("ct", "Default")
        override_tag = OoxmlNamespaces.qn("ct", "Override")
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for child in root:
            if child.tag == default_tag:
                extension = child.get("Extension", "").lstrip(".").casefold()
                content_type = child.get("ContentType", "")
                if not extension or not content_type:
                    raise InvalidOoxmlError(
                        "Content type Default requires Extension and ContentType"
                    )
                if extension in defaults:
                    raise InvalidOoxmlError(
                        f"Duplicate content type default for extension {extension!r}"
                    )
                defaults[extension] = content_type
            elif child.tag == override_tag:
                raw_part = child.get("PartName", "")
                content_type = child.get("ContentType", "")
                part = raw_part.removeprefix("/")
                if (
                    not raw_part.startswith("/")
                    or not part
                    or "\\" in part
                    or ".." in PurePosixPath(part).parts
                    or not content_type
                ):
                    raise InvalidOoxmlError(f"Invalid content type Override PartName {raw_part!r}")
                if part in overrides:
                    raise InvalidOoxmlError(f"Duplicate content type override for part {part!r}")
                overrides[part] = content_type
            else:
                raise InvalidOoxmlError(f"Unexpected element {child.tag!r} in content types part")
        return ContentTypes(defaults=defaults, overrides=overrides)
