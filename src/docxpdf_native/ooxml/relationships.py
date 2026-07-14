"""Safe parsing and package-relative resolution of OOXML relationships."""

from __future__ import annotations

from pathlib import PurePosixPath
from urllib.parse import unquote, urlsplit
from xml.etree.ElementTree import Element

from pydantic import BaseModel, ConfigDict

from docxpdf_native.exceptions import RelationshipError
from docxpdf_native.ooxml.namespaces import OoxmlNamespaces
from docxpdf_native.ooxml.xml import SafeXmlParser


class Relationship(BaseModel):
    """A parsed relationship whose internal target has been normalized."""

    model_config = ConfigDict(frozen=True)

    relationship_id: str
    relationship_type: str
    target: str
    target_mode: str = "Internal"
    source_part: str | None = None
    resolved_target: str | None = None

    @property
    def is_external(self) -> bool:
        """Whether this relationship points outside the OPC package."""

        return self.target_mode.casefold() == "external"


class RelationshipSet(BaseModel):
    """Relationships owned by one package part (or by the package root)."""

    model_config = ConfigDict(frozen=True)

    source_part: str | None
    relationships: tuple[Relationship, ...] = ()

    def by_id(self, relationship_id: str) -> Relationship:
        """Return one relationship or report the precise owner and missing id."""

        for relation in self.relationships:
            if relation.relationship_id == relationship_id:
                return relation
        owner = self.source_part or "package root"
        raise RelationshipError(f"Relationship {relationship_id!r} does not exist for {owner!r}")


class RelationshipParser:
    """Parse relationship parts without performing any external I/O."""

    def __init__(self, xml_parser: SafeXmlParser, *, allow_external: bool = False) -> None:
        self._xml_parser = xml_parser
        self._allow_external = allow_external

    def parse(
        self,
        data: bytes,
        *,
        relationship_part: str,
        source_part: str | None,
    ) -> RelationshipSet:
        """Parse and normalize all relationships from a ``.rels`` part."""

        root = self._xml_parser.parse(data, part_name=relationship_part)
        expected_root = OoxmlNamespaces.qn("pr", "Relationships")
        expected_child = OoxmlNamespaces.qn("pr", "Relationship")
        if root.tag != expected_root:
            raise RelationshipError(
                f"Relationship part {relationship_part!r} has unexpected root {root.tag!r}"
            )

        seen: set[str] = set()
        relationships: list[Relationship] = []
        for element in root:
            if element.tag != expected_child:
                raise RelationshipError(
                    f"Unexpected element {element.tag!r} in relationship part {relationship_part!r}"
                )
            relation = self._parse_one(
                element,
                relationship_part=relationship_part,
                source_part=source_part,
            )
            if relation.relationship_id in seen:
                raise RelationshipError(
                    f"Duplicate relationship Id {relation.relationship_id!r} in "
                    f"{relationship_part!r}"
                )
            seen.add(relation.relationship_id)
            relationships.append(relation)
        return RelationshipSet(source_part=source_part, relationships=tuple(relationships))

    def _parse_one(
        self,
        element: Element,
        *,
        relationship_part: str,
        source_part: str | None,
    ) -> Relationship:
        relation_id = element.get("Id")
        relation_type = element.get("Type")
        target = element.get("Target")
        if not relation_id or not relation_type or not target:
            raise RelationshipError(
                f"Relationship in {relationship_part!r} must have Id, Type, and Target"
            )
        target_mode = element.get("TargetMode", "Internal")
        if target_mode.casefold() not in {"internal", "external"}:
            raise RelationshipError(
                f"Relationship {relation_id!r} in {relationship_part!r} has invalid "
                f"TargetMode {target_mode!r}"
            )
        if target_mode.casefold() == "external":
            if not self._allow_external:
                raise RelationshipError(
                    f"External relationship {relation_id!r} in {relationship_part!r} is disabled"
                )
            resolved_target = None
        else:
            resolved_target = self.resolve_internal_target(target, source_part=source_part)
        return Relationship(
            relationship_id=relation_id,
            relationship_type=relation_type,
            target=target,
            target_mode=target_mode.title(),
            source_part=source_part,
            resolved_target=resolved_target,
        )

    @staticmethod
    def resolve_internal_target(target: str, *, source_part: str | None) -> str:
        """Resolve a relationship target while preventing package-root escape."""

        decoded = unquote(target)
        split = urlsplit(decoded)
        if split.scheme or split.netloc:
            raise RelationshipError(
                f"Internal relationship target {target!r} contains an external URI"
            )
        if split.query or split.fragment:
            raise RelationshipError(
                f"Internal relationship target {target!r} contains a query or fragment"
            )
        path = split.path
        if not path or "\\" in path or "\x00" in path:
            raise RelationshipError(f"Invalid internal relationship target {target!r}")

        if path.startswith("/"):
            base_parts: list[str] = []
        elif source_part is None:
            base_parts = []
        else:
            base_parts = list(PurePosixPath(source_part).parent.parts)

        for component in path.split("/"):
            if component in {"", "."}:
                continue
            if component == "..":
                if not base_parts:
                    raise RelationshipError(f"Relationship target {target!r} escapes package root")
                base_parts.pop()
                continue
            if component.endswith(":"):
                raise RelationshipError(f"Invalid internal relationship target {target!r}")
            base_parts.append(component)

        if not base_parts:
            raise RelationshipError(f"Invalid empty relationship target {target!r}")
        return "/".join(base_parts)


def relationship_owner(relationship_part: str) -> str | None:
    """Map an OPC relationship-part name to the part that owns it."""

    if relationship_part == "_rels/.rels":
        return None
    if relationship_part.startswith("_rels/") and relationship_part.endswith(".rels"):
        filename = relationship_part.removeprefix("_rels/").removesuffix(".rels")
        if filename and "/" not in filename:
            return filename
    marker = "/_rels/"
    if marker not in relationship_part or not relationship_part.endswith(".rels"):
        raise RelationshipError(
            f"Relationship part has an invalid package location: {relationship_part!r}"
        )
    directory, filename = relationship_part.split(marker, maxsplit=1)
    if not directory or not filename or "/" in filename:
        raise RelationshipError(
            f"Relationship part has an invalid package location: {relationship_part!r}"
        )
    owner_name = filename.removesuffix(".rels")
    if not owner_name:
        raise RelationshipError(
            f"Relationship part has an invalid package location: {relationship_part!r}"
        )
    return f"{directory}/{owner_name}"
