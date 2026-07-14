"""Bounded, non-extracting reader for untrusted DOCX/OPC ZIP packages."""

from __future__ import annotations

import logging
import stat
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import BinaryIO, ClassVar
from zipfile import BadZipFile, LargeZipFile, ZipFile, ZipInfo

from pydantic import BaseModel, ConfigDict, Field

from docxpdf_native.exceptions import (
    InvalidDocxError,
    MissingPartError,
    RelationshipError,
)
from docxpdf_native.models.options import ResourceLimits
from docxpdf_native.ooxml.content_types import ContentTypes, ContentTypesParser
from docxpdf_native.ooxml.relationships import (
    RelationshipParser,
    RelationshipSet,
    relationship_owner,
)
from docxpdf_native.ooxml.xml import SafeXmlParser
from docxpdf_native.security.limits import ResourceLimitGuard

logger = logging.getLogger(__name__)


class OoxmlPackage(BaseModel):
    """An immutable in-memory view of a validated DOCX package.

    No archive member is ever extracted to disk.  This removes the write-side
    ZIP-slip risk and makes all resource accounting happen before OOXML parsing.
    """

    model_config = ConfigDict(frozen=True)

    REQUIRED_PARTS: ClassVar[tuple[str, ...]] = (
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
    )

    parts: dict[str, bytes] = Field(repr=False)
    content_types: ContentTypes
    relationship_sets: dict[str, RelationshipSet] = Field(repr=False)
    limits: ResourceLimits

    @classmethod
    def open(
        cls,
        source: Path | str | bytes | BinaryIO,
        *,
        limits: ResourceLimits | None = None,
        allow_external_relationships: bool = False,
    ) -> OoxmlPackage:
        """Read and validate a DOCX source without extracting archive members."""

        selected_limits = limits or ResourceLimits()
        archive_data = cls._read_source(source, maximum=selected_limits.max_docx_file_size)
        parts = cls._read_archive(archive_data, limits=selected_limits)
        cls._require_parts(parts)

        xml_parser = SafeXmlParser(
            max_part_size=selected_limits.max_xml_part_size,
            max_depth=selected_limits.max_xml_depth,
        )
        content_types = ContentTypesParser(xml_parser).parse(parts["[Content_Types].xml"])
        relationship_sets = cls._parse_relationships(
            parts,
            xml_parser=xml_parser,
            allow_external=allow_external_relationships,
        )
        cls._validate_relationship_targets(parts, relationship_sets)
        cls._validate_relationship_cycles(relationship_sets)
        cls._validate_root_document_relationship(relationship_sets)
        return cls(
            parts=parts,
            content_types=content_types,
            relationship_sets=relationship_sets,
            limits=selected_limits,
        )

    @property
    def part_names(self) -> tuple[str, ...]:
        """Return package part names in deterministic lexical order."""

        return tuple(sorted(self.parts))

    def has_part(self, part_name: str) -> bool:
        """Return whether a normalized part name exists."""

        return part_name in self.parts

    def read_part(self, part_name: str) -> bytes:
        """Return one part or raise a location-rich missing-part error."""

        try:
            return self.parts[part_name]
        except KeyError as error:
            raise MissingPartError(f"DOCX package is missing part {part_name!r}") from error

    def relationships_for(self, source_part: str | None) -> RelationshipSet:
        """Return relationships for a part, using an empty set when none exist."""

        key = source_part or ""
        relation_set = self.relationship_sets.get(key)
        if relation_set is not None:
            return relation_set
        return RelationshipSet(source_part=source_part)

    @staticmethod
    def _read_source(source: Path | str | bytes | BinaryIO, *, maximum: int) -> bytes:
        if isinstance(source, bytes):
            data = source
        elif isinstance(source, (Path, str)):
            path = Path(source)
            try:
                size = path.stat().st_size
            except OSError as error:
                raise InvalidDocxError(f"Cannot read DOCX path {path!s}: {error}") from error
            ResourceLimitGuard.ensure_at_most(
                actual=size,
                maximum=maximum,
                resource="DOCX file size",
            )
            try:
                with path.open("rb") as stream:
                    data = stream.read(maximum + 1)
            except OSError as error:
                raise InvalidDocxError(f"Cannot read DOCX path {path!s}: {error}") from error
        else:
            try:
                data = source.read(maximum + 1)
            except (OSError, ValueError) as error:
                raise InvalidDocxError(f"Cannot read DOCX binary stream: {error}") from error
            if not isinstance(data, bytes):
                raise InvalidDocxError("DOCX binary stream returned non-bytes data")
        ResourceLimitGuard.ensure_at_most(
            actual=len(data),
            maximum=maximum,
            resource="DOCX file size",
        )
        return data

    @classmethod
    def _read_archive(cls, data: bytes, *, limits: ResourceLimits) -> dict[str, bytes]:
        parts: dict[str, bytes] = {}
        try:
            with ZipFile(BytesIO(data), "r", allowZip64=True) as archive:
                infos = archive.infolist()
                cls._validate_infos(infos, limits=limits)
                for info in sorted(infos, key=lambda item: item.filename):
                    if info.is_dir():
                        continue
                    try:
                        content = archive.read(info)
                    except (BadZipFile, RuntimeError, NotImplementedError, OSError) as error:
                        raise InvalidDocxError(
                            f"Cannot read ZIP part {info.filename!r}: {error}"
                        ) from error
                    if len(content) != info.file_size:
                        raise InvalidDocxError(
                            f"ZIP part {info.filename!r} size differs from its directory entry"
                        )
                    parts[info.filename] = content
        except (BadZipFile, LargeZipFile, OSError) as error:
            raise InvalidDocxError(f"Invalid DOCX ZIP package: {error}") from error
        return parts

    @classmethod
    def _validate_infos(cls, infos: list[ZipInfo], *, limits: ResourceLimits) -> None:
        seen: set[str] = set()
        expanded = 0
        for info in infos:
            cls._validate_part_name(info.filename)
            if info.filename in seen:
                raise InvalidDocxError(f"Duplicate ZIP part {info.filename!r}")
            seen.add(info.filename)
            if info.flag_bits & 0x1:
                raise InvalidDocxError(f"Encrypted ZIP part is not supported: {info.filename!r}")
            unix_mode = info.external_attr >> 16
            if unix_mode and stat.S_ISLNK(unix_mode):
                raise InvalidDocxError(f"Symbolic-link ZIP part is not allowed: {info.filename!r}")
            expanded = ResourceLimitGuard.checked_add(
                current=expanded,
                increment=info.file_size,
                maximum=limits.max_expanded_size,
                resource="DOCX expanded size",
            )
            if not info.is_dir() and cls._is_xml_part(info.filename):
                ResourceLimitGuard.ensure_at_most(
                    actual=info.file_size,
                    maximum=limits.max_xml_part_size,
                    resource=f"XML part {info.filename!r} size",
                )
            if not info.is_dir() and cls._is_image_part(info.filename):
                ResourceLimitGuard.ensure_at_most(
                    actual=info.file_size,
                    maximum=limits.max_image_size,
                    resource=f"image part {info.filename!r} size",
                )

    @staticmethod
    def _validate_part_name(name: str) -> None:
        path = PurePosixPath(name)
        invalid = (
            not name
            or name.startswith("/")
            or "\\" in name
            or "\x00" in name
            or path.is_absolute()
            or any(component in {"", ".", ".."} for component in path.parts)
            or any(":" in component for component in path.parts)
        )
        if invalid:
            raise InvalidDocxError(f"Unsafe ZIP part name {name!r}")

    @staticmethod
    def _is_xml_part(name: str) -> bool:
        lowered = name.casefold()
        return lowered.endswith((".xml", ".rels"))

    @staticmethod
    def _is_image_part(name: str) -> bool:
        normalized = f"/{name.casefold()}"
        return "/media/" in normalized

    @classmethod
    def _require_parts(cls, parts: dict[str, bytes]) -> None:
        for required in cls.REQUIRED_PARTS:
            if required not in parts:
                raise MissingPartError(f"DOCX package is missing required part {required!r}")

    @staticmethod
    def _parse_relationships(
        parts: dict[str, bytes],
        *,
        xml_parser: SafeXmlParser,
        allow_external: bool,
    ) -> dict[str, RelationshipSet]:
        parser = RelationshipParser(xml_parser, allow_external=allow_external)
        parsed: dict[str, RelationshipSet] = {}
        for part_name in sorted(parts):
            if not part_name.endswith(".rels"):
                continue
            source = relationship_owner(part_name)
            if source is not None and source not in parts:
                raise RelationshipError(
                    f"Relationship part {part_name!r} belongs to missing part {source!r}"
                )
            key = source or ""
            parsed[key] = parser.parse(
                parts[part_name],
                relationship_part=part_name,
                source_part=source,
            )
        return parsed

    @staticmethod
    def _validate_relationship_targets(
        parts: dict[str, bytes],
        relationship_sets: dict[str, RelationshipSet],
    ) -> None:
        for relation_set in relationship_sets.values():
            for relation in relation_set.relationships:
                target = relation.resolved_target
                if target is not None and target not in parts:
                    owner = relation.source_part or "package root"
                    raise RelationshipError(
                        f"Relationship {relation.relationship_id!r} from {owner!r} targets "
                        f"missing part {target!r}"
                    )

    @staticmethod
    def _validate_relationship_cycles(relationship_sets: dict[str, RelationshipSet]) -> None:
        graph: dict[str, tuple[str, ...]] = {}
        for key, relation_set in relationship_sets.items():
            targets = tuple(
                relation.resolved_target
                for relation in relation_set.relationships
                if relation.resolved_target is not None
            )
            graph[key] = targets

        # 0 = unseen, 1 = on the active path, 2 = completely visited.  An
        # explicit stack avoids making attacker-controlled relationship depth
        # consume the Python call stack.
        states: dict[str, int] = {}
        for start in sorted(graph):
            if states.get(start, 0) != 0:
                continue
            states[start] = 1
            path = [start]
            positions = {start: 0}
            stack: list[tuple[str, int]] = [(start, 0)]
            while stack:
                node, target_index = stack[-1]
                targets = graph[node]
                if target_index >= len(targets):
                    stack.pop()
                    path.pop()
                    positions.pop(node)
                    states[node] = 2
                    continue
                target = targets[target_index]
                stack[-1] = (node, target_index + 1)
                if target not in graph:
                    continue
                state = states.get(target, 0)
                if state == 1:
                    cycle_start = positions[target]
                    cycle = " -> ".join((*path[cycle_start:], target))
                    raise RelationshipError(f"Relationship cycle detected: {cycle}")
                if state == 2:
                    continue
                states[target] = 1
                positions[target] = len(path)
                path.append(target)
                stack.append((target, 0))

    @staticmethod
    def _validate_root_document_relationship(
        relationship_sets: dict[str, RelationshipSet],
    ) -> None:
        root = relationship_sets.get("")
        if root is None:
            raise RelationshipError("Package root relationship part is missing")
        office_document_suffix = "/officeDocument"
        if not any(
            relation.relationship_type.endswith(office_document_suffix)
            and relation.resolved_target == "word/document.xml"
            for relation in root.relationships
        ):
            raise RelationshipError(
                "Package root has no officeDocument relationship to 'word/document.xml'"
            )
