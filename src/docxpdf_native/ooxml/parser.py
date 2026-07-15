"""OOXML-to-model parser for the documented native rendering subset."""

from __future__ import annotations

import logging
import math
import re
import unicodedata
from collections.abc import Iterator
from decimal import Decimal
from pathlib import Path
from typing import BinaryIO, Literal
from xml.etree.ElementTree import Element

from pydantic import BaseModel, ConfigDict, ValidationError

from docxpdf_native.abstractions import DocumentParser, UnsupportedFeatureHandler
from docxpdf_native.exceptions import InvalidOoxmlError, RelationshipError, ResourceLimitError
from docxpdf_native.models.document import (
    BorderModel,
    CellMargins,
    DocumentMetadata,
    DocumentModel,
    HeaderFooterModel,
    ImageModel,
    NumberingDefinition,
    NumberingLevel,
    ParagraphModel,
    PlaceholderModel,
    RunModel,
    SectionModel,
    TableBorders,
    TableCellModel,
    TableModel,
    TableRowModel,
)
from docxpdf_native.models.options import ConversionOptions, ResourceLimits
from docxpdf_native.models.results import ConversionWarning, UnsupportedFeature
from docxpdf_native.models.styles import ParagraphProperties, RunProperties, TabStop, ThemeFonts
from docxpdf_native.ooxml.namespaces import OoxmlNamespaces
from docxpdf_native.ooxml.package import OoxmlPackage
from docxpdf_native.ooxml.raster import RasterImageConverter
from docxpdf_native.ooxml.styles import OoxmlStylesParser, ParsedStyleSheet, ThemeFontParser
from docxpdf_native.ooxml.unsupported import (
    LenientUnsupportedFeatureHandler,
    StrictUnsupportedFeatureHandler,
    UnsupportedFeatureDetector,
)
from docxpdf_native.ooxml.xml import SafeXmlParser

logger = logging.getLogger(__name__)

DocumentSource = Path | str | bytes | BinaryIO
BlockModel = ParagraphModel | TableModel
Visual = ImageModel | PlaceholderModel

# VML (``style="width:...;height:...")`` and legacy ``w:dxaOrig``/``w:dyaOrig``
# lengths are expressed in a handful of CSS-like units; VML's own default unit
# (a bare number) is points.
_VML_UNITS_TO_POINTS: dict[str, float] = {
    "pt": 1.0,
    "in": 72.0,
    "cm": 72.0 / 2.54,
    "mm": 72.0 / 25.4,
    "px": 0.75,
    "pc": 12.0,
}
_VML_LENGTH_PATTERN = re.compile(r"^\s*(-?[0-9]*\.?[0-9]+)\s*([a-zA-Z%]*)\s*$")


class _FieldFrame:
    """Mutable state for one active ``w:fldChar`` begin/separate/end span."""

    __slots__ = ("instruction", "phase", "result_runs")

    def __init__(self) -> None:
        self.phase: Literal["instruction", "result"] = "instruction"
        self.instruction: str = ""
        self.result_runs: list[RunModel] = []


class _FieldTracker:
    """Track nested complex fields within one paragraph.

    Only the outermost field's cached result is ever emitted.  Anything
    nested (typically the field's own instruction operands, e.g. the ``PAGE``
    field inside ``IF { PAGE } = 1 ...``) is discarded, matching how Word
    itself only bakes the outermost cached result into the visible run
    stream.
    """

    def __init__(self) -> None:
        self._stack: list[_FieldFrame] = []

    @property
    def depth(self) -> int:
        return len(self._stack)

    def begin(self) -> None:
        self._stack.append(_FieldFrame())

    def separate(self) -> None:
        if self._stack:
            self._stack[-1].phase = "result"

    def record_instruction(self, text: str) -> None:
        if self._stack and self._stack[-1].phase == "instruction":
            self._stack[-1].instruction += text

    def should_buffer_result(self) -> bool:
        return len(self._stack) == 1 and self._stack[-1].phase == "result"

    def buffer(self, run: RunModel) -> None:
        self._stack[-1].result_runs.append(run)

    def end(self) -> list[RunModel]:
        if not self._stack:
            return []
        frame = self._stack.pop()
        if self._stack:
            return []
        return self._finalize(frame)

    def flush_incomplete(self) -> list[RunModel]:
        """Force-close any field left open at the end of a paragraph.

        Complex fields are tracked per paragraph; one that never receives a
        matching ``end`` (rare, e.g. a field spanning multiple paragraphs)
        still surfaces whatever result text it had captured so far, rather
        than silently losing it.
        """

        if not self._stack:
            return []
        while len(self._stack) > 1:
            self._stack.pop()
        frame = self._stack.pop()
        if frame.phase != "result":
            return []
        return self._finalize(frame)

    @staticmethod
    def _finalize(frame: _FieldFrame) -> list[RunModel]:
        instruction = frame.instruction.strip()
        command = instruction.split(maxsplit=1)[0].upper() if instruction else ""
        sentinel = {"PAGE": "{{PAGE}}", "NUMPAGES": "{{NUMPAGES}}"}.get(command)
        if sentinel is None:
            return frame.result_runs
        if frame.result_runs:
            first = frame.result_runs[0]
            return [
                RunModel(
                    text=sentinel,
                    style_id=first.style_id,
                    properties=first.properties,
                    preserve_space=first.preserve_space,
                    hidden=first.hidden,
                    source_index=first.source_index,
                )
            ]
        return [RunModel(text=sentinel)]


class _ParseState(BaseModel):
    model_config = ConfigDict(validate_assignment=True)

    limits: ResourceLimits
    paragraphs: int = 0
    runs: int = 0
    table_cells: int = 0
    blocks: int = 0

    def count_paragraph(self, *, part_name: str) -> int:
        index = self.paragraphs
        self.paragraphs += 1
        if self.paragraphs > self.limits.max_paragraphs:
            raise ResourceLimitError(
                f"paragraph count exceeds limit {self.limits.max_paragraphs}",
                part=part_name,
            )
        return index

    def count_run(self, *, part_name: str) -> int:
        index = self.runs
        self.runs += 1
        if self.runs > self.limits.max_runs:
            raise ResourceLimitError(
                f"run count exceeds limit {self.limits.max_runs}",
                part=part_name,
            )
        return index

    def count_cell(self, *, part_name: str) -> int:
        index = self.table_cells
        self.table_cells += 1
        if self.table_cells > self.limits.max_table_cells:
            raise ResourceLimitError(
                f"table cell count exceeds limit {self.limits.max_table_cells}",
                part=part_name,
            )
        return index

    def next_block(self) -> int:
        index = self.blocks
        self.blocks += 1
        return index


class OoxmlDocumentParser(DocumentParser):
    """Parse an OPC package into PDF-independent Pydantic document models."""

    def __init__(
        self,
        *,
        limits: ResourceLimits | None = None,
        unsupported_handler: UnsupportedFeatureHandler | None = None,
    ) -> None:
        self._limits = limits
        self._unsupported_handler = unsupported_handler
        self._warnings: tuple[ConversionWarning, ...] = ()
        self._unsupported_features: tuple[UnsupportedFeature, ...] = ()
        self._placeholder_features: list[UnsupportedFeature] = []
        self._placeholder_warnings: list[ConversionWarning] = []

    @property
    def warnings(self) -> tuple[ConversionWarning, ...]:
        """Warnings produced by the most recent parse call."""

        return self._warnings

    @property
    def unsupported_features(self) -> tuple[UnsupportedFeature, ...]:
        """Unsupported features found by the most recent parse call."""

        return self._unsupported_features

    def parse(
        self,
        source: DocumentSource,
        *,
        options: ConversionOptions | None = None,
    ) -> DocumentModel:
        """Parse a DOCX path, byte string, or binary stream.

        The wider source type is intentional: it is compatible with the ABC's
        byte contract and is also the concrete parser used by the public API.
        """

        selected_options = options or ConversionOptions()
        self._placeholder_features = []
        self._placeholder_warnings = []
        limits = self._limits or selected_options.resource_limits
        package = OoxmlPackage.open(
            source,
            limits=limits,
            # External relationships (hyperlinks, external OLE targets, ...)
            # are common in real-world DOCX files and are always tolerated at
            # the package level: their target bytes are never fetched, so
            # merely referencing an external target is not a security or
            # correctness concern in itself.
            allow_external_relationships=True,
        )
        state = _ParseState(limits=limits)
        xml_parser = SafeXmlParser(
            max_part_size=limits.max_xml_part_size,
            max_depth=limits.max_xml_depth,
        )
        root = xml_parser.parse(
            package.read_part("word/document.xml"),
            part_name="word/document.xml",
        )
        if root.tag != OoxmlNamespaces.qn("w", "document"):
            raise InvalidOoxmlError(
                "Main document has an unexpected root element",
                part="word/document.xml",
                location="/w:document",
            )
        handler = self._unsupported_handler
        if handler is None:
            handler = (
                StrictUnsupportedFeatureHandler()
                if selected_options.strict
                else LenientUnsupportedFeatureHandler()
            )
        detector = UnsupportedFeatureDetector(handler)
        detector.scan_package(package)
        detector.scan_element(root, part_name="word/document.xml")
        optional_roots: dict[str, Element] = {}
        for part_name in package.part_names:
            if part_name == "word/document.xml" or not part_name.casefold().endswith(".xml"):
                continue
            optional_root = xml_parser.parse(package.read_part(part_name), part_name=part_name)
            optional_roots[part_name] = optional_root
            # VML elements under settings/shapeDefaults are drawing defaults,
            # not rendered document shapes. Content-bearing parts are still scanned.
            if part_name != "word/settings.xml":
                detector.scan_element(optional_root, part_name=part_name)
        self._validate_optional_part_roots(optional_roots)
        body = root.find(OoxmlNamespaces.qn("w", "body"))
        if body is None:
            raise InvalidOoxmlError(
                "Main document has no w:body",
                part="word/document.xml",
                location="/w:document",
            )
        try:
            sections = self._parse_sections(
                body,
                package=package,
                xml_parser=xml_parser,
                state=state,
            )
            style_sheet = self._parse_style_sheet(optional_roots.get("word/styles.xml"))
            theme_fonts = self._parse_theme_fonts(optional_roots.get("word/theme/theme1.xml"))
            numbering = self._parse_numbering(optional_roots.get("word/numbering.xml"))
            metadata = self._parse_metadata(
                optional_roots.get("docProps/core.xml"),
                optional_roots.get("docProps/app.xml"),
            )
            source_name = str(source) if isinstance(source, (Path, str)) else None
            handler_warnings = (
                handler.warnings if isinstance(handler, LenientUnsupportedFeatureHandler) else ()
            )
            handler_features = (
                handler.features if isinstance(handler, LenientUnsupportedFeatureHandler) else ()
            )
            # Placeholder diagnostics are collected while parsing sections above
            # and always surface, in both strict and lenient mode: a same-size
            # placeholder is a reported degradation, not a silent one, so it
            # never blocks strict mode the way a truly unsupported feature does.
            self._warnings = (*handler_warnings, *self._placeholder_warnings)
            self._unsupported_features = (*handler_features, *self._placeholder_features)
            return DocumentModel(
                sections=sections,
                styles=style_sheet.styles,
                defaults=style_sheet.defaults,
                theme_fonts=theme_fonts,
                numbering=numbering,
                metadata=metadata,
                source_name=source_name,
            )
        except ValidationError as error:
            raise InvalidOoxmlError(
                f"OOXML value failed model validation: {error.errors(include_url=False)}",
                part="word/document.xml",
            ) from error

    @staticmethod
    def _validate_optional_part_roots(roots: dict[str, Element]) -> None:
        exact_roots = {
            "word/settings.xml": OoxmlNamespaces.qn("w", "settings"),
            "word/fontTable.xml": OoxmlNamespaces.qn("w", "fonts"),
            "docProps/core.xml": OoxmlNamespaces.qn("cp", "coreProperties"),
            "docProps/app.xml": OoxmlNamespaces.qn("ep", "Properties"),
        }
        for part_name, root in roots.items():
            expected = exact_roots.get(part_name)
            if part_name.startswith("word/header") and part_name.endswith(".xml"):
                expected = OoxmlNamespaces.qn("w", "hdr")
            elif part_name.startswith("word/footer") and part_name.endswith(".xml"):
                expected = OoxmlNamespaces.qn("w", "ftr")
            if expected is not None and root.tag != expected:
                raise InvalidOoxmlError(
                    f"OOXML part {part_name!r} has unexpected root element {root.tag!r}",
                    part=part_name,
                    location="/",
                )

    @staticmethod
    def _parse_style_sheet(root: Element | None) -> ParsedStyleSheet:
        if root is None:
            return ParsedStyleSheet()
        if root.tag != OoxmlNamespaces.qn("w", "styles"):
            raise InvalidOoxmlError(
                "Styles part has an unexpected root element",
                part="word/styles.xml",
            )
        try:
            return OoxmlStylesParser.parse(root)
        except (TypeError, ValueError) as error:
            raise InvalidOoxmlError(
                f"Invalid styles part: {error}",
                part="word/styles.xml",
            ) from error

    @staticmethod
    def _parse_theme_fonts(root: Element | None) -> ThemeFonts:
        if root is None:
            return ThemeFonts()
        if root.tag != OoxmlNamespaces.qn("a", "theme"):
            raise InvalidOoxmlError(
                "Theme part has an unexpected root element",
                part="word/theme/theme1.xml",
            )
        return ThemeFontParser.parse(root)

    def _parse_numbering(self, root: Element | None) -> tuple[NumberingDefinition, ...]:
        if root is None:
            return ()
        if root.tag != OoxmlNamespaces.qn("w", "numbering"):
            raise InvalidOoxmlError(
                "Numbering part has an unexpected root element",
                part="word/numbering.xml",
            )
        abstract_levels: dict[int, tuple[NumberingLevel, ...]] = {}
        for abstract in root.findall(OoxmlNamespaces.qn("w", "abstractNum")):
            abstract_id = self._required_int(abstract, "w", "abstractNumId")
            levels: list[NumberingLevel] = []
            for level in abstract.findall(OoxmlNamespaces.qn("w", "lvl")):
                level_index = self._required_int(level, "w", "ilvl")
                start_element = level.find(OoxmlNamespaces.qn("w", "start"))
                format_element = level.find(OoxmlNamespaces.qn("w", "numFmt"))
                text_element = level.find(OoxmlNamespaces.qn("w", "lvlText"))
                start = (
                    self._required_int(start_element, "w", "val")
                    if start_element is not None
                    else 1
                )
                number_format = (
                    format_element.get(OoxmlNamespaces.qn("w", "val"), "decimal")
                    if format_element is not None
                    else "decimal"
                )
                level_text = (
                    text_element.get(OoxmlNamespaces.qn("w", "val"), f"%{level_index + 1}.")
                    if text_element is not None
                    else f"%{level_index + 1}."
                )
                paragraph_properties, _ = self._parse_paragraph_properties(
                    level.find(OoxmlNamespaces.qn("w", "pPr"))
                )
                run_properties, _ = self._parse_run_properties(
                    level.find(OoxmlNamespaces.qn("w", "rPr"))
                )
                levels.append(
                    NumberingLevel(
                        level=level_index,
                        number_format=number_format,
                        text=level_text,
                        start=max(0, start),
                        paragraph=paragraph_properties,
                        run=run_properties,
                    )
                )
            abstract_levels[abstract_id] = tuple(levels)

        definitions: list[NumberingDefinition] = []
        seen: set[int] = set()
        for numbering in root.findall(OoxmlNamespaces.qn("w", "num")):
            numbering_id = self._required_int(numbering, "w", "numId")
            abstract_reference = numbering.find(OoxmlNamespaces.qn("w", "abstractNumId"))
            referenced_abstract_id = (
                self._required_int(abstract_reference, "w", "val")
                if abstract_reference is not None
                else None
            )
            if numbering_id in seen:
                raise InvalidOoxmlError(
                    f"Duplicate numbering id {numbering_id}",
                    part="word/numbering.xml",
                )
            seen.add(numbering_id)
            definitions.append(
                NumberingDefinition(
                    numbering_id=numbering_id,
                    abstract_numbering_id=referenced_abstract_id,
                    levels=(
                        abstract_levels.get(referenced_abstract_id, ())
                        if referenced_abstract_id is not None
                        else ()
                    ),
                )
            )
        return tuple(definitions)

    @staticmethod
    def _parse_metadata(
        core_root: Element | None,
        app_root: Element | None,
    ) -> DocumentMetadata:
        values: dict[str, str] = {}
        core_mapping = {
            "title": "title",
            "subject": "subject",
            "creator": "creator",
            "keywords": "keywords",
            "description": "description",
        }
        if core_root is not None:
            for child in core_root:
                local_name = child.tag.rsplit("}", maxsplit=1)[-1]
                field = core_mapping.get(local_name)
                if field is not None and child.text:
                    values[field] = child.text
        if app_root is not None:
            for child in app_root:
                if child.tag.rsplit("}", maxsplit=1)[-1] == "Application" and child.text:
                    values["application"] = child.text
                    break
        return DocumentMetadata.model_validate(values)

    def _parse_sections(
        self,
        body: Element,
        *,
        package: OoxmlPackage,
        xml_parser: SafeXmlParser,
        state: _ParseState,
    ) -> tuple[SectionModel, ...]:
        sections: list[SectionModel] = []
        current: list[BlockModel] = []
        final_properties: Element | None = None
        paragraph_tag = OoxmlNamespaces.qn("w", "p")
        table_tag = OoxmlNamespaces.qn("w", "tbl")
        section_tag = OoxmlNamespaces.qn("w", "sectPr")
        paragraph_properties_tag = OoxmlNamespaces.qn("w", "pPr")

        for child in self._flatten_block_children(body):
            if child.tag == paragraph_tag:
                paragraph = self._parse_paragraph(
                    child,
                    package=package,
                    part_name="word/document.xml",
                    state=state,
                )
                current.append(paragraph)
                properties = child.find(paragraph_properties_tag)
                section_properties = (
                    properties.find(section_tag) if properties is not None else None
                )
                if section_properties is not None:
                    sections.append(
                        self._make_section(
                            current,
                            section_properties,
                            package=package,
                            xml_parser=xml_parser,
                            state=state,
                            source_index=len(sections),
                        )
                    )
                    current = []
            elif child.tag == table_tag:
                current.append(
                    self._parse_table(
                        child,
                        package=package,
                        part_name="word/document.xml",
                        state=state,
                    )
                )
            elif child.tag == section_tag:
                final_properties = child

        if current or not sections or final_properties is not None:
            sections.append(
                self._make_section(
                    current,
                    final_properties,
                    package=package,
                    xml_parser=xml_parser,
                    state=state,
                    source_index=len(sections),
                )
            )
        return tuple(sections)

    def _make_section(
        self,
        blocks: list[BlockModel],
        properties: Element | None,
        *,
        package: OoxmlPackage,
        xml_parser: SafeXmlParser,
        state: _ParseState,
        source_index: int,
    ) -> SectionModel:
        values: dict[str, object] = {
            "blocks": tuple(blocks),
            "source_index": source_index,
        }
        if properties is not None:
            page_size = properties.find(OoxmlNamespaces.qn("w", "pgSz"))
            if page_size is not None:
                width = self._optional_int(page_size, "w", "w")
                height = self._optional_int(page_size, "w", "h")
                if width is not None:
                    values["page_width"] = self.twips_to_points(width)
                if height is not None:
                    values["page_height"] = self.twips_to_points(height)
                orientation = page_size.get(OoxmlNamespaces.qn("w", "orient"), "portrait")
                if orientation not in {"portrait", "landscape"}:
                    raise InvalidOoxmlError(f"Invalid page orientation {orientation!r}")
                values["orientation"] = orientation

            margins = properties.find(OoxmlNamespaces.qn("w", "pgMar"))
            if margins is not None:
                for attribute, field in (
                    ("top", "margin_top"),
                    ("right", "margin_right"),
                    ("bottom", "margin_bottom"),
                    ("left", "margin_left"),
                    ("header", "header_distance"),
                    ("footer", "footer_distance"),
                ):
                    value = self._optional_int(margins, "w", attribute)
                    if value is not None:
                        values[field] = max(0.0, self.twips_to_points(value))

            number_type = properties.find(OoxmlNamespaces.qn("w", "pgNumType"))
            if number_type is not None:
                start = self._optional_int(number_type, "w", "start")
                if start is not None:
                    values["page_number_start"] = max(0, start)

            break_element = properties.find(OoxmlNamespaces.qn("w", "type"))
            if break_element is not None:
                break_value = break_element.get(OoxmlNamespaces.qn("w", "val"), "nextPage")
                break_mapping = {
                    "continuous": "continuous",
                    "nextPage": "next_page",
                    "evenPage": "even_page",
                    "oddPage": "odd_page",
                }
                mapped = break_mapping.get(break_value)
                if mapped is None:
                    raise InvalidOoxmlError(f"Invalid section break type {break_value!r}")
                values["section_break"] = mapped

            columns = properties.find(OoxmlNamespaces.qn("w", "cols"))
            if columns is not None:
                column_count = self._optional_int(columns, "w", "num") or 1
                values["columns"] = max(1, column_count)

            doc_grid = properties.find(OoxmlNamespaces.qn("w", "docGrid"))
            if doc_grid is not None:
                grid_type = doc_grid.get(OoxmlNamespaces.qn("w", "type"), "default")
                if grid_type not in {"default", "lines", "linesAndChars", "snapToChars"}:
                    raise InvalidOoxmlError(f"Invalid docGrid type {grid_type!r}")
                values["doc_grid_type"] = grid_type
                line_pitch = self._optional_int(doc_grid, "w", "linePitch")
                if line_pitch is not None and line_pitch > 0:
                    values["doc_grid_line_pitch"] = self.twips_to_points(line_pitch)

            values["headers"] = self._parse_header_footer_references(
                properties,
                package=package,
                xml_parser=xml_parser,
                state=state,
                reference_name="headerReference",
            )
            values["footers"] = self._parse_header_footer_references(
                properties,
                package=package,
                xml_parser=xml_parser,
                state=state,
                reference_name="footerReference",
            )
        return SectionModel.model_validate(values)

    def _parse_header_footer_references(
        self,
        section_properties: Element,
        *,
        package: OoxmlPackage,
        xml_parser: SafeXmlParser,
        state: _ParseState,
        reference_name: str,
    ) -> tuple[HeaderFooterModel, ...]:
        references: list[HeaderFooterModel] = []
        relationship_set = package.relationships_for("word/document.xml")
        for reference in section_properties.findall(OoxmlNamespaces.qn("w", reference_name)):
            relation_id = reference.get(OoxmlNamespaces.qn("r", "id"))
            if not relation_id:
                raise RelationshipError(
                    f"{reference_name} is missing r:id",
                    part="word/document.xml",
                )
            relation = relationship_set.by_id(relation_id)
            if relation.resolved_target is None:
                raise RelationshipError(
                    f"{reference_name} {relation_id!r} cannot be external",
                    part="word/document.xml",
                )
            part_name = relation.resolved_target
            root = xml_parser.parse(package.read_part(part_name), part_name=part_name)
            blocks = self._parse_story_blocks(
                root,
                package=package,
                part_name=part_name,
                state=state,
            )
            kind = reference.get(OoxmlNamespaces.qn("w", "type"), "default")
            if kind not in {"default", "first", "even"}:
                raise InvalidOoxmlError(
                    f"Invalid {reference_name} type {kind!r}",
                    part="word/document.xml",
                )
            references.append(
                HeaderFooterModel(
                    kind=kind,
                    relationship_id=relation_id,
                    part_name=part_name,
                    blocks=blocks,
                )
            )
        return tuple(references)

    def _parse_story_blocks(
        self,
        root: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
    ) -> tuple[BlockModel, ...]:
        blocks: list[BlockModel] = []
        for child in self._flatten_block_children(root):
            if child.tag == OoxmlNamespaces.qn("w", "p"):
                blocks.append(
                    self._parse_paragraph(
                        child,
                        package=package,
                        part_name=part_name,
                        state=state,
                    )
                )
            elif child.tag == OoxmlNamespaces.qn("w", "tbl"):
                blocks.append(
                    self._parse_table(
                        child,
                        package=package,
                        part_name=part_name,
                        state=state,
                    )
                )
        return tuple(blocks)

    # -- Block-level container flattening -----------------------------------
    #
    # ``w:sdt`` (content controls) and ``mc:AlternateContent`` can wrap
    # ordinary block content (``w:p``/``w:tbl``, possibly nested further).
    # Flattening them here means every caller that walks block children only
    # ever sees the real content, never the wrapper, without special-casing
    # every call site.

    def _flatten_block_children(self, parent: Element) -> Iterator[Element]:
        for child in parent:
            yield from self._flatten_block_element(child)

    def _flatten_block_element(self, element: Element) -> Iterator[Element]:
        if element.tag == OoxmlNamespaces.qn("w", "sdt"):
            content = element.find(OoxmlNamespaces.qn("w", "sdtContent"))
            if content is not None:
                yield from self._flatten_block_children(content)
            return
        if element.tag == OoxmlNamespaces.qn("mc", "AlternateContent"):
            yield from self._flatten_alternate_content(element)
            return
        yield element

    def _flatten_alternate_content(self, element: Element) -> Iterator[Element]:
        # Markup-compatibility fallback: prefer the modern mc:Choice branches
        # we understand, otherwise render whatever mc:Fallback provides (this
        # is usually the legacy VML representation of the same drawing).
        chosen = element.find(OoxmlNamespaces.qn("mc", "Fallback"))
        if chosen is None:
            chosen = element.find(OoxmlNamespaces.qn("mc", "Choice"))
        if chosen is not None:
            yield from self._flatten_block_children(chosen)

    # -- Inline-level container flattening -----------------------------------
    #
    # ``w:hyperlink``, ``w:ins``, ``w:moveTo``, and ``w:smartTag`` wrap runs
    # transparently (their content should render exactly like an ordinary
    # run stream).  ``w:del``/``w:moveFrom`` wrap rejected/relocated content
    # that must not render.  ``w:sdt`` and ``mc:AlternateContent`` reuse the
    # same unwrapping rules as the block-level case above.  Everything else
    # (bookmarks, proofing errors, comment anchors, paragraph properties) is
    # simply not a recognized leaf and is ignored by the caller.

    _TRANSPARENT_INLINE_CONTAINERS = (
        "hyperlink",
        "ins",
        "moveTo",
        "smartTag",
    )
    _DROPPED_INLINE_CONTAINERS = ("del", "moveFrom")

    def _flatten_inline_children(self, parent: Element) -> Iterator[Element]:
        for child in parent:
            yield from self._flatten_inline_element(child)

    def _flatten_inline_element(self, element: Element) -> Iterator[Element]:
        if element.tag in {
            OoxmlNamespaces.qn("w", name) for name in self._TRANSPARENT_INLINE_CONTAINERS
        }:
            yield from self._flatten_inline_children(element)
            return
        if element.tag in {
            OoxmlNamespaces.qn("w", name) for name in self._DROPPED_INLINE_CONTAINERS
        }:
            return
        if element.tag == OoxmlNamespaces.qn("w", "sdt"):
            content = element.find(OoxmlNamespaces.qn("w", "sdtContent"))
            if content is not None:
                yield from self._flatten_inline_children(content)
            return
        if element.tag == OoxmlNamespaces.qn("mc", "AlternateContent"):
            chosen = element.find(OoxmlNamespaces.qn("mc", "Fallback"))
            if chosen is None:
                chosen = element.find(OoxmlNamespaces.qn("mc", "Choice"))
            if chosen is not None:
                yield from self._flatten_inline_children(chosen)
            return
        yield element

    def _record_placeholder(self, *, part_name: str, label: str) -> None:
        """Record a non-fatal diagnostic for a same-size placeholder.

        Placeholders are always reported, in both strict and lenient mode:
        the content's footprint is preserved (so pagination matches the
        source document) but its exact appearance is not, and that
        degradation should never be silent.
        """

        location = f"placeholder[{len(self._placeholder_features)}]"
        feature = UnsupportedFeature(
            name="content_placeholder",
            part=part_name,
            element=label,
            location=location,
            status="placeholder",
            workaround=(
                "Install the 'images' extra (Pillow) for broader image format "
                "support, or replace the content with a native PNG/JPEG image."
            ),
        )
        warning = ConversionWarning(
            code="content_placeholder",
            message=(
                f"{label} could not be rendered natively and was replaced by a "
                "same-size placeholder"
            ),
            part=part_name,
            location=location,
            feature=feature,
        )
        self._placeholder_features.append(feature)
        self._placeholder_warnings.append(warning)

    def _parse_paragraph(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
    ) -> ParagraphModel:
        paragraph_index = state.count_paragraph(part_name=part_name)
        properties_element = element.find(OoxmlNamespaces.qn("w", "pPr"))
        properties, style_id = self._parse_paragraph_properties(properties_element)
        runs: list[RunModel] = []
        tracker = _FieldTracker()
        run_tag = OoxmlNamespaces.qn("w", "r")
        fld_simple_tag = OoxmlNamespaces.qn("w", "fldSimple")
        fld_char_tag = OoxmlNamespaces.qn("w", "fldChar")
        instr_text_tag = OoxmlNamespaces.qn("w", "instrText")
        for child in self._flatten_inline_children(element):
            if child.tag == run_tag:
                field_chars = child.findall(fld_char_tag)
                if field_chars:
                    for field_char in field_chars:
                        field_type = field_char.get(OoxmlNamespaces.qn("w", "fldCharType"), "")
                        if field_type == "begin":
                            tracker.begin()
                        elif field_type == "separate":
                            tracker.separate()
                        elif field_type == "end":
                            runs.extend(tracker.end())
                    continue
                instr_texts = child.findall(instr_text_tag)
                if instr_texts:
                    for instr_text in instr_texts:
                        tracker.record_instruction(instr_text.text or "")
                    continue
                produced = self._parse_run(
                    child,
                    package=package,
                    part_name=part_name,
                    state=state,
                )
                self._route_field_content(produced, tracker=tracker, runs=runs)
            elif child.tag == fld_simple_tag:
                self._parse_fld_simple(
                    child,
                    package=package,
                    part_name=part_name,
                    state=state,
                    runs=runs,
                )
        runs.extend(tracker.flush_incomplete())
        return ParagraphModel(
            runs=tuple(runs),
            style_id=style_id,
            properties=properties,
            source_index=paragraph_index,
        )

    @staticmethod
    def _route_field_content(
        produced: list[RunModel],
        *,
        tracker: _FieldTracker,
        runs: list[RunModel],
    ) -> None:
        if tracker.depth == 0:
            runs.extend(produced)
        elif tracker.should_buffer_result():
            for run in produced:
                tracker.buffer(run)
        # else: nested field instruction/result content is discarded, matching
        # how Word only bakes the outermost field's cached result into view.

    def _parse_fld_simple(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
        runs: list[RunModel],
    ) -> None:
        field_runs: list[RunModel] = []
        for run_element in element.findall(OoxmlNamespaces.qn("w", "r")):
            field_runs.extend(
                self._parse_run(
                    run_element,
                    package=package,
                    part_name=part_name,
                    state=state,
                )
            )
        instruction = element.get(OoxmlNamespaces.qn("w", "instr"), "")
        command = instruction.strip().split(maxsplit=1)[0].upper() if instruction.strip() else ""
        sentinel = {"PAGE": "{{PAGE}}", "NUMPAGES": "{{NUMPAGES}}"}.get(command)
        if sentinel is None:
            # Any other field code (DATE, FILENAME, REF, ...) is rendered as
            # its cached child-run text, exactly as Word last computed it.
            runs.extend(field_runs)
        elif field_runs:
            first = field_runs[0]
            runs.append(
                RunModel(
                    text=sentinel,
                    style_id=first.style_id,
                    properties=first.properties,
                    preserve_space=first.preserve_space,
                    hidden=first.hidden,
                    source_index=first.source_index,
                )
            )
        else:
            runs.append(
                RunModel(
                    text=sentinel,
                    source_index=state.count_run(part_name=part_name),
                )
            )

    def _parse_paragraph_properties(
        self,
        element: Element | None,
    ) -> tuple[ParagraphProperties, str | None]:
        if element is None:
            return ParagraphProperties(), None
        values: dict[str, object] = {}
        style = element.find(OoxmlNamespaces.qn("w", "pStyle"))
        style_id = style.get(OoxmlNamespaces.qn("w", "val")) if style is not None else None

        alignment = element.find(OoxmlNamespaces.qn("w", "jc"))
        if alignment is not None:
            raw = alignment.get(OoxmlNamespaces.qn("w", "val"), "left")
            mapping = {
                "left": "left",
                "start": "left",
                "center": "center",
                "right": "right",
                "end": "right",
                "both": "justify",
                "distribute": "distribute",
            }
            mapped = mapping.get(raw)
            if mapped is None:
                raise InvalidOoxmlError(f"Unsupported paragraph alignment {raw!r}")
            values["alignment"] = mapped

        indent = element.find(OoxmlNamespaces.qn("w", "ind"))
        if indent is not None:
            for attribute, field in (
                ("left", "left_indent"),
                ("start", "left_indent"),
                ("right", "right_indent"),
                ("end", "right_indent"),
                ("firstLine", "first_line_indent"),
                ("hanging", "hanging_indent"),
            ):
                indent_value = self._optional_int(indent, "w", attribute)
                if indent_value is not None:
                    values[field] = self.twips_to_points(indent_value)

        spacing = element.find(OoxmlNamespaces.qn("w", "spacing"))
        if spacing is not None:
            before = self._optional_int(spacing, "w", "before")
            after = self._optional_int(spacing, "w", "after")
            if before is not None:
                values["space_before"] = max(0.0, self.twips_to_points(before))
            if after is not None:
                values["space_after"] = max(0.0, self.twips_to_points(after))
            line = self._optional_int(spacing, "w", "line")
            if line is not None and line > 0:
                rule = spacing.get(OoxmlNamespaces.qn("w", "lineRule"), "auto")
                rule_mapping = {"auto": "auto", "atLeast": "at_least", "exact": "exact"}
                mapped_rule = rule_mapping.get(rule)
                if mapped_rule is None:
                    raise InvalidOoxmlError(f"Invalid line spacing rule {rule!r}")
                values["line_spacing_rule"] = mapped_rule
                values["line_spacing"] = (
                    round(line / 240.0, 6) if mapped_rule == "auto" else self.twips_to_points(line)
                )

        for child_name, field in (
            ("keepNext", "keep_next"),
            ("keepLines", "keep_lines"),
            ("pageBreakBefore", "page_break_before"),
            ("widowControl", "widow_control"),
        ):
            child = element.find(OoxmlNamespaces.qn("w", child_name))
            if child is not None:
                values[field] = self._on_off(child)

        tabs = element.find(OoxmlNamespaces.qn("w", "tabs"))
        if tabs is not None:
            parsed_tabs: list[TabStop] = []
            for tab in tabs.findall(OoxmlNamespaces.qn("w", "tab")):
                position = self._optional_int(tab, "w", "pos")
                if position is None:
                    raise InvalidOoxmlError("Tab stop is missing w:pos")
                alignment_value = tab.get(OoxmlNamespaces.qn("w", "val"), "left")
                leader_value = tab.get(OoxmlNamespaces.qn("w", "leader"), "none")
                leader_mapping = {
                    "none": "none",
                    "dot": "dot",
                    "hyphen": "hyphen",
                    "underscore": "underscore",
                    "middleDot": "middle_dot",
                    "heavy": "heavy",
                }
                leader = leader_mapping.get(leader_value)
                if alignment_value not in {
                    "left",
                    "center",
                    "right",
                    "decimal",
                    "bar",
                    "clear",
                    "num",
                }:
                    raise InvalidOoxmlError(f"Invalid tab alignment {alignment_value!r}")
                if leader is None:
                    raise InvalidOoxmlError(f"Invalid tab leader {leader_value!r}")
                parsed_tabs.append(
                    TabStop(
                        position=self.twips_to_points(position),
                        alignment=alignment_value,
                        leader=leader,
                    )
                )
            values["tabs"] = tuple(parsed_tabs)

        numbering = element.find(OoxmlNamespaces.qn("w", "numPr"))
        if numbering is not None:
            numbering_id = numbering.find(OoxmlNamespaces.qn("w", "numId"))
            numbering_level = numbering.find(OoxmlNamespaces.qn("w", "ilvl"))
            if numbering_id is not None:
                values["numbering_id"] = self._required_int(numbering_id, "w", "val")
            if numbering_level is not None:
                values["numbering_level"] = self._required_int(numbering_level, "w", "val")
        return ParagraphProperties.model_validate(values), style_id

    def _parse_run(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
    ) -> list[RunModel]:
        properties_element = element.find(OoxmlNamespaces.qn("w", "rPr"))
        properties, style_id = self._parse_run_properties(properties_element)
        hidden = properties.hidden is True
        parsed: list[RunModel] = []
        for child in element:
            if child is properties_element:
                continue
            run_index = state.count_run(part_name=part_name)
            common = {
                "style_id": style_id,
                "properties": properties,
                "hidden": hidden,
                "source_index": run_index,
            }
            if child.tag in {
                OoxmlNamespaces.qn("w", "t"),
                OoxmlNamespaces.qn("w", "delText"),
            }:
                text = unicodedata.normalize("NFC", child.text or "")
                preserve = child.get(OoxmlNamespaces.qn("xml", "space")) == "preserve"
                parsed.append(RunModel(text=text, preserve_space=preserve, **common))
            elif child.tag in {
                OoxmlNamespaces.qn("w", "br"),
                OoxmlNamespaces.qn("w", "cr"),
            }:
                break_value = child.get(OoxmlNamespaces.qn("w", "type"), "textWrapping")
                break_type: Literal["line", "page"] = "page" if break_value == "page" else "line"
                parsed.append(RunModel(break_type=break_type, **common))
            elif child.tag == OoxmlNamespaces.qn("w", "tab"):
                parsed.append(RunModel(tab=True, **common))
            elif child.tag == OoxmlNamespaces.qn("w", "drawing"):
                for drawing_visual in self._parse_drawings(
                    child, package=package, part_name=part_name
                ):
                    parsed.append(self._visual_run(drawing_visual, **common))
            elif child.tag == OoxmlNamespaces.qn("w", "pict"):
                pict_visual = self._parse_vml_visual(
                    child, package=package, part_name=part_name, is_object=False
                )
                if pict_visual is not None:
                    parsed.append(self._visual_run(pict_visual, **common))
            elif child.tag == OoxmlNamespaces.qn("w", "object"):
                object_visual = self._parse_vml_visual(
                    child, package=package, part_name=part_name, is_object=True
                )
                if object_visual is not None:
                    parsed.append(self._visual_run(object_visual, **common))
            else:
                state.runs -= 1
        if not parsed and len(element) == (1 if properties_element is not None else 0):
            parsed.append(
                RunModel(
                    style_id=style_id,
                    properties=properties,
                    hidden=hidden,
                    source_index=state.count_run(part_name=part_name),
                )
            )
        return parsed

    def _parse_run_properties(
        self,
        element: Element | None,
    ) -> tuple[RunProperties, str | None]:
        if element is None:
            return RunProperties(), None
        values: dict[str, object] = {}
        style = element.find(OoxmlNamespaces.qn("w", "rStyle"))
        style_id = style.get(OoxmlNamespaces.qn("w", "val")) if style is not None else None
        fonts = element.find(OoxmlNamespaces.qn("w", "rFonts"))
        if fonts is not None:
            attributes = {
                "ascii": ("ascii_font", "asciiTheme"),
                "hAnsi": ("high_ansi_font", "hAnsiTheme"),
                "eastAsia": ("east_asia_font", "eastAsiaTheme"),
                "cs": ("complex_script_font", "csTheme"),
            }
            for attribute, (field, theme_attribute) in attributes.items():
                value = fonts.get(OoxmlNamespaces.qn("w", attribute))
                if not value:
                    theme_value = fonts.get(OoxmlNamespaces.qn("w", theme_attribute))
                    value = f"+{theme_value}" if theme_value else None
                if value:
                    values[field] = value
            family = values.get("ascii_font") or values.get("high_ansi_font")
            if family:
                values["font_family"] = family
        size = element.find(OoxmlNamespaces.qn("w", "sz"))
        if size is not None:
            values["font_size"] = self.half_points_to_points(self._required_int(size, "w", "val"))
        for child_name, field in (("b", "bold"), ("i", "italic"), ("strike", "strike")):
            child = element.find(OoxmlNamespaces.qn("w", child_name))
            if child is not None:
                values[field] = self._on_off(child)
        underline = element.find(OoxmlNamespaces.qn("w", "u"))
        if underline is not None:
            raw_underline = underline.get(OoxmlNamespaces.qn("w", "val"), "single")
            values["underline"] = False if raw_underline in {"none", "0"} else raw_underline
        color = element.find(OoxmlNamespaces.qn("w", "color"))
        if color is not None:
            raw_color = color.get(OoxmlNamespaces.qn("w", "val"))
            try:
                values["color"] = OoxmlNamespaces.normalize_color(raw_color)
            except ValueError as error:
                raise InvalidOoxmlError(str(error)) from error
        highlight = element.find(OoxmlNamespaces.qn("w", "highlight"))
        if highlight is not None:
            value = highlight.get(OoxmlNamespaces.qn("w", "val"))
            try:
                values["highlight"] = OoxmlNamespaces.normalize_highlight(value)
            except ValueError as error:
                raise InvalidOoxmlError(str(error)) from error
        vertical_align = element.find(OoxmlNamespaces.qn("w", "vertAlign"))
        if vertical_align is not None:
            value = vertical_align.get(OoxmlNamespaces.qn("w", "val"), "baseline")
            if value not in {"baseline", "superscript", "subscript"}:
                raise InvalidOoxmlError(f"Invalid vertical alignment {value!r}")
            values["vertical_align"] = value
        spacing = element.find(OoxmlNamespaces.qn("w", "spacing"))
        if spacing is not None:
            values["character_spacing"] = self.twips_to_points(
                self._required_int(spacing, "w", "val")
            )
        hidden = element.find(OoxmlNamespaces.qn("w", "vanish"))
        if hidden is not None:
            values["hidden"] = self._on_off(hidden)
        return RunProperties.model_validate(values), style_id

    @staticmethod
    def _visual_run(visual: Visual, **common: object) -> RunModel:
        if isinstance(visual, ImageModel):
            return RunModel(image=visual, **common)
        return RunModel(placeholder=visual, **common)

    @staticmethod
    def _safe_dimension(value: float) -> float:
        return value if value > 0 else 0.01

    # -- DrawingML (wp:inline / wp:anchor) -----------------------------------

    def _parse_drawings(
        self,
        drawing: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
    ) -> list[Visual]:
        elements = [
            *drawing.findall(f".//{OoxmlNamespaces.qn('wp', 'inline')}"),
            *drawing.findall(f".//{OoxmlNamespaces.qn('wp', 'anchor')}"),
        ]
        visuals: list[Visual] = []
        for element in elements:
            visual = self._parse_one_drawing(element, package=package, part_name=part_name)
            if visual is not None:
                visuals.append(visual)
        return visuals

    def _parse_one_drawing(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
    ) -> Visual | None:
        extent = element.find(OoxmlNamespaces.qn("wp", "extent"))
        if extent is None:
            # No declared size means no reliable footprint to reserve; skip
            # rather than guess, this is rare and indicates malformed markup.
            return None
        width = self._safe_dimension(self.emu_to_points(self._required_plain_int(extent, "cx")))
        height = self._safe_dimension(self.emu_to_points(self._required_plain_int(extent, "cy")))
        document_properties = element.find(OoxmlNamespaces.qn("wp", "docPr"))
        description = None
        if document_properties is not None:
            description = document_properties.get("descr") or document_properties.get("title")
        blip = element.find(f".//{OoxmlNamespaces.qn('a', 'blip')}")
        if blip is not None:
            relation_id = blip.get(OoxmlNamespaces.qn("r", "embed")) or blip.get(
                OoxmlNamespaces.qn("r", "link")
            )
            image = (
                self._resolve_relationship_image(
                    relation_id,
                    package=package,
                    part_name=part_name,
                    width=width,
                    height=height,
                    description=description,
                )
                if relation_id
                else None
            )
            if image is not None:
                return image
            self._record_placeholder(part_name=part_name, label="[Image]")
            return PlaceholderModel(label="[Image]", width=width, height=height)
        label = self._drawing_placeholder_label(element)
        self._record_placeholder(part_name=part_name, label=label)
        return PlaceholderModel(label=label, width=width, height=height)

    @staticmethod
    def _drawing_placeholder_label(element: Element) -> str:
        if element.find(f".//{OoxmlNamespaces.qn('c', 'chart')}") is not None:
            return "[Chart]"
        if element.find(f".//{OoxmlNamespaces.qn('dgm', 'relIds')}") is not None:
            return "[SmartArt]"
        if (
            element.find(f".//{OoxmlNamespaces.qn('wps', 'txbx')}") is not None
            or element.find(f".//{OoxmlNamespaces.qn('w', 'txbxContent')}") is not None
        ):
            return "[Text Box]"
        return "[Shape]"

    # -- VML (w:pict / w:object) and embedded OLE objects --------------------

    def _parse_vml_visual(
        self,
        container: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        is_object: bool,
    ) -> Visual | None:
        shape = container.find(f".//{OoxmlNamespaces.qn('v', 'shape')}")
        if shape is None:
            return None
        size = self._parse_vml_style_size(shape.get("style")) or self._vml_fallback_size(container)
        if size is None:
            return None
        width, height = self._safe_dimension(size[0]), self._safe_dimension(size[1])
        object_label = "[Embedded Object]" if is_object else "[Image]"
        image_data = shape.find(OoxmlNamespaces.qn("v", "imagedata"))
        if image_data is not None:
            relation_id = image_data.get(OoxmlNamespaces.qn("r", "id"))
            image = (
                self._resolve_relationship_image(
                    relation_id,
                    package=package,
                    part_name=part_name,
                    width=width,
                    height=height,
                    description=None,
                )
                if relation_id
                else None
            )
            if image is not None:
                return image
            self._record_placeholder(part_name=part_name, label=object_label)
            return PlaceholderModel(label=object_label, width=width, height=height)
        has_textpath = shape.find(f".//{OoxmlNamespaces.qn('v', 'textpath')}") is not None
        if is_object:
            label = "[Embedded Object]"
        elif has_textpath:
            label = "[WordArt]"
        else:
            label = "[Shape]"
        self._record_placeholder(part_name=part_name, label=label)
        return PlaceholderModel(label=label, width=width, height=height)

    @staticmethod
    def _vml_fallback_size(container: Element) -> tuple[float, float] | None:
        # ``w:object`` may omit a usable VML ``style`` and instead carry the
        # original object size directly, in twentieths of a point.
        width = OoxmlDocumentParser._optional_int(container, "w", "dxaOrig")
        height = OoxmlDocumentParser._optional_int(container, "w", "dyaOrig")
        if width is None or height is None:
            return None
        return (
            OoxmlDocumentParser.twips_to_points(width),
            OoxmlDocumentParser.twips_to_points(height),
        )

    @classmethod
    def _parse_vml_style_size(cls, style: str | None) -> tuple[float, float] | None:
        if not style:
            return None
        properties: dict[str, str] = {}
        for declaration in style.split(";"):
            if ":" not in declaration:
                continue
            key, _, raw_value = declaration.partition(":")
            properties[key.strip().casefold()] = raw_value.strip()
        width = cls._parse_vml_length(properties.get("width"))
        height = cls._parse_vml_length(properties.get("height"))
        if width is None or height is None:
            return None
        return width, height

    @staticmethod
    def _parse_vml_length(value: str | None) -> float | None:
        if not value:
            return None
        match = _VML_LENGTH_PATTERN.fullmatch(value)
        if match is None:
            return None
        number = float(match.group(1))
        unit = match.group(2).casefold() or "pt"
        factor = _VML_UNITS_TO_POINTS.get(unit)
        if factor is None:
            return None
        length = number * factor
        return length if length > 0 else None

    # -- Shared relationship-backed image resolution -------------------------

    def _resolve_relationship_image(
        self,
        relation_id: str,
        *,
        package: OoxmlPackage,
        part_name: str,
        width: float,
        height: float,
        description: str | None,
    ) -> ImageModel | None:
        try:
            relation = package.relationships_for(part_name).by_id(relation_id)
        except RelationshipError:
            return None
        if relation.is_external or relation.resolved_target is None:
            # External bytes are never fetched; the caller falls back to a
            # dimension-preserving placeholder instead.
            return None
        image_part = relation.resolved_target
        if not package.has_part(image_part):
            return None
        raw_content_type = package.content_types.for_part(image_part)
        content_type = "image/jpeg" if raw_content_type == "image/jpg" else raw_content_type
        data = package.read_part(image_part)
        if content_type not in {"image/png", "image/jpeg"}:
            converted = RasterImageConverter.convert_to_png(data)
            if converted is None:
                return None
            data = converted
            content_type = "image/png"
        return ImageModel(
            relationship_id=relation_id,
            part_name=image_part,
            content_type=content_type,
            data=data,
            width=width,
            height=height,
            description=description,
        )

    def _parse_table(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
    ) -> TableModel:
        values: dict[str, object] = {"source_index": state.next_block()}
        properties = element.find(OoxmlNamespaces.qn("w", "tblPr"))
        table_margins = CellMargins()
        if properties is not None:
            width = properties.find(OoxmlNamespaces.qn("w", "tblW"))
            if width is not None and width.get(OoxmlNamespaces.qn("w", "type"), "dxa") == "dxa":
                width_value = self._optional_int(width, "w", "w")
                if width_value is not None and width_value > 0:
                    values["width"] = self.twips_to_points(width_value)
            layout = properties.find(OoxmlNamespaces.qn("w", "tblLayout"))
            values["autofit"] = not (
                layout is not None
                and layout.get(OoxmlNamespaces.qn("w", "type"), "autofit") == "fixed"
            )
            alignment = properties.find(OoxmlNamespaces.qn("w", "jc"))
            if alignment is not None:
                raw_alignment = alignment.get(OoxmlNamespaces.qn("w", "val"), "left")
                if raw_alignment in {"left", "center", "right"}:
                    values["alignment"] = raw_alignment
            indent = properties.find(OoxmlNamespaces.qn("w", "tblInd"))
            if indent is not None:
                indent_value = self._optional_int(indent, "w", "w")
                if indent_value is not None:
                    values["left_indent"] = self.twips_to_points(indent_value)
            margin_element = properties.find(OoxmlNamespaces.qn("w", "tblCellMar"))
            if margin_element is not None:
                table_margins = self._parse_cell_margins(margin_element)
                values["cell_margins"] = table_margins
            border_element = properties.find(OoxmlNamespaces.qn("w", "tblBorders"))
            if border_element is not None:
                values["borders"] = self._parse_borders(border_element)

        grid = element.find(OoxmlNamespaces.qn("w", "tblGrid"))
        if grid is not None:
            widths: list[float] = []
            for column in grid.findall(OoxmlNamespaces.qn("w", "gridCol")):
                column_width = self._optional_int(column, "w", "w")
                if column_width is not None and column_width > 0:
                    widths.append(self.twips_to_points(column_width))
            values["grid_widths"] = tuple(widths)

        row_tag = OoxmlNamespaces.qn("w", "tr")
        rows = tuple(
            self._parse_table_row(
                row,
                package=package,
                part_name=part_name,
                state=state,
                table_margins=table_margins,
            )
            for row in self._flatten_block_children(element)
            if row.tag == row_tag
        )
        values["rows"] = rows
        return TableModel.model_validate(values)

    def _parse_table_row(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
        table_margins: CellMargins,
    ) -> TableRowModel:
        values: dict[str, object] = {}
        properties = element.find(OoxmlNamespaces.qn("w", "trPr"))
        if properties is not None:
            height = properties.find(OoxmlNamespaces.qn("w", "trHeight"))
            if height is not None:
                raw_height = self._optional_int(height, "w", "val")
                if raw_height is not None and raw_height > 0:
                    values["height"] = self.twips_to_points(raw_height)
                rule = height.get(OoxmlNamespaces.qn("w", "hRule"), "auto")
                rule_mapping = {"auto": "auto", "atLeast": "at_least", "exact": "exact"}
                mapped_rule = rule_mapping.get(rule)
                if mapped_rule is None:
                    raise InvalidOoxmlError(f"Invalid table row height rule {rule!r}")
                values["height_rule"] = mapped_rule
            cant_split = properties.find(OoxmlNamespaces.qn("w", "cantSplit"))
            if cant_split is not None:
                values["cant_split"] = self._on_off(cant_split)
            repeat = properties.find(OoxmlNamespaces.qn("w", "tblHeader"))
            if repeat is not None:
                values["repeat_header"] = self._on_off(repeat)
        cell_tag = OoxmlNamespaces.qn("w", "tc")
        values["cells"] = tuple(
            self._parse_table_cell(
                cell,
                package=package,
                part_name=part_name,
                state=state,
                table_margins=table_margins,
            )
            for cell in self._flatten_block_children(element)
            if cell.tag == cell_tag
        )
        return TableRowModel.model_validate(values)

    def _parse_table_cell(
        self,
        element: Element,
        *,
        package: OoxmlPackage,
        part_name: str,
        state: _ParseState,
        table_margins: CellMargins,
    ) -> TableCellModel:
        cell_index = state.count_cell(part_name=part_name)
        values: dict[str, object] = {
            "source_index": cell_index,
            "margins": table_margins,
        }
        properties = element.find(OoxmlNamespaces.qn("w", "tcPr"))
        if properties is not None:
            width = properties.find(OoxmlNamespaces.qn("w", "tcW"))
            if width is not None and width.get(OoxmlNamespaces.qn("w", "type"), "dxa") == "dxa":
                width_value = self._optional_int(width, "w", "w")
                if width_value is not None and width_value > 0:
                    values["width"] = self.twips_to_points(width_value)
            span = properties.find(OoxmlNamespaces.qn("w", "gridSpan"))
            if span is not None:
                values["grid_span"] = max(1, self._required_int(span, "w", "val"))
            merge = properties.find(OoxmlNamespaces.qn("w", "vMerge"))
            if merge is not None:
                merge_value = merge.get(OoxmlNamespaces.qn("w", "val"), "continue")
                if merge_value not in {"restart", "continue"}:
                    raise InvalidOoxmlError(f"Invalid vertical merge value {merge_value!r}")
                values["vertical_merge"] = merge_value
            margins = properties.find(OoxmlNamespaces.qn("w", "tcMar"))
            if margins is not None:
                values["margins"] = self._parse_cell_margins(margins)
            borders = properties.find(OoxmlNamespaces.qn("w", "tcBorders"))
            if borders is not None:
                values["borders"] = self._parse_borders(borders)
            shading = properties.find(OoxmlNamespaces.qn("w", "shd"))
            if shading is not None:
                try:
                    values["background_color"] = OoxmlNamespaces.normalize_color(
                        shading.get(OoxmlNamespaces.qn("w", "fill"))
                    )
                except ValueError as error:
                    raise InvalidOoxmlError(str(error)) from error
            vertical_alignment = properties.find(OoxmlNamespaces.qn("w", "vAlign"))
            if vertical_alignment is not None:
                raw_alignment = vertical_alignment.get(OoxmlNamespaces.qn("w", "val"), "top")
                if raw_alignment not in {"top", "center", "bottom"}:
                    raise InvalidOoxmlError(
                        f"Invalid table cell vertical alignment {raw_alignment!r}"
                    )
                values["vertical_alignment"] = raw_alignment

        paragraph_tag = OoxmlNamespaces.qn("w", "p")
        paragraphs = tuple(
            self._parse_paragraph(
                paragraph,
                package=package,
                part_name=part_name,
                state=state,
            )
            for paragraph in self._flatten_block_children(element)
            if paragraph.tag == paragraph_tag
        )
        values["paragraphs"] = paragraphs
        if paragraphs and paragraphs[0].properties.alignment is not None:
            values["text_alignment"] = paragraphs[0].properties.alignment
        return TableCellModel.model_validate(values)

    def _parse_cell_margins(self, element: Element) -> CellMargins:
        values: dict[str, float] = {}
        for side in ("top", "right", "bottom", "left"):
            side_element = element.find(OoxmlNamespaces.qn("w", side))
            if side_element is None:
                continue
            raw = self._optional_int(side_element, "w", "w")
            if raw is not None:
                values[side] = max(0.0, self.twips_to_points(raw))
        return CellMargins.model_validate(values)

    def _parse_borders(self, element: Element) -> TableBorders:
        values: dict[str, BorderModel | None] = {}
        side_mapping = {
            "top": "top",
            "right": "right",
            "bottom": "bottom",
            "left": "left",
            "insideH": "inside_horizontal",
            "insideV": "inside_vertical",
        }
        for child_name, field in side_mapping.items():
            child = element.find(OoxmlNamespaces.qn("w", child_name))
            if child is None:
                continue
            style = child.get(OoxmlNamespaces.qn("w", "val"), "single")
            if style in {"nil", "none"}:
                values[field] = None
                continue
            size = self._optional_int(child, "w", "sz") or 4
            try:
                color = (
                    OoxmlNamespaces.normalize_color(child.get(OoxmlNamespaces.qn("w", "color")))
                    or "000000"
                )
            except ValueError as error:
                raise InvalidOoxmlError(str(error)) from error
            values[field] = BorderModel(style=style, width=round(size / 8.0, 4), color=color)
        return TableBorders.model_validate(values)

    @staticmethod
    def twips_to_points(value: int) -> float:
        """Convert twentieths of a point to points with deterministic rounding."""

        return OoxmlDocumentParser._to_points(value, divisor=20)

    @staticmethod
    def emu_to_points(value: int) -> float:
        """Convert English Metric Units to points."""

        return OoxmlDocumentParser._to_points(value, divisor=12_700)

    @staticmethod
    def half_points_to_points(value: int) -> float:
        """Convert half-points to points."""

        return OoxmlDocumentParser._to_points(value, divisor=2)

    @staticmethod
    def _to_points(value: int, *, divisor: int) -> float:
        converted = float(Decimal(value) / Decimal(divisor))
        if not math.isfinite(converted) or abs(converted) > 1_000_000:
            raise InvalidOoxmlError("OOXML length is outside the supported point range")
        return round(converted, 6)

    @staticmethod
    def _on_off(element: Element) -> bool:
        value = element.get(OoxmlNamespaces.qn("w", "val"))
        try:
            return OoxmlNamespaces.parse_on_off(value)
        except ValueError as error:
            raise InvalidOoxmlError(str(error)) from error

    @staticmethod
    def _optional_int(element: Element, prefix: str, name: str) -> int | None:
        value = element.get(OoxmlNamespaces.qn(prefix, name))
        if value is None:
            return None
        try:
            return int(value)
        except ValueError as error:
            raise InvalidOoxmlError(
                f"Attribute {prefix}:{name} must be an integer, got {value!r}"
            ) from error

    @staticmethod
    def _required_int(element: Element, prefix: str, name: str) -> int:
        value = OoxmlDocumentParser._optional_int(element, prefix, name)
        if value is None:
            raise InvalidOoxmlError(f"Missing required integer attribute {prefix}:{name}")
        return value

    @staticmethod
    def _required_plain_int(element: Element, name: str) -> int:
        value = element.get(name)
        if value is None:
            raise InvalidOoxmlError(f"Missing required integer attribute {name}")
        try:
            return int(value)
        except ValueError as error:
            raise InvalidOoxmlError(
                f"Attribute {name} must be an integer, got {value!r}"
            ) from error
