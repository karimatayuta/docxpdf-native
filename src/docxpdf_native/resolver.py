"""Build the fully resolved document model used by the layout engine."""

from __future__ import annotations

import unicodedata
from pathlib import Path

from docxpdf_native.abstractions import FontResolver, StyleResolver
from docxpdf_native.exceptions import InvalidOoxmlError
from docxpdf_native.models import (
    DocumentModel,
    HeaderFooterModel,
    NumberingDefinition,
    ParagraphModel,
    ParagraphProperties,
    ResolvedDocumentModel,
    ResolvedParagraphModel,
    ResolvedRunModel,
    ResolvedSectionModel,
    RunModel,
    RunProperties,
    TableCellModel,
    TableModel,
)
from docxpdf_native.ooxml.styles import OoxmlStyleResolver


class NativeStyleResolver(StyleResolver):
    """Resolve OOXML inheritance, direct formatting, themes, and concrete fonts."""

    def __init__(self, font_resolver: FontResolver) -> None:
        self._font_resolver = font_resolver
        self._number_counters: dict[tuple[int, int], int] = {}
        self._font_paths: dict[str, Path] = {}

    @property
    def font_paths(self) -> dict[str, Path]:
        """Return the concrete font files selected during the latest resolution."""
        return dict(self._font_paths)

    def resolve(self, document: DocumentModel) -> ResolvedDocumentModel:
        style_resolver = OoxmlStyleResolver(
            styles=document.styles,
            document_run_defaults=document.defaults.run,
            document_paragraph_defaults=document.defaults.paragraph,
            theme_fonts=document.theme_fonts,
        )
        default_paragraph_style = next(
            (
                style.style_id
                for style in document.styles
                if style.style_type == "paragraph" and style.is_default
            ),
            None,
        )
        sections: list[ResolvedSectionModel] = []
        all_paragraphs: list[ResolvedParagraphModel] = []
        for section in document.sections:
            blocks: list[ResolvedParagraphModel | TableModel] = []
            for block in section.blocks:
                if isinstance(block, ParagraphModel):
                    resolved = self._resolve_paragraph(
                        block,
                        style_resolver=style_resolver,
                        default_style_id=default_paragraph_style,
                        numbering=document.numbering,
                    )
                    blocks.append(resolved)
                    all_paragraphs.append(resolved)
                else:
                    blocks.append(
                        self._resolve_table(
                            block,
                            style_resolver=style_resolver,
                            default_style_id=default_paragraph_style,
                            numbering=document.numbering,
                        )
                    )
            headers = tuple(
                self._resolve_header_footer(
                    header,
                    style_resolver=style_resolver,
                    default_style_id=default_paragraph_style,
                    numbering=document.numbering,
                )
                for header in section.headers
            )
            footers = tuple(
                self._resolve_header_footer(
                    footer,
                    style_resolver=style_resolver,
                    default_style_id=default_paragraph_style,
                    numbering=document.numbering,
                )
                for footer in section.footers
            )
            values = section.model_dump(exclude={"blocks", "headers", "footers", "columns"})
            sections.append(
                ResolvedSectionModel(
                    **values,
                    blocks=tuple(blocks),
                    headers=headers,
                    footers=footers,
                )
            )
        return ResolvedDocumentModel(
            sections=tuple(sections),
            paragraphs=tuple(all_paragraphs),
            source=document,
        )

    def _resolve_paragraph(
        self,
        paragraph: ParagraphModel,
        *,
        style_resolver: OoxmlStyleResolver,
        default_style_id: str | None,
        numbering: tuple[NumberingDefinition, ...],
    ) -> ResolvedParagraphModel:
        paragraph_style_id = paragraph.style_id or default_style_id
        try:
            paragraph_style = style_resolver.resolve(
                paragraph_style_id=paragraph_style_id,
                paragraph_direct=paragraph.properties,
            )
            runs = tuple(
                resolved_run
                for run_index, run in enumerate(paragraph.runs)
                for resolved_run in self._resolve_run(
                    run,
                    paragraph=paragraph,
                    run_index=run_index,
                    paragraph_style_id=paragraph_style_id,
                    paragraph_direct=paragraph.properties,
                    style_resolver=style_resolver,
                )
            )
        except ValueError as error:
            raise InvalidOoxmlError(
                f"style resolution failed: {error}",
                part="word/styles.xml",
                location=f"paragraph[{paragraph.source_index}]",
                cause=error,
            ) from error
        properties = paragraph_style.paragraph
        numbering_label = self._numbering_label(properties, numbering)
        return ResolvedParagraphModel(
            runs=runs,
            alignment=properties.alignment or "left",
            left_indent=properties.left_indent or 0.0,
            right_indent=properties.right_indent or 0.0,
            first_line_indent=properties.first_line_indent or 0.0,
            hanging_indent=properties.hanging_indent or 0.0,
            space_before=properties.space_before or 0.0,
            space_after=properties.space_after or 0.0,
            line_spacing=properties.line_spacing or 1.0,
            line_spacing_rule=properties.line_spacing_rule or "auto",
            keep_next=bool(properties.keep_next),
            keep_lines=bool(properties.keep_lines),
            page_break_before=bool(properties.page_break_before),
            widow_control=(properties.widow_control is not False),
            tabs=properties.tabs,
            numbering_label=numbering_label,
            source_index=paragraph.source_index,
        )

    def _resolve_run(
        self,
        run: RunModel,
        *,
        paragraph: ParagraphModel,
        run_index: int,
        paragraph_style_id: str | None,
        paragraph_direct: ParagraphProperties,
        style_resolver: OoxmlStyleResolver,
    ) -> tuple[ResolvedRunModel, ...]:
        effective = style_resolver.resolve(
            paragraph_style_id=paragraph_style_id,
            character_style_id=run.style_id,
            paragraph_direct=paragraph_direct,
            run_direct=run.properties,
        ).run
        result: list[ResolvedRunModel] = []
        for segment, east_asia in self._script_segments(run.text):
            requested_font = self._requested_font(effective, east_asia=east_asia)
            resolved_font = self._font_resolver.resolve(
                requested_font,
                east_asia=east_asia,
                paragraph_index=paragraph.source_index,
                run_index=run_index,
            )
            if resolved_font.path is not None:
                self._font_paths.setdefault(resolved_font.family, resolved_font.path)
            result.append(
                ResolvedRunModel(
                    text=segment,
                    font_name=resolved_font.family,
                    font_path=resolved_font.path,
                    font_size=effective.font_size or 11.0,
                    bold=bool(effective.bold),
                    italic=bool(effective.italic),
                    underline=effective.underline or False,
                    strike=bool(effective.strike),
                    color=effective.color or "000000",
                    highlight=effective.highlight,
                    vertical_align=effective.vertical_align or "baseline",
                    character_spacing=effective.character_spacing or 0.0,
                    break_type=run.break_type,
                    tab=run.tab,
                    image=run.image,
                    hidden=run.hidden or bool(effective.hidden),
                    source_index=run.source_index,
                )
            )
        return tuple(result)

    def _resolve_table(
        self,
        table: TableModel,
        *,
        style_resolver: OoxmlStyleResolver,
        default_style_id: str | None,
        numbering: tuple[NumberingDefinition, ...],
    ) -> TableModel:
        rows = tuple(
            row.model_copy(
                update={
                    "cells": tuple(
                        self._resolve_cell(
                            cell,
                            style_resolver=style_resolver,
                            default_style_id=default_style_id,
                            numbering=numbering,
                        )
                        for cell in row.cells
                    )
                }
            )
            for row in table.rows
        )
        return table.model_copy(update={"rows": rows})

    def _resolve_cell(
        self,
        cell: TableCellModel,
        *,
        style_resolver: OoxmlStyleResolver,
        default_style_id: str | None,
        numbering: tuple[NumberingDefinition, ...],
    ) -> TableCellModel:
        paragraphs = tuple(
            self._resolved_to_raw(
                self._resolve_paragraph(
                    paragraph,
                    style_resolver=style_resolver,
                    default_style_id=default_style_id,
                    numbering=numbering,
                )
            )
            for paragraph in cell.paragraphs
        )
        return cell.model_copy(update={"paragraphs": paragraphs})

    def _resolve_header_footer(
        self,
        model: HeaderFooterModel,
        *,
        style_resolver: OoxmlStyleResolver,
        default_style_id: str | None,
        numbering: tuple[NumberingDefinition, ...],
    ) -> HeaderFooterModel:
        blocks = tuple(
            self._resolved_to_raw(
                self._resolve_paragraph(
                    block,
                    style_resolver=style_resolver,
                    default_style_id=default_style_id,
                    numbering=numbering,
                )
            )
            if isinstance(block, ParagraphModel)
            else self._resolve_table(
                block,
                style_resolver=style_resolver,
                default_style_id=default_style_id,
                numbering=numbering,
            )
            for block in model.blocks
        )
        return model.model_copy(update={"blocks": blocks})

    @staticmethod
    def _resolved_to_raw(paragraph: ResolvedParagraphModel) -> ParagraphModel:
        properties = ParagraphProperties(
            alignment=paragraph.alignment,
            left_indent=paragraph.left_indent,
            right_indent=paragraph.right_indent,
            first_line_indent=paragraph.first_line_indent,
            hanging_indent=paragraph.hanging_indent,
            space_before=paragraph.space_before,
            space_after=paragraph.space_after,
            line_spacing=paragraph.line_spacing,
            line_spacing_rule=paragraph.line_spacing_rule,
            keep_next=paragraph.keep_next,
            keep_lines=paragraph.keep_lines,
            page_break_before=paragraph.page_break_before,
            widow_control=paragraph.widow_control,
            tabs=paragraph.tabs,
        )
        runs = tuple(
            RunModel(
                text=run.text,
                properties=RunProperties(
                    font_family=run.font_name,
                    font_path=run.font_path,
                    font_size=run.font_size,
                    bold=run.bold,
                    italic=run.italic,
                    underline=run.underline,
                    strike=run.strike,
                    color=run.color,
                    highlight=run.highlight,
                    vertical_align=run.vertical_align,
                    character_spacing=run.character_spacing,
                    hidden=run.hidden,
                ),
                break_type=run.break_type,
                tab=run.tab,
                image=run.image,
                hidden=run.hidden,
                source_index=run.source_index,
            )
            for run in paragraph.runs
        )
        return ParagraphModel(runs=runs, properties=properties, source_index=paragraph.source_index)

    def _numbering_label(
        self,
        properties: ParagraphProperties,
        numbering: tuple[NumberingDefinition, ...],
    ) -> str | None:
        if properties.numbering_id is None:
            return None
        level = properties.numbering_level or 0
        definition = next(
            (item for item in numbering if item.numbering_id == properties.numbering_id),
            None,
        )
        level_model = (
            next(
                (item for item in definition.levels if item.level == level),
                None,
            )
            if definition is not None
            else None
        )
        if level_model is None:
            return "•"
        key = (properties.numbering_id, level)
        value = self._number_counters.get(key, level_model.start - 1) + 1
        self._number_counters[key] = value
        if level_model.number_format == "bullet":
            return level_model.text
        return level_model.text.replace(f"%{level + 1}", str(value))

    @staticmethod
    def _requested_font(properties: RunProperties, *, east_asia: bool) -> str:
        if east_asia:
            return (
                properties.east_asia_font
                or properties.font_family
                or properties.high_ansi_font
                or properties.ascii_font
                or "Helvetica"
            )
        return (
            properties.ascii_font
            or properties.high_ansi_font
            or properties.font_family
            or properties.east_asia_font
            or "Helvetica"
        )

    @staticmethod
    def _contains_east_asia(text: str) -> bool:
        for character in text:
            codepoint = ord(character)
            if (
                0x3040 <= codepoint <= 0x30FF
                or 0x3400 <= codepoint <= 0x9FFF
                or 0xF900 <= codepoint <= 0xFAFF
                or unicodedata.east_asian_width(character) in {"W", "F"}
            ):
                return True
        return False

    @classmethod
    def _script_segments(cls, text: str) -> tuple[tuple[str, bool], ...]:
        if not text:
            return (("", False),)
        result: list[tuple[str, bool]] = []
        current = ""
        current_east_asia = False
        for character in text:
            if current and cls._extends_previous_cluster(character, current):
                current += character
                continue
            east_asia = cls._contains_east_asia(character)
            if not current:
                current = character
                current_east_asia = east_asia
                continue
            if east_asia == current_east_asia:
                current += character
                continue
            result.append((current, current_east_asia))
            current = character
            current_east_asia = east_asia
        result.append((current, current_east_asia))
        return tuple(result)

    @staticmethod
    def _extends_previous_cluster(character: str, current: str) -> bool:
        codepoint = ord(character)
        return (
            current.endswith("\u200d")
            or character == "\u200d"
            or unicodedata.combining(character) != 0
            or 0xFE00 <= codepoint <= 0xFE0F
            or 0xE0100 <= codepoint <= 0xE01EF
            or 0x1F3FB <= codepoint <= 0x1F3FF
        )
