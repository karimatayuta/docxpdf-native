from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Literal

from pydantic import Field, model_validator

from docxpdf_native.models.base import FrozenModel
from docxpdf_native.models.fonts import FontSubstitution
from docxpdf_native.models.styles import (
    DocumentDefaults,
    ParagraphAlignment,
    ParagraphProperties,
    RunProperties,
    StyleModel,
    TabStop,
    ThemeFonts,
)


class ImageModel(FrozenModel):
    relationship_id: str
    part_name: str
    content_type: Literal["image/png", "image/jpeg"]
    data: bytes = Field(repr=False)
    width: float = Field(gt=0)
    height: float = Field(gt=0)
    description: str | None = None
    source_index: int = Field(default=0, ge=0)


class PlaceholderModel(FrozenModel):
    """A same-size stand-in for content that cannot be rendered natively.

    Used for embedded objects, EMF/WMF previews, unrecognized drawings, and
    other constructs whose exact appearance is out of scope but whose page
    footprint must still be reserved so pagination matches the source DOCX.
    """

    label: str
    width: float = Field(gt=0)
    height: float = Field(gt=0)
    source_index: int = Field(default=0, ge=0)


class RunModel(FrozenModel):
    text: str = ""
    style_id: str | None = None
    properties: RunProperties = Field(default_factory=RunProperties)
    break_type: Literal["line", "page"] | None = None
    tab: bool = False
    image: ImageModel | None = None
    placeholder: PlaceholderModel | None = None
    preserve_space: bool = False
    hidden: bool = False
    source_index: int = Field(default=0, ge=0)

    @model_validator(mode="after")
    def validate_content_kind(self) -> RunModel:
        content_kinds = sum(
            (
                bool(self.text),
                self.break_type is not None,
                self.tab,
                self.image is not None,
                self.placeholder is not None,
            )
        )
        if content_kinds > 1:
            raise ValueError(
                "a run may contain only one of text, break, tab, image, or placeholder"
            )
        return self


class ParagraphModel(FrozenModel):
    runs: tuple[RunModel, ...] = ()
    style_id: str | None = None
    properties: ParagraphProperties = Field(default_factory=ParagraphProperties)
    source_index: int = Field(default=0, ge=0)

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)


class CellMargins(FrozenModel):
    top: float = Field(default=0, ge=0)
    right: float = Field(default=0, ge=0)
    bottom: float = Field(default=0, ge=0)
    left: float = Field(default=0, ge=0)


class BorderModel(FrozenModel):
    style: str = "single"
    width: float = Field(default=0.5, ge=0)
    color: str = "000000"


class TableBorders(FrozenModel):
    top: BorderModel | None = None
    right: BorderModel | None = None
    bottom: BorderModel | None = None
    left: BorderModel | None = None
    inside_horizontal: BorderModel | None = None
    inside_vertical: BorderModel | None = None


class TableCellModel(FrozenModel):
    paragraphs: tuple[ParagraphModel, ...] = ()
    width: float | None = Field(default=None, gt=0)
    grid_span: int = Field(default=1, ge=1)
    vertical_merge: Literal["restart", "continue"] | None = None
    margins: CellMargins = Field(default_factory=CellMargins)
    borders: TableBorders = Field(default_factory=TableBorders)
    background_color: str | None = None
    vertical_alignment: Literal["top", "center", "bottom"] = "top"
    text_alignment: ParagraphAlignment | None = None
    source_index: int = Field(default=0, ge=0)


class TableRowModel(FrozenModel):
    cells: tuple[TableCellModel, ...] = ()
    height: float | None = Field(default=None, gt=0)
    height_rule: Literal["auto", "at_least", "exact"] = "auto"
    cant_split: bool = False
    repeat_header: bool = False
    source_index: int = Field(default=0, ge=0)


class TableModel(FrozenModel):
    rows: tuple[TableRowModel, ...] = ()
    grid_widths: tuple[float, ...] = ()
    width: float | None = Field(default=None, gt=0)
    style_id: str | None = None
    autofit: bool = False
    alignment: Literal["left", "center", "right"] = "left"
    left_indent: float = 0
    cell_margins: CellMargins = Field(default_factory=CellMargins)
    borders: TableBorders = Field(default_factory=TableBorders)
    source_index: int = Field(default=0, ge=0)


class HeaderFooterModel(FrozenModel):
    kind: Literal["default", "first", "even"] = "default"
    relationship_id: str | None = None
    part_name: str | None = None
    blocks: tuple[ParagraphModel | TableModel, ...] = ()


class NumberingLevel(FrozenModel):
    level: int = Field(ge=0)
    number_format: str = "decimal"
    text: str = "%1."
    start: int = 1
    paragraph: ParagraphProperties = Field(default_factory=ParagraphProperties)
    run: RunProperties = Field(default_factory=RunProperties)


class NumberingDefinition(FrozenModel):
    numbering_id: int
    abstract_numbering_id: int | None = None
    levels: tuple[NumberingLevel, ...] = ()


class DocumentMetadata(FrozenModel):
    title: str | None = None
    subject: str | None = None
    creator: str | None = None
    keywords: str | None = None
    description: str | None = None
    application: str | None = None
    custom: Mapping[str, str] = Field(default_factory=dict)


class SectionModel(FrozenModel):
    blocks: tuple[ParagraphModel | TableModel, ...] = ()
    page_width: float = Field(default=595.28, gt=0)
    page_height: float = Field(default=841.89, gt=0)
    orientation: Literal["portrait", "landscape"] = "portrait"
    margin_top: float = Field(default=72, ge=0)
    margin_right: float = Field(default=72, ge=0)
    margin_bottom: float = Field(default=72, ge=0)
    margin_left: float = Field(default=72, ge=0)
    header_distance: float = Field(default=36, ge=0)
    footer_distance: float = Field(default=36, ge=0)
    page_number_start: int | None = Field(default=None, ge=0)
    section_break: Literal["continuous", "next_page", "even_page", "odd_page"] = "next_page"
    headers: tuple[HeaderFooterModel, ...] = ()
    footers: tuple[HeaderFooterModel, ...] = ()
    columns: int = Field(default=1, ge=1)
    doc_grid_type: Literal["default", "lines", "linesAndChars", "snapToChars"] = "default"
    doc_grid_line_pitch: float | None = Field(default=None, gt=0)
    source_index: int = Field(default=0, ge=0)


class DocumentModel(FrozenModel):
    sections: tuple[SectionModel, ...] = ()
    styles: tuple[StyleModel, ...] = ()
    defaults: DocumentDefaults = Field(default_factory=DocumentDefaults)
    theme_fonts: ThemeFonts = Field(default_factory=ThemeFonts)
    numbering: tuple[NumberingDefinition, ...] = ()
    metadata: DocumentMetadata = Field(default_factory=DocumentMetadata)
    source_name: str | None = None


class ResolvedRunModel(FrozenModel):
    text: str = ""
    font_name: str = "Helvetica"
    font_path: Path | None = None
    font_size: float = Field(default=11, gt=0)
    bold: bool = False
    italic: bool = False
    underline: bool | str = False
    strike: bool = False
    color: str = "000000"
    highlight: str | None = None
    vertical_align: Literal["baseline", "superscript", "subscript"] = "baseline"
    character_spacing: float = 0
    break_type: Literal["line", "page"] | None = None
    tab: bool = False
    image: ImageModel | None = None
    placeholder: PlaceholderModel | None = None
    hidden: bool = False
    source_index: int = Field(default=0, ge=0)


class ResolvedParagraphModel(FrozenModel):
    runs: tuple[ResolvedRunModel, ...] = ()
    alignment: ParagraphAlignment = "left"
    left_indent: float = 0
    right_indent: float = 0
    first_line_indent: float = 0
    hanging_indent: float = 0
    space_before: float = 0
    space_after: float = 0
    line_spacing: float = Field(default=1, gt=0)
    line_spacing_rule: Literal["auto", "at_least", "exact"] = "auto"
    keep_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    widow_control: bool = True
    tabs: tuple[TabStop, ...] = ()
    numbering_label: str | None = None
    source_index: int = Field(default=0, ge=0)


class ResolvedSectionModel(FrozenModel):
    blocks: tuple[ResolvedParagraphModel | TableModel, ...] = ()
    page_width: float = Field(default=595.28, gt=0)
    page_height: float = Field(default=841.89, gt=0)
    orientation: Literal["portrait", "landscape"] = "portrait"
    margin_top: float = Field(default=72, ge=0)
    margin_right: float = Field(default=72, ge=0)
    margin_bottom: float = Field(default=72, ge=0)
    margin_left: float = Field(default=72, ge=0)
    header_distance: float = Field(default=36, ge=0)
    footer_distance: float = Field(default=36, ge=0)
    page_number_start: int | None = Field(default=None, ge=0)
    section_break: Literal["continuous", "next_page", "even_page", "odd_page"] = "next_page"
    headers: tuple[HeaderFooterModel, ...] = ()
    footers: tuple[HeaderFooterModel, ...] = ()
    doc_grid_type: Literal["default", "lines", "linesAndChars", "snapToChars"] = "default"
    doc_grid_line_pitch: float | None = Field(default=None, gt=0)
    source_index: int = Field(default=0, ge=0)


class ResolvedDocumentModel(FrozenModel):
    sections: tuple[ResolvedSectionModel | SectionModel, ...] = ()
    paragraphs: tuple[ResolvedParagraphModel, ...] = ()
    source: DocumentModel | None = None
    font_substitutions: tuple[FontSubstitution, ...] = ()
