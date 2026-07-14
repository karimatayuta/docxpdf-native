from __future__ import annotations

from pathlib import Path
from typing import Literal

from pydantic import Field

from docxpdf_native.models.base import FrozenModel

ParagraphAlignment = Literal["left", "center", "right", "justify", "distribute"]
LineSpacingRule = Literal["auto", "at_least", "exact"]
StyleType = Literal["paragraph", "character", "table", "numbering"]


class TabStop(FrozenModel):
    position: float
    alignment: Literal["left", "center", "right", "decimal", "bar", "clear", "num"] = "left"
    leader: Literal["none", "dot", "hyphen", "underscore", "middle_dot", "heavy"] = "none"


class RunProperties(FrozenModel):
    font_family: str | None = None
    font_path: Path | None = None
    east_asia_font: str | None = None
    ascii_font: str | None = None
    high_ansi_font: str | None = None
    complex_script_font: str | None = None
    font_size: float | None = Field(default=None, gt=0)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | str | None = None
    strike: bool | None = None
    color: str | None = None
    highlight: str | None = None
    vertical_align: Literal["baseline", "superscript", "subscript"] | None = None
    character_spacing: float | None = None
    hidden: bool | None = None


class ParagraphProperties(FrozenModel):
    alignment: ParagraphAlignment | None = None
    left_indent: float | None = None
    right_indent: float | None = None
    first_line_indent: float | None = None
    hanging_indent: float | None = None
    space_before: float | None = None
    space_after: float | None = None
    line_spacing: float | None = Field(default=None, gt=0)
    line_spacing_rule: LineSpacingRule | None = None
    keep_next: bool | None = None
    keep_lines: bool | None = None
    page_break_before: bool | None = None
    widow_control: bool | None = None
    tabs: tuple[TabStop, ...] = ()
    numbering_id: int | None = None
    numbering_level: int | None = Field(default=None, ge=0)


class StyleModel(FrozenModel):
    style_id: str
    style_type: StyleType
    name: str | None = None
    based_on: str | None = None
    next_style: str | None = None
    link: str | None = None
    is_default: bool = False
    paragraph: ParagraphProperties = Field(default_factory=ParagraphProperties)
    run: RunProperties = Field(default_factory=RunProperties)


class ThemeFonts(FrozenModel):
    major_latin: str | None = None
    major_east_asia: str | None = None
    major_complex_script: str | None = None
    minor_latin: str | None = None
    minor_east_asia: str | None = None
    minor_complex_script: str | None = None


class DocumentDefaults(FrozenModel):
    paragraph: ParagraphProperties = Field(default_factory=ParagraphProperties)
    run: RunProperties = Field(default_factory=RunProperties)
