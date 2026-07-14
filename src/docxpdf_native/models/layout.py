from __future__ import annotations

from pathlib import Path
from typing import Literal

from pydantic import Field

from docxpdf_native.models.base import FrozenModel
from docxpdf_native.models.document import TableBorders
from docxpdf_native.models.results import ConversionWarning


class BlockBox(FrozenModel):
    kind: str = "block"
    x: float
    y: float
    width: float = Field(ge=0)
    height: float = Field(ge=0)
    source_start: int = Field(default=0, ge=0)
    source_end: int = Field(default=0, ge=0)


class TextFragment(FrozenModel):
    kind: Literal["text"] = "text"
    x: float
    y: float
    width: float = Field(ge=0)
    height: float = Field(ge=0)
    text: str
    font_name: str
    font_path: Path | None = None
    font_size: float = Field(gt=0)
    bold: bool = False
    italic: bool = False
    underline: bool | str = False
    strike: bool = False
    color: str = "000000"
    highlight: str | None = None
    character_spacing: float = 0
    baseline_shift: float = 0
    source_start: int = Field(default=0, ge=0)
    source_end: int = Field(default=0, ge=0)


class ImageBox(BlockBox):
    kind: Literal["image"] = "image"
    image_data: bytes = Field(default=b"", exclude=True, repr=False)
    content_type: Literal["image/png", "image/jpeg"]
    part_name: str | None = None
    relationship_id: str | None = None


class LineBox(BlockBox):
    kind: Literal["line"] = "line"
    used_width: float = Field(default=0, ge=0)
    ascent: float = Field(default=0, ge=0)
    descent: float = Field(default=0, ge=0)
    baseline: float = Field(default=0, ge=0)
    fragments: tuple[TextFragment | ImageBox, ...] = ()


class ParagraphBox(BlockBox):
    kind: Literal["paragraph"] = "paragraph"
    lines: tuple[LineBox, ...] = ()
    paragraph_index: int = Field(default=0, ge=0)
    continued_from_previous_page: bool = False
    continues_on_next_page: bool = False


class CellBox(BlockBox):
    kind: Literal["cell"] = "cell"
    row_index: int = Field(ge=0)
    column_index: int = Field(ge=0)
    row_span: int = Field(default=1, ge=1)
    column_span: int = Field(default=1, ge=1)
    blocks: tuple[ParagraphBox | ImageBox, ...] = ()
    background_color: str | None = None
    borders: TableBorders = Field(default_factory=TableBorders)
    vertical_alignment: Literal["top", "center", "bottom"] = "top"


class TableBox(BlockBox):
    kind: Literal["table"] = "table"
    cells: tuple[CellBox, ...] = ()
    column_widths: tuple[float, ...] = ()
    row_heights: tuple[float, ...] = ()
    table_index: int = Field(default=0, ge=0)
    continued_from_previous_page: bool = False
    continues_on_next_page: bool = False
    borders: TableBorders = Field(default_factory=TableBorders)


class PageRegion(FrozenModel):
    x: float
    y: float
    width: float = Field(ge=0)
    height: float = Field(ge=0)


class PageModel(FrozenModel):
    number: int = Field(ge=1)
    width: float = Field(gt=0)
    height: float = Field(gt=0)
    section_index: int = Field(default=0, ge=0)
    section_page_number: int | None = Field(default=None, ge=0)
    header_region: PageRegion | None = None
    body_region: PageRegion | None = None
    footer_region: PageRegion | None = None
    header: tuple[ParagraphBox | TableBox | ImageBox | BlockBox, ...] = ()
    body: tuple[ParagraphBox | TableBox | ImageBox | BlockBox, ...] = ()
    footer: tuple[ParagraphBox | TableBox | ImageBox | BlockBox, ...] = ()
    source_start: int = Field(default=0, ge=0)
    source_end: int = Field(default=0, ge=0)


class LayoutDocument(FrozenModel):
    pages: tuple[PageModel, ...] = ()
    warnings: tuple[ConversionWarning, ...] = ()

    @property
    def page_count(self) -> int:
        return len(self.pages)
