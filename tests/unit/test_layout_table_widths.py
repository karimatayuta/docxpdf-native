from __future__ import annotations

import pytest

from docxpdf_native.abstractions import TextMeasurer
from docxpdf_native.exceptions import LayoutError
from docxpdf_native.layout.engine import NativeLayoutEngine
from docxpdf_native.models import (
    ConversionOptions,
    ParagraphModel,
    ResolvedDocumentModel,
    ResolvedFont,
    ResolvedSectionModel,
    RunModel,
    TableCellModel,
    TableModel,
    TableRowModel,
    TextMeasurement,
)


class _FixedTextMeasurer(TextMeasurer):
    def measure(
        self,
        text: str,
        font: ResolvedFont,
        font_size: float,
        *,
        character_spacing: float = 0,
    ) -> TextMeasurement:
        return TextMeasurement(width=len(text) * 5, ascent=8, descent=2)


def _layout(table: TableModel, *, page_height: float = 200):
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=220,
                page_height=page_height,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )
    return NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())


def test_table_width_controls_equal_columns_when_grid_is_absent() -> None:
    table = TableModel(
        width=120,
        rows=(
            TableRowModel(
                cells=(TableCellModel(), TableCellModel()),
            ),
        ),
    )

    layout = _layout(table)

    table_box = layout.pages[0].body[0]
    assert table_box.column_widths == (60.0, 60.0)  # type: ignore[union-attr]


def test_cell_widths_control_columns_when_grid_is_absent() -> None:
    table = TableModel(
        rows=(
            TableRowModel(
                cells=(TableCellModel(width=40), TableCellModel(width=80)),
            ),
        ),
    )

    layout = _layout(table)

    table_box = layout.pages[0].body[0]
    assert table_box.column_widths == (40.0, 80.0)  # type: ignore[union-attr]


def test_cell_text_alignment_overrides_paragraph_alignment() -> None:
    table = TableModel(
        grid_widths=(100,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(runs=(RunModel(text="x"),)),),
                        text_alignment="right",
                    ),
                ),
            ),
        ),
    )

    layout = _layout(table)

    table_box = layout.pages[0].body[0]
    cell = table_box.cells[0]  # type: ignore[union-attr]
    fragment = cell.blocks[0].lines[0].fragments[0]  # type: ignore[union-attr]
    assert fragment.x == cell.x + cell.width - fragment.width


def test_vertical_merge_crossing_page_boundary_is_explicitly_rejected() -> None:
    table = TableModel(
        grid_widths=(100,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(),),
                        vertical_merge="restart",
                    ),
                ),
                height=30,
                height_rule="exact",
            ),
            TableRowModel(
                cells=(TableCellModel(vertical_merge="continue"),),
                height=30,
                height_rule="exact",
            ),
        ),
    )

    with pytest.raises(LayoutError, match="vertically merged cell"):
        _layout(table, page_height=50)
