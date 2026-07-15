from __future__ import annotations

import pytest

from docxpdf_native.abstractions import TextMeasurer
from docxpdf_native.exceptions import LayoutError, PageLimitExceededError
from docxpdf_native.layout.engine import NativeLayoutEngine
from docxpdf_native.models import (
    BorderModel,
    ConversionOptions,
    HeaderFooterModel,
    ImageModel,
    ParagraphModel,
    ParagraphProperties,
    PlaceholderModel,
    ResolvedDocumentModel,
    ResolvedFont,
    ResolvedParagraphModel,
    ResolvedRunModel,
    ResolvedSectionModel,
    ResourceLimits,
    RunModel,
    SectionModel,
    TableBorders,
    TableCellModel,
    TableModel,
    TableRowModel,
    TabStop,
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
        width = len(text) * font_size * 0.5
        if text:
            width += max(0, len(text) - 1) * character_spacing
        return TextMeasurement(width=width, ascent=font_size * 0.8, descent=font_size * 0.2)


def test_layout_engine_places_short_paragraph_on_one_page() -> None:
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(
            ResolvedParagraphModel(
                runs=(ResolvedRunModel(text="hello", font_name="Helvetica", font_size=12),),
                source_index=0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 1


def test_layout_engine_preserves_paragraph_text_in_fragment() -> None:
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(
            ResolvedParagraphModel(
                runs=(ResolvedRunModel(text="hello", font_name="Helvetica", font_size=12),),
                source_index=0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    paragraph = layout.pages[0].body[0]
    assert paragraph.lines[0].fragments[0].text == "hello"  # type: ignore[union-attr]


def test_layout_engine_applies_super_and_subscript_baseline_shifts() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(
                text="super",
                font_name="Helvetica",
                font_size=12,
                vertical_align="superscript",
            ),
            ResolvedRunModel(
                text="sub",
                font_name="Helvetica",
                font_size=12,
                vertical_align="subscript",
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    fragments = layout.pages[0].body[0].lines[0].fragments  # type: ignore[union-attr]
    assert fragments[0].baseline_shift > 0  # type: ignore[union-attr]
    assert fragments[1].baseline_shift < 0  # type: ignore[union-attr]


def test_layout_engine_wraps_text_to_available_width() -> None:
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=80,
                page_height=200,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
        paragraphs=(
            ResolvedParagraphModel(
                runs=(ResolvedRunModel(text="one two three", font_name="Helvetica", font_size=12),),
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    paragraph = layout.pages[0].body[0]
    assert tuple(line.fragments[0].text for line in paragraph.lines) == (  # type: ignore[union-attr]
        "one two ",
        "three",
    )


def test_layout_engine_stacks_paragraphs_without_overlap() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_after=6,
    )
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(paragraph, paragraph.model_copy(update={"source_index": 1})),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    first, second = layout.pages[0].body
    assert second.y >= first.y + first.height


def test_layout_engine_collapses_paragraph_spacing_when_after_exceeds_before() -> None:
    # Word does not sum a paragraph's space_after with the next paragraph's
    # space_before -- only the larger of the two applies. Here after (20) >
    # before (10), so the second paragraph's own space_before contributes
    # nothing extra: its text starts right where the first paragraph ends.
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_after=20,
        source_index=0,
    )
    second = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_before=10,
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(first, second),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    first_box, second_box = layout.pages[0].body
    # first_box.height = 0 (before) + 12 (line) + 20 (after)
    assert first_box.height == 32.0  # type: ignore[union-attr]
    assert second_box.lines[0].y == first_box.y + first_box.height  # type: ignore[union-attr]


def test_layout_engine_uses_larger_space_before_when_it_exceeds_previous_after() -> None:
    # Here before (15) > after (5), so only the uncovered remainder (10)
    # of the second paragraph's space_before is added after the first
    # paragraph's box (which already contributes its own space_after).
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_after=5,
        source_index=0,
    )
    second = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_before=15,
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(first, second),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    first_box, second_box = layout.pages[0].body
    assert second_box.lines[0].y == first_box.y + first_box.height + 10  # type: ignore[union-attr]


def test_layout_engine_resets_paragraph_spacing_collapse_across_a_table() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_after=20,
        source_index=0,
    )
    table = TableModel(
        grid_widths=(80,),
        rows=(TableRowModel(cells=(TableCellModel(),)),),
        source_index=1,
    )
    second = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line", font_name="Helvetica", font_size=12),),
        space_before=10,
        source_index=2,
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(first, table, second)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    _, table_box, second_box = layout.pages[0].body
    # The table does not participate in spacing collapse, so the paragraph
    # after it keeps its full space_before.
    assert second_box.lines[0].y == table_box.y + table_box.height + 10  # type: ignore[union-attr]


def test_layout_engine_collapses_paragraph_spacing_in_table_cells() -> None:
    first = ParagraphModel(
        runs=(RunModel(text="line"),),
        properties=ParagraphProperties(space_after=20),
    )
    second = ParagraphModel(
        runs=(RunModel(text="line"),),
        properties=ParagraphProperties(space_before=10),
    )
    table = TableModel(
        grid_widths=(80,),
        rows=(TableRowModel(cells=(TableCellModel(paragraphs=(first, second)),)),),
    )
    document = ResolvedDocumentModel(sections=(ResolvedSectionModel(blocks=(table,)),))

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    cell = layout.pages[0].body[0].cells[0]  # type: ignore[union-attr]
    first_box, second_box = cell.blocks
    assert second_box.lines[0].y == first_box.y + first_box.height  # type: ignore[union-attr]


def test_layout_engine_paginates_complete_paragraphs() -> None:
    paragraphs = tuple(
        ResolvedParagraphModel(
            runs=(ResolvedRunModel(text=f"line {index}", font_name="Helvetica", font_size=12),),
            source_index=index,
        )
        for index in range(4)
    )
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=120,
                page_height=40,
                margin_left=5,
                margin_right=5,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
        paragraphs=paragraphs,
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2


def test_layout_engine_honors_page_break_before() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first", font_name="Helvetica", font_size=12),),
        source_index=0,
    )
    second = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="second", font_name="Helvetica", font_size=12),),
        page_break_before=True,
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(first, second),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2


def test_layout_engine_honors_explicit_page_break_run() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="before", font_name="Helvetica", font_size=12),
            ResolvedRunModel(break_type="page"),
            ResolvedRunModel(text="after", font_name="Helvetica", font_size=12),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(paragraph,),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2


def test_layout_engine_honors_explicit_line_break_run() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="before", font_name="Helvetica", font_size=12),
            ResolvedRunModel(break_type="line"),
            ResolvedRunModel(text="after", font_name="Helvetica", font_size=12),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(SectionModel(),),
        paragraphs=(paragraph,),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    box = layout.pages[0].body[0]
    assert tuple("".join(fragment.text for fragment in line.fragments) for line in box.lines) == (  # type: ignore[union-attr]
        "before",
        "after",
    )


def test_layout_engine_splits_long_paragraph_between_pages() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(
                text="one two three four five six seven eight nine ten",
                font_name="Helvetica",
                font_size=12,
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=80,
                page_height=40,
                margin_left=10,
                margin_right=10,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
        paragraphs=(paragraph,),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 3


def test_layout_engine_moves_keep_lines_paragraph_to_next_page() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first", font_name="Helvetica", font_size=12),),
        space_after=6,
        source_index=0,
    )
    kept = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three", font_name="Helvetica", font_size=12),),
        keep_lines=True,
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=80,
                page_height=46,
                margin_left=10,
                margin_right=10,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
        paragraphs=(first, kept),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(box.paragraph_index for box in layout.pages[0].body) == (0,)  # type: ignore[union-attr]


def test_layout_engine_keeps_keep_next_pair_together() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first", font_name="Helvetica", font_size=12),),
        space_after=12,
        source_index=0,
    )
    heading = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="heading", font_name="Helvetica", font_size=12),),
        keep_next=True,
        source_index=1,
    )
    following = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="following", font_name="Helvetica", font_size=12),),
        source_index=2,
    )
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=100,
                page_height=46,
                margin_left=5,
                margin_right=5,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
        paragraphs=(first, heading, following),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(box.paragraph_index for box in layout.pages[0].body) == (0,)  # type: ignore[union-attr]


def test_layout_engine_avoids_single_widow_line_at_page_bottom() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first", font_name="Helvetica", font_size=12),),
        space_after=12,
        source_index=0,
    )
    wrapped = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(
                text="one two three four five six",
                font_name="Helvetica",
                font_size=12,
            ),
        ),
        widow_control=True,
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=80,
                page_height=47,
                margin_left=10,
                margin_right=10,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
        paragraphs=(first, wrapped),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(box.paragraph_index for box in layout.pages[0].body) == (0,)  # type: ignore[union-attr]


def test_layout_engine_starts_section_with_its_page_size() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first", font_name="Helvetica", font_size=12),),
        source_index=0,
    )
    second = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="second", font_name="Helvetica", font_size=12),),
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(blocks=(first,), page_width=200, page_height=300),
            ResolvedSectionModel(blocks=(second,), page_width=300, page_height=200),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple((page.width, page.height) for page in layout.pages) == (
        (200.0, 300.0),
        (300.0, 200.0),
    )


def test_layout_engine_uses_fixed_table_grid_widths() -> None:
    table = TableModel(
        grid_widths=(40, 60),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text="A"),)),)),
                    TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text="B"),)),)),
                ),
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=200,
                page_height=300,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    table_box = layout.pages[0].body[0]
    assert table_box.column_widths == (40.0, 60.0)  # type: ignore[union-attr]


def test_layout_engine_reports_deterministic_autofit_fallback() -> None:
    table = TableModel(
        autofit=True,
        rows=(
            TableRowModel(
                cells=(TableCellModel(paragraphs=(ParagraphModel(),)),),
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(table,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert [warning.code for warning in layout.warnings] == ["table-autofit-fallback"]


def test_layout_engine_preserves_mixed_paragraph_table_order() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="before", font_name="Helvetica", font_size=12),),
        source_index=0,
    )
    table = TableModel(
        grid_widths=(80,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(runs=(RunModel(text="cell"),)),),
                    ),
                ),
            ),
        ),
        source_index=1,
    )
    last = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="after", font_name="Helvetica", font_size=12),),
        source_index=2,
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(first, table, last)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert [block.kind for block in layout.pages[0].body] == [
        "paragraph",
        "table",
        "paragraph",
    ]


def test_layout_engine_paginates_table_by_rows() -> None:
    rows = tuple(
        TableRowModel(
            cells=(
                TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text=f"row {index}"),)),)),
            ),
            height=20,
            height_rule="exact",
            cant_split=True,
        )
        for index in range(5)
    )
    table = TableModel(grid_widths=(100,), rows=rows)
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=70,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 3


def test_layout_engine_repeats_table_header_row_on_each_page() -> None:
    header = TableRowModel(
        cells=(TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text="Header"),)),)),),
        height=10,
        height_rule="exact",
        repeat_header=True,
    )
    rows = tuple(
        TableRowModel(
            cells=(
                TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text=f"row {index}"),)),)),
            ),
            height=20,
            height_rule="exact",
        )
        for index in range(4)
    )
    table = TableModel(grid_widths=(100,), rows=(header, *rows))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=70,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert all(any(cell.row_index == 0 for cell in page.body[0].cells) for page in layout.pages)  # type: ignore[union-attr]


def test_layout_engine_represents_vertical_cell_merge_as_row_span() -> None:
    table = TableModel(
        grid_widths=(50,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(runs=(RunModel(text="merged"),)),),
                        vertical_merge="restart",
                    ),
                ),
            ),
            TableRowModel(cells=(TableCellModel(vertical_merge="continue"),)),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=100,
                page_height=200,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    table_box = layout.pages[0].body[0]
    merged = next(cell for cell in table_box.cells if cell.row_index == 0)  # type: ignore[union-attr]
    assert merged.row_span == 2


def test_layout_engine_preserves_cell_border_and_background() -> None:
    borders = TableBorders(top=BorderModel(width=2, color="FF0000"))
    table = TableModel(
        grid_widths=(50,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        background_color="FFFF00",
                        borders=borders,
                    ),
                ),
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(table,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    cell = layout.pages[0].body[0].cells[0]  # type: ignore[union-attr]
    assert (cell.background_color, cell.borders.top.width) == ("FFFF00", 2.0)


def test_layout_engine_bottom_aligns_cell_content() -> None:
    table = TableModel(
        grid_widths=(50,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(runs=(RunModel(text="bottom"),)),),
                        vertical_alignment="bottom",
                    ),
                ),
                height=40,
                height_rule="exact",
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(table,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    cell = layout.pages[0].body[0].cells[0]  # type: ignore[union-attr]
    assert cell.blocks[0].y + cell.blocks[0].height == cell.y + cell.height


def test_layout_engine_scales_inline_image_to_body_width() -> None:
    image = ImageModel(
        relationship_id="rId1",
        part_name="word/media/image.png",
        content_type="image/png",
        data=b"png",
        width=100,
        height=50,
    )
    paragraph = ResolvedParagraphModel(runs=(ResolvedRunModel(image=image),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                page_height=100,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    fragment = layout.pages[0].body[0].lines[0].fragments[0]  # type: ignore[union-attr]
    assert (fragment.width, fragment.height) == (60.0, 30.0)


def test_layout_engine_scales_placeholder_to_body_width_and_preserves_label() -> None:
    placeholder = PlaceholderModel(label="[Embedded Object]", width=100, height=50)
    paragraph = ResolvedParagraphModel(runs=(ResolvedRunModel(placeholder=placeholder),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                page_height=100,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    fragment = layout.pages[0].body[0].lines[0].fragments[0]  # type: ignore[union-attr]
    assert (fragment.width, fragment.height) == (60.0, 30.0)
    assert fragment.label == "[Embedded Object]"  # type: ignore[union-attr]


def test_layout_engine_advances_tab_to_next_default_stop() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="A", font_name="Helvetica", font_size=12),
            ResolvedRunModel(tab=True),
            ResolvedRunModel(text="B", font_name="Helvetica", font_size=12),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    fragments = layout.pages[0].body[0].lines[0].fragments  # type: ignore[union-attr]
    assert fragments[1].x - fragments[0].x == 36.0


def test_layout_engine_resolves_page_and_total_page_fields() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="{{PAGE}}/{{NUMPAGES}}", font_name="Helvetica"),),
        source_index=0,
    )
    second = first.model_copy(update={"page_break_before": True, "source_index": 1})
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(first, second)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(page.body[0].lines[0].fragments[0].text for page in layout.pages) == (  # type: ignore[union-attr]
        "1/2",
        "2/2",
    )


def test_layout_engine_places_header_and_footer_content() -> None:
    header = HeaderFooterModel(
        blocks=(ParagraphModel(runs=(RunModel(text="Header"),)),),
    )
    footer = HeaderFooterModel(
        blocks=(ParagraphModel(runs=(RunModel(text="Page {{PAGE}}"),)),),
    )
    body = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="Body", font_name="Helvetica"),),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(body,), headers=(header,), footers=(footer,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    page = layout.pages[0]
    assert (
        page.header[0].lines[0].fragments[0].text,  # type: ignore[union-attr]
        page.footer[0].lines[0].fragments[0].text,  # type: ignore[union-attr]
    ) == ("Header", "Page 1")


def test_layout_engine_gives_empty_paragraph_a_line_height() -> None:
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(ResolvedParagraphModel(),)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].height > 0


def test_layout_engine_places_numbering_label_before_text() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="item", font_name="Helvetica"),),
        numbering_label="1.",
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    text = "".join(
        fragment.text
        for fragment in layout.pages[0].body[0].lines[0].fragments  # type: ignore[union-attr]
    )
    assert text == "1. item"


def test_layout_engine_applies_hanging_indent_after_first_line() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three", font_name="Helvetica", font_size=12),),
        left_indent=20,
        hanging_indent=10,
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=100,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    lines = layout.pages[0].body[0].lines  # type: ignore[union-attr]
    assert lines[1].x - lines[0].x == 10.0


def test_layout_engine_justifies_nonfinal_line_to_available_width() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three", font_name="Helvetica", font_size=12),),
        alignment="justify",
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    line = layout.pages[0].body[0].lines[0]  # type: ignore[union-attr]
    assert line.used_width == line.width


def test_layout_engine_enforces_page_limit_after_section_layout() -> None:
    paragraph = ResolvedParagraphModel(runs=(ResolvedRunModel(text="page"),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(blocks=(paragraph,)),
            ResolvedSectionModel(blocks=(paragraph,)),
        ),
    )
    options = ConversionOptions(resource_limits=ResourceLimits(max_pages=1))

    with pytest.raises(PageLimitExceededError, match="page limit exceeded"):
        NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=options)


def test_layout_engine_selects_first_even_and_default_headers() -> None:
    headers = (
        HeaderFooterModel(
            kind="default",
            blocks=(ParagraphModel(runs=(RunModel(text="default"),)),),
        ),
        HeaderFooterModel(
            kind="first",
            blocks=(ParagraphModel(runs=(RunModel(text="first"),)),),
        ),
        HeaderFooterModel(
            kind="even",
            blocks=(ParagraphModel(runs=(RunModel(text="even"),)),),
        ),
    )
    paragraphs = tuple(
        ResolvedParagraphModel(
            runs=(ResolvedRunModel(text=f"body {index}"),),
            page_break_before=index > 0,
        )
        for index in range(3)
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=paragraphs, headers=headers),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(
        page.header[0].lines[0].fragments[0].text  # type: ignore[union-attr]
        for page in layout.pages
    ) == ("first", "even", "default")


def test_layout_engine_lays_out_table_in_header_region() -> None:
    table = TableModel(
        grid_widths=(80,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(runs=(RunModel(text="header cell"),)),),
                    ),
                ),
            ),
        ),
    )
    header = HeaderFooterModel(blocks=(table,))
    body = ResolvedParagraphModel(runs=(ResolvedRunModel(text="body"),))
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(body,), headers=(header,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].header[0].kind == "table"


def test_layout_engine_uses_physical_page_number_without_section_start() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="{{PAGE}}/{{NUMPAGES}}"),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_number_start=None,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].fragments[0].text == "1/1"  # type: ignore[union-attr]


def test_layout_engine_repositions_image_before_resolved_page_field() -> None:
    image = ImageModel(
        relationship_id="rIdImage",
        part_name="word/media/image.png",
        content_type="image/png",
        data=b"png",
        width=20,
        height=10,
    )
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(image=image), ResolvedRunModel(text="{{PAGE}}")),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,), page_number_start=None),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    fragments = layout.pages[0].body[0].lines[0].fragments  # type: ignore[union-attr]
    assert (
        fragments[0].kind,
        fragments[1].text,  # type: ignore[union-attr]
        fragments[1].x - fragments[0].x,
    ) == ("image", "1", 20.0)


def test_layout_engine_fills_remaining_page_room_before_splitting_table() -> None:
    # body.height is 50; the paragraph takes 11pt, leaving 39pt. Word's
    # default lets a row split across pages, so the table starts right where
    # the paragraph left off, row 0 (20pt) fits whole, and row 1 (20pt) is
    # split to use up the rest of the remaining room (19pt) rather than
    # moving whole to a fresh page.
    paragraph = ResolvedParagraphModel(runs=(ResolvedRunModel(text="before"),))
    table = TableModel(
        grid_widths=(80,),
        rows=tuple(
            TableRowModel(
                cells=(TableCellModel(paragraphs=(ParagraphModel(),)),),
                height=20,
                height_rule="exact",
            )
            for _ in range(4)
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph, table),
                page_width=100,
                page_height=70,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(tuple(block.kind for block in page.body) for page in layout.pages) == (
        ("paragraph", "table"),
        ("table",),
    )
    assert [warning.code for warning in layout.warnings] == ["table_row_split"]
    first_page_table = layout.pages[0].body[1]
    assert first_page_table.height == 39.0  # type: ignore[union-attr]
    assert [cell.row_index for cell in first_page_table.cells] == [0, 1]  # type: ignore[union-attr]
    second_page_table = layout.pages[1].body[0]
    assert [cell.row_index for cell in second_page_table.cells] == [2, 3]  # type: ignore[union-attr]


def test_layout_engine_abandons_cramped_table_start_without_overlapping_content() -> None:
    # Regression test for a bug found while implementing the "fill remaining
    # room" behaviour above: when a table starts mid-page but even its first
    # (header) row doesn't fit the sliver of room left over -- though it
    # would fit a fresh page -- that row must be anchored to the fresh
    # page's top, not to the cramped mid-page position. The bug placed it at
    # the stale mid-page y, causing it to visually overflow past the page
    # boundary and overlap whatever came next.
    filler = ResolvedParagraphModel(runs=(ResolvedRunModel(text="x"),), space_after=35.0)
    header = TableRowModel(
        cells=(TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text="H"),)),)),),
        height=10,
        height_rule="exact",
        repeat_header=True,
    )
    data_rows = tuple(
        TableRowModel(
            cells=(TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text=f"D{i}"),)),)),),
            height=20,
            height_rule="exact",
        )
        for i in range(3)
    )
    table = TableModel(grid_widths=(80,), rows=(header, *data_rows))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(filler, table),
                page_width=100,
                page_height=70,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    # Only ~4pt is left after the filler paragraph -- not enough for even
    # the 10pt header row -- so the whole table must move to a fresh page.
    assert tuple(tuple(block.kind for block in page.body) for page in layout.pages) == (
        ("paragraph",),
        ("table",),
        ("table",),
    )
    for page in layout.pages[1:]:
        table_box = page.body[0]
        assert table_box.y == 10.0  # type: ignore[union-attr]
        for cell in table_box.cells:  # type: ignore[union-attr]
            assert cell.y >= table_box.y  # type: ignore[union-attr]
            assert cell.y + cell.height <= table_box.y + table_box.height + 1e-6  # type: ignore[union-attr]


def test_layout_engine_moves_table_with_vertical_merge_whole_to_fresh_page() -> None:
    # A table containing a vertically merged cell is kept on the old, safe
    # path of moving whole to a fresh page when it doesn't fit the current
    # page's remainder -- starting it mid-page risks a merge boundary
    # landing on a chunk split, which _split_table_box cannot represent.
    paragraph = ResolvedParagraphModel(runs=(ResolvedRunModel(text="before"),))
    table = TableModel(
        grid_widths=(80,),
        rows=(
            TableRowModel(
                cells=(TableCellModel(paragraphs=(ParagraphModel(),), vertical_merge="restart"),),
                height=20,
                height_rule="exact",
            ),
            TableRowModel(
                cells=(TableCellModel(vertical_merge="continue"),),
                height=20,
                height_rule="exact",
            ),
            TableRowModel(
                cells=(TableCellModel(paragraphs=(ParagraphModel(),)),),
                height=20,
                height_rule="exact",
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph, table),
                page_width=100,
                page_height=70,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    # 39pt remain after the paragraph -- enough for the merged rows (40pt)
    # to fit if the table were allowed to start mid-page, but the table
    # must still move whole to a fresh page because it contains a merge.
    assert tuple(tuple(block.kind for block in page.body) for page in layout.pages) == (
        ("paragraph",),
        ("table",),
        ("table",),
    )


def test_layout_engine_splits_row_at_remaining_room_before_trying_a_fresh_page() -> None:
    # Word's default ("allow row to break across pages") splits an
    # oversized row at whatever room is left on the current page, rather
    # than moving it whole to a fresh page first. Here the row (9
    # paragraphs, 99pt) would fit a fresh 110pt body on its own, but must
    # still be split at the 90pt left after the filler paragraph instead of
    # being moved whole to page 2.
    filler = ResolvedParagraphModel(runs=(ResolvedRunModel(text="x"),), space_after=9.0)
    row = TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(9)),))
    table = TableModel(grid_widths=(100,), rows=(row,))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(filler, table),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert [warning.code for warning in layout.warnings] == ["table_row_split"]
    first_page_table = layout.pages[0].body[1]
    assert first_page_table.height == 90.0  # type: ignore[union-attr]
    second_page_table = layout.pages[1].body[0]
    assert second_page_table.height == 9.0  # type: ignore[union-attr]


def test_layout_engine_moves_cant_split_row_to_fresh_page_instead_of_splitting_it() -> None:
    # A cantSplit row must still move whole to a fresh page rather than
    # being cut at the room remaining on the current page, even though a
    # splittable row in the same spot would be split there instead (see
    # test_layout_engine_splits_row_at_remaining_room_before_trying_a_fresh_page).
    filler = ResolvedParagraphModel(runs=(ResolvedRunModel(text="x"),), space_after=9.0)
    row = TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(9)),), cant_split=True)
    table = TableModel(grid_widths=(100,), rows=(row,))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(filler, table),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert layout.warnings == ()
    assert tuple(tuple(block.kind for block in page.body) for page in layout.pages) == (
        ("paragraph",),
        ("table",),
    )
    table_box = layout.pages[1].body[0]
    assert table_box.height == 99.0  # type: ignore[union-attr]


def test_layout_engine_honors_page_break_before_in_mixed_section() -> None:
    first = ResolvedParagraphModel(runs=(ResolvedRunModel(text="before"),))
    table = TableModel(
        rows=(TableRowModel(cells=(TableCellModel(),)),),
    )
    last = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="after"),),
        page_break_before=True,
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(first, table, last)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(tuple(block.kind for block in page.body) for page in layout.pages) == (
        ("paragraph", "table"),
        ("paragraph",),
    )


def test_layout_engine_keeps_next_pair_together_in_mixed_section() -> None:
    filler = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="filler"),),
        space_after=36,
        source_index=0,
    )
    heading = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="heading"),),
        keep_next=True,
        source_index=1,
    )
    following = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="following"),),
        source_index=2,
    )
    table = TableModel(rows=(TableRowModel(cells=(TableCellModel(),)),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(filler, heading, following, table),
                page_width=120,
                page_height=80,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(block.paragraph_index for block in layout.pages[1].body[:2]) == (1, 2)  # type: ignore[union-attr]


def test_layout_engine_splits_long_paragraph_in_mixed_section() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three four five six seven eight nine ten"),),
        source_index=4,
    )
    table = TableModel(rows=(TableRowModel(cells=(TableCellModel(),)),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph, table),
                page_width=80,
                page_height=40,
                margin_left=10,
                margin_right=10,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    paragraph_segments = tuple(
        block for page in layout.pages for block in page.body if block.kind == "paragraph"
    )
    assert (
        paragraph_segments[0].continues_on_next_page,  # type: ignore[union-attr]
        paragraph_segments[-1].continued_from_previous_page,  # type: ignore[union-attr]
    ) == (True, True)


def test_layout_engine_moves_keep_lines_paragraph_in_mixed_section() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first"),),
        space_after=6,
        source_index=0,
    )
    kept = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three"),),
        keep_lines=True,
        source_index=1,
    )
    table = TableModel(rows=(TableRowModel(cells=(TableCellModel(),)),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(first, kept, table),
                page_width=80,
                page_height=46,
                margin_left=10,
                margin_right=10,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert tuple(block.paragraph_index for block in layout.pages[1].body[:1]) == (1,)  # type: ignore[union-attr]


def test_layout_engine_rejects_keep_lines_paragraph_taller_than_mixed_page() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three four five six"),),
        keep_lines=True,
        source_index=7,
    )
    table = TableModel(rows=(TableRowModel(cells=(TableCellModel(),)),))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph, table),
                page_width=60,
                page_height=30,
                margin_left=5,
                margin_right=5,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
    )

    with pytest.raises(LayoutError, match=r"paragraph\[7\]"):
        NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())


def test_layout_engine_overflows_table_row_taller_than_page_body_when_content_is_atomic() -> None:
    # An empty (content-less) row cannot be split, so it is placed whole and
    # allowed to overflow instead of raising -- conversions must never fail
    # just because one row is taller than the page.
    table = TableModel(
        rows=(
            TableRowModel(
                cells=(TableCellModel(),),
                height=60,
                height_rule="exact",
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_height=70,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 1
    assert [warning.code for warning in layout.warnings] == ["table_row_overflow"]


def test_layout_engine_drops_repeated_header_when_data_row_cannot_fit_below_it() -> None:
    # The data row fits a fresh page on its own (30 <= 50) but not once the
    # repeated header (30) is subtracted from the 50pt body. Rather than
    # failing, the header repeat is skipped for that one page.
    table = TableModel(
        rows=(
            TableRowModel(
                cells=(TableCellModel(),),
                height=30,
                height_rule="exact",
                repeat_header=True,
            ),
            TableRowModel(
                cells=(TableCellModel(),),
                height=30,
                height_rule="exact",
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_height=70,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert [block.kind for block in layout.pages[0].body] == ["table"]
    assert [block.kind for block in layout.pages[1].body] == ["table"]
    assert layout.warnings == ()


def _paragraphs(count: int) -> tuple[ParagraphModel, ...]:
    return tuple(ParagraphModel(runs=(RunModel(text=f"l{index}"),)) for index in range(count))


def test_layout_engine_splits_table_row_across_two_pages() -> None:
    # Each paragraph lays out as one 11pt line with this measurer/font size,
    # so 20 paragraphs need 220pt -- exactly two 110pt page bodies.
    table = TableModel(
        grid_widths=(100,),
        rows=(TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(20)),)),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert [warning.code for warning in layout.warnings] == ["table_row_split"]
    first, second = (page.body[0] for page in layout.pages)
    assert (first.continued_from_previous_page, first.continues_on_next_page) == (False, True)  # type: ignore[union-attr]
    assert (second.continued_from_previous_page, second.continues_on_next_page) == (True, False)  # type: ignore[union-attr]
    assert len(first.cells[0].blocks) == 10  # type: ignore[union-attr]
    assert len(second.cells[0].blocks) == 10  # type: ignore[union-attr]
    first_texts = [
        fragment.text
        for block in first.cells[0].blocks  # type: ignore[union-attr]
        for line in block.lines
        for fragment in line.fragments
    ]
    second_texts = [
        fragment.text
        for block in second.cells[0].blocks  # type: ignore[union-attr]
        for line in block.lines
        for fragment in line.fragments
    ]
    assert first_texts == [f"l{index}" for index in range(10)]
    assert second_texts == [f"l{index}" for index in range(10, 20)]


def test_layout_engine_splits_table_row_across_three_or_more_pages() -> None:
    # 30 paragraphs at 11pt need 330pt -- exactly three 110pt page bodies.
    table = TableModel(
        grid_widths=(100,),
        rows=(TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(30)),)),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 3
    assert [warning.code for warning in layout.warnings] == ["table_row_split", "table_row_split"]
    boxes = [page.body[0] for page in layout.pages]
    assert [b.continued_from_previous_page for b in boxes] == [False, True, True]  # type: ignore[union-attr]
    assert [b.continues_on_next_page for b in boxes] == [True, True, False]  # type: ignore[union-attr]
    assert [len(b.cells[0].blocks) for b in boxes] == [10, 10, 10]  # type: ignore[union-attr]


def test_layout_engine_keeps_cant_split_row_whole_when_it_fits_a_fresh_page() -> None:
    # The second row is marked cantSplit and fits a fresh 110pt page body on
    # its own (110pt); it must move to the next page whole, never split.
    rows = (
        TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(1)),)),
        TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(10)),), cant_split=True),
    )
    table = TableModel(grid_widths=(100,), rows=rows)
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert layout.warnings == ()
    assert [cell.row_index for cell in layout.pages[0].body[0].cells] == [0]  # type: ignore[union-attr]
    assert [cell.row_index for cell in layout.pages[1].body[0].cells] == [1]  # type: ignore[union-attr]


def test_layout_engine_splits_cant_split_row_when_taller_than_page_body() -> None:
    # cantSplit only protects a row that could fit a page on its own; once a
    # single row is taller than the whole body (220pt > 110pt) Word ends up
    # splitting it too, so we must fall back to splitting instead of failing.
    table = TableModel(
        grid_widths=(100,),
        rows=(
            TableRowModel(
                cells=(TableCellModel(paragraphs=_paragraphs(20)),),
                cant_split=True,
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert [warning.code for warning in layout.warnings] == ["table_row_split"]


def test_layout_engine_repeats_header_row_across_split_data_row_pages() -> None:
    # The data row (25 paragraphs, 275pt) does not fit even a fresh 110pt
    # body, so it is split; every page that carries a slice of it must still
    # repeat the header row at the top, per the existing repeat-header
    # mechanism. Word's default lets a row break at the room remaining after
    # the header, so the first page already carries header + a data slice
    # instead of a wasted header-only page.
    header = TableRowModel(
        cells=(TableCellModel(paragraphs=(ParagraphModel(runs=(RunModel(text="H"),)),)),),
        repeat_header=True,
    )
    data = TableRowModel(cells=(TableCellModel(paragraphs=_paragraphs(25)),))
    table = TableModel(grid_widths=(100,), rows=(header, data))
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=140,
                page_height=130,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 3
    assert [warning.code for warning in layout.warnings] == ["table_row_split", "table_row_split"]
    # Every page repeats row 0 (the header) ahead of the row-1 data slice --
    # including the first page, since the header row itself always fits.
    for page in layout.pages:
        assert [cell.row_index for cell in page.body[0].cells] == [0, 1]  # type: ignore[union-attr]
    data_texts = [
        fragment.text
        for page in layout.pages
        for cell in page.body[0].cells  # type: ignore[union-attr]
        if cell.row_index == 1
        for block in cell.blocks
        for line in block.lines
        for fragment in line.fragments
    ]
    assert data_texts == [f"l{index}" for index in range(25)]


def test_layout_engine_extends_and_scales_incomplete_table_grid() -> None:
    table = TableModel(
        grid_widths=(200,),
        rows=(TableRowModel(cells=(TableCellModel(), TableCellModel())),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=120,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    table_box = layout.pages[0].body[0]
    assert (table_box.column_widths, table_box.width) == ((80.0, 20.0), 100.0)  # type: ignore[union-attr]


@pytest.mark.parametrize(
    ("alignment", "expected_x"),
    (("center", 40.0), ("right", 70.0)),
)
def test_layout_engine_positions_narrow_table_by_alignment(
    alignment: str,
    expected_x: float,
) -> None:
    table = TableModel(
        grid_widths=(40,),
        rows=(TableRowModel(cells=(TableCellModel(),)),),
        alignment=alignment,  # type: ignore[arg-type]
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                page_width=120,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].x == expected_x


def test_layout_engine_center_aligns_cell_content_vertically() -> None:
    table = TableModel(
        grid_widths=(60,),
        rows=(
            TableRowModel(
                cells=(
                    TableCellModel(
                        paragraphs=(ParagraphModel(runs=(RunModel(text="center"),)),),
                        vertical_alignment="center",
                    ),
                ),
                height=40,
                height_rule="exact",
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(table,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    cell = layout.pages[0].body[0].cells[0]  # type: ignore[union-attr]
    paragraph = cell.blocks[0]
    assert paragraph.y + paragraph.height / 2 == cell.y + cell.height / 2


def test_layout_engine_moves_split_paragraph_when_no_line_fits_remaining_space() -> None:
    first = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="first"),),
        space_after=12,
        source_index=0,
    )
    long = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three four five six"),),
        widow_control=False,
        source_index=1,
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(first, long),
                page_width=80,
                page_height=40,
                margin_left=10,
                margin_right=10,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[1].body[0].paragraph_index == 1  # type: ignore[union-attr]


def test_layout_engine_rejects_keep_lines_paragraph_taller_than_plain_page() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="one two three four five six"),),
        keep_lines=True,
        source_index=8,
    )
    document = ResolvedDocumentModel(
        sections=(
            SectionModel(
                page_width=60,
                page_height=30,
                margin_left=5,
                margin_right=5,
                margin_top=5,
                margin_bottom=5,
            ),
        ),
        paragraphs=(paragraph,),
    )

    with pytest.raises(LayoutError, match=r"paragraph\[8\]"):
        NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())


def test_layout_engine_keeps_trailing_explicit_page_break_as_empty_page() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="before"), ResolvedRunModel(break_type="page")),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert (
        layout.page_count,
        layout.pages[1].body[0].lines[0].fragments,  # type: ignore[union-attr]
    ) == (2, ())


def test_layout_engine_prefixes_numbering_label_to_empty_paragraph() -> None:
    paragraph = ResolvedParagraphModel(numbering_label="1.")
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].fragments[0].text == "1. "  # type: ignore[union-attr]


def test_layout_engine_skips_hidden_and_empty_runs_and_uses_explicit_tab_stop() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="secret", hidden=True),
            ResolvedRunModel(text="A"),
            ResolvedRunModel(tab=True),
            ResolvedRunModel(),
            ResolvedRunModel(text="B"),
        ),
        tabs=(TabStop(position=50),),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    fragments = layout.pages[0].body[0].lines[0].fragments  # type: ignore[union-attr]
    assert (
        tuple(fragment.text for fragment in fragments),  # type: ignore[union-attr]
        fragments[1].x - fragments[0].x,
    ) == (("A", "B"), 50.0)


def test_layout_engine_preserves_empty_lines_around_hidden_and_empty_runs() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="hidden", hidden=True),
            ResolvedRunModel(break_type="line"),
            ResolvedRunModel(),
            ResolvedRunModel(break_type="line"),
            ResolvedRunModel(text="shown"),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    lines = layout.pages[0].body[0].lines  # type: ignore[union-attr]
    assert tuple("".join(fragment.text for fragment in line.fragments) for line in lines) == (
        "",
        "",
        "shown",
    )


@pytest.mark.parametrize(
    ("rule", "spacing", "expected_height"),
    (("exact", 17.0, 17.0), ("at_least", 20.0, 20.0)),
)
def test_layout_engine_applies_fixed_and_minimum_line_spacing(
    rule: str,
    spacing: float,
    expected_height: float,
) -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line"),),
        line_spacing_rule=rule,  # type: ignore[arg-type]
        line_spacing=spacing,
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].height == expected_height  # type: ignore[union-attr]


@pytest.mark.parametrize("alignment", ("center", "right"))
def test_layout_engine_offsets_center_and_right_aligned_lines(alignment: str) -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="line"),),
        alignment=alignment,  # type: ignore[arg-type]
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=120,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    line = layout.pages[0].body[0].lines[0]  # type: ignore[union-attr]
    expected = 10.0 + (line.width - line.used_width) / (2 if alignment == "center" else 1)
    assert line.x == expected


def test_layout_engine_falls_back_when_justified_token_has_no_break_slots() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="abcdefghijklmno", font_size=12),),
        alignment="justify",
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert len(layout.pages[0].body[0].lines[0].fragments) == 1  # type: ignore[union-attr]


@pytest.mark.parametrize(
    ("alignment", "line_spacing_rule", "line_spacing"),
    (("center", "exact", 17.0), ("right", "at_least", 20.0)),
)
def test_layout_engine_aligns_multi_run_line_with_configured_spacing(
    alignment: str,
    line_spacing_rule: str,
    line_spacing: float,
) -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="styled ", bold=True),
            ResolvedRunModel(text="line", italic=True),
        ),
        alignment=alignment,  # type: ignore[arg-type]
        line_spacing_rule=line_spacing_rule,  # type: ignore[arg-type]
        line_spacing=line_spacing,
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=160,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    line = layout.pages[0].body[0].lines[0]  # type: ignore[union-attr]
    expected_x = 10.0 + (line.width - line.used_width) / (2 if alignment == "center" else 1)
    assert (line.x, line.height) == (expected_x, line_spacing)


def test_layout_engine_wraps_multiple_styled_runs_without_losing_styles() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="one two ", bold=True),
            ResolvedRunModel(text="three four", italic=True),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    lines = layout.pages[0].body[0].lines  # type: ignore[union-attr]
    assert (
        len(lines) > 1,
        any(fragment.bold for line in lines for fragment in line.fragments),
        any(fragment.italic for line in lines for fragment in line.fragments),
    ) == (True, True, True)


def test_layout_engine_justifies_nonfinal_multi_run_line_at_space() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="one ", bold=True),
            ResolvedRunModel(text="two three", italic=True),
        ),
        alignment="justify",
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                margin_left=10,
                margin_right=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    first_line = layout.pages[0].body[0].lines[0]  # type: ignore[union-attr]
    assert (first_line.used_width, first_line.width, len(first_line.fragments) > 1) == (
        first_line.width,
        first_line.width,
        True,
    )


def test_layout_engine_skips_hidden_run_during_multi_run_wrapping() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(text="visible ", bold=True),
            ResolvedRunModel(text="secret", hidden=True),
            ResolvedRunModel(),
            ResolvedRunModel(text="text", italic=True),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(ResolvedSectionModel(blocks=(paragraph,)),),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    text = "".join(
        fragment.text
        for line in layout.pages[0].body[0].lines  # type: ignore[union-attr]
        for fragment in line.fragments
    )
    assert text == "visible text"


def test_layout_engine_snaps_line_height_to_doc_grid() -> None:
    # font_size=12 measures to a natural line height of 12pt with
    # _FixedTextMeasurer; a "lines" docGrid with an 18pt pitch must round
    # that up to a single, full 18pt grid line (Word's w:docGrid type="lines"
    # behaviour for the default Japanese template).
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="hello", font_name="Helvetica", font_size=12),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                doc_grid_type="lines",
                doc_grid_line_pitch=18.0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].height == 18.0  # type: ignore[union-attr]


def test_layout_engine_snaps_line_height_to_next_doc_grid_multiple() -> None:
    # font_size=20 measures to a 20pt natural line, which needs two 18pt
    # grid lines (ceil(20/18) == 2) to fully contain it.
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="hello", font_name="Helvetica", font_size=20),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                doc_grid_type="lines",
                doc_grid_line_pitch=18.0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].height == 36.0  # type: ignore[union-attr]


def test_layout_engine_does_not_snap_exact_line_spacing_to_doc_grid() -> None:
    # Word always honours an explicit fixed line height over the grid.
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="hello", font_name="Helvetica", font_size=12),),
        line_spacing_rule="exact",
        line_spacing=12.0,
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                doc_grid_type="lines",
                doc_grid_line_pitch=18.0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].height == 12.0  # type: ignore[union-attr]


def test_layout_engine_ignores_doc_grid_when_type_is_default() -> None:
    # w:docGrid w:type="default" (the common case) does not snap lines even
    # if a linePitch happens to be present.
    paragraph = ResolvedParagraphModel(
        runs=(ResolvedRunModel(text="hello", font_name="Helvetica", font_size=12),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                doc_grid_type="default",
                doc_grid_line_pitch=18.0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.pages[0].body[0].lines[0].height == pytest.approx(12.0)  # type: ignore[union-attr]


def test_layout_engine_snaps_table_cell_text_to_doc_grid() -> None:
    table = TableModel(
        grid_widths=(80,),
        rows=(TableRowModel(cells=(TableCellModel(paragraphs=(ParagraphModel(),)),)),),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                doc_grid_type="lines",
                doc_grid_line_pitch=18.0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    cell = layout.pages[0].body[0].cells[0]  # type: ignore[union-attr]
    # RunModel's default font size is 11pt; the empty paragraph falls back
    # to the same 11pt metric, so it too should round up to the 18pt grid.
    assert cell.blocks[0].lines[0].height == 18.0  # type: ignore[union-attr]


def test_layout_engine_does_not_snap_vertically_merged_cell_text_to_doc_grid() -> None:
    # Snapping a merged cell's height would change how many rows its
    # row-span covers, risking turning a previously fine layout into one
    # where the merge lands on a page-body chunk boundary that
    # _split_table_box cannot represent. Merged cells keep their natural,
    # unsnapped line heights.
    table = TableModel(
        grid_widths=(80,),
        rows=(
            TableRowModel(
                cells=(TableCellModel(paragraphs=(ParagraphModel(),), vertical_merge="restart"),),
            ),
            TableRowModel(cells=(TableCellModel(vertical_merge="continue"),)),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(table,),
                doc_grid_type="lines",
                doc_grid_line_pitch=18.0,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    merged_cell = next(
        cell
        for cell in layout.pages[0].body[0].cells  # type: ignore[union-attr]
        if cell.row_span == 2
    )
    assert merged_cell.blocks[0].lines[0].height == 11.0
