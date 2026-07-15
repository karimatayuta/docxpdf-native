"""Top-left-origin page layout independent of the PDF backend."""

from __future__ import annotations

import math
from collections import deque
from collections.abc import Callable
from typing import cast

from pydantic import BaseModel, ConfigDict, Field

from docxpdf_native.abstractions import LayoutEngine, TextMeasurer
from docxpdf_native.exceptions import LayoutError, PageLimitExceededError
from docxpdf_native.layout.japanese_breaking import (
    JapaneseLineBreakingRules,
    UnicodeText,
)
from docxpdf_native.layout.line_breaking import LineBreaker
from docxpdf_native.models import (
    BlockBox,
    CellBox,
    ConversionOptions,
    ConversionWarning,
    HeaderFooterModel,
    ImageBox,
    LayoutDocument,
    LineBox,
    PageModel,
    PageRegion,
    ParagraphBox,
    ParagraphModel,
    PlaceholderBox,
    ResolvedDocumentModel,
    ResolvedFont,
    ResolvedParagraphModel,
    ResolvedRunModel,
    ResolvedSectionModel,
    RunProperties,
    TableBox,
    TableCellModel,
    TableModel,
    TextFragment,
)


class _StyledCluster(BaseModel):
    model_config = ConfigDict(frozen=True)

    text: str
    start: int
    end: int
    width: float = Field(ge=0)
    run_index: int = Field(ge=0)
    run: ResolvedRunModel


# Rows shorter than this many points are treated as fully consumed when comparing
# against the remaining page height, absorbing floating point rounding noise.
_LAYOUT_EPSILON = 1e-6


class _RowUnit(BaseModel):
    """A queued table row (or the remainder of a previously split row)."""

    model_config = ConfigDict(frozen=True)

    row_index: int
    row_top: float
    height: float
    cells: tuple[CellBox, ...]
    cant_split: bool = False


class NativeLayoutEngine(LayoutEngine):
    """Lay out resolved document content using point units."""

    def __init__(self, text_measurer: TextMeasurer) -> None:
        self._text_measurer = text_measurer

    def layout(
        self,
        document: ResolvedDocumentModel,
        *,
        options: ConversionOptions,
    ) -> LayoutDocument:
        sections = document.sections or (ResolvedSectionModel(blocks=document.paragraphs),)
        has_section_blocks = any(section.blocks for section in sections)
        pages: list[PageModel] = []
        warnings: list[ConversionWarning] = []
        next_display_page_number = 1
        for section_index, section in enumerate(sections):
            resolved_section = (
                section
                if isinstance(section, ResolvedSectionModel)
                else ResolvedSectionModel.model_validate(section.model_dump(exclude={"columns"}))
            )
            paragraphs = tuple(
                block
                for block in resolved_section.blocks
                if isinstance(block, ResolvedParagraphModel)
            )
            tables = tuple(
                block for block in resolved_section.blocks if isinstance(block, TableModel)
            )
            warnings.extend(
                ConversionWarning(
                    code="table-autofit-fallback",
                    message=("Table AutoFit was replaced with deterministic equal-width columns."),
                    part="word/document.xml",
                    location=f"section[{section_index}]/table[{table_index}]",
                )
                for table_index, table in enumerate(tables)
                if table.autofit
            )
            if not has_section_blocks and section_index == 0:
                paragraphs = document.paragraphs
            if section_index > 0 and self._needs_parity_page(
                resolved_section.section_break,
                next_physical_page=len(pages) + 1,
            ):
                pages.append(
                    self._blank_page(
                        resolved_section,
                        number=len(pages) + 1,
                        section_index=section_index,
                        display_number=next_display_page_number,
                    )
                )
                next_display_page_number += 1
            if tables and paragraphs:
                section_pages = self._layout_mixed_section(
                    resolved_section,
                    section_index=section_index,
                    first_page_number=len(pages) + 1,
                    warnings=warnings,
                )
            elif tables:
                section_pages = self._layout_tables_only(
                    resolved_section,
                    tables=tables,
                    section_index=section_index,
                    first_page_number=len(pages) + 1,
                    warnings=warnings,
                )
            else:
                section_pages = self._layout_section(
                    resolved_section,
                    paragraphs=paragraphs,
                    section_index=section_index,
                    first_page_number=len(pages) + 1,
                )
            continuous = (
                section_index > 0
                and resolved_section.section_break == "continuous"
                and bool(pages)
                and self._same_page_geometry(pages[-1], resolved_section)
            )
            display_start = resolved_section.page_number_start
            if display_start is None:
                if continuous:
                    display_start = pages[-1].section_page_number or pages[-1].number
                else:
                    display_start = next_display_page_number
            section_pages = tuple(
                page.model_copy(update={"section_page_number": display_start + index})
                for index, page in enumerate(section_pages)
            )
            section_pages = self._add_header_footer(
                section_pages,
                section=resolved_section,
            )
            merged = False
            if continuous:
                merged = self._merge_continuous_first_page(pages, section_pages)
            if (
                continuous
                and not merged
                and resolved_section.page_number_start is None
                and display_start != next_display_page_number
            ):
                display_start = next_display_page_number
                section_pages = tuple(
                    page.model_copy(update={"section_page_number": display_start + index})
                    for index, page in enumerate(section_pages)
                )
                section_pages = self._add_header_footer(
                    section_pages,
                    section=resolved_section,
                )
            if merged:
                pages.extend(section_pages[1:])
                next_display_page_number = display_start + len(section_pages)
            else:
                pages.extend(section_pages)
                next_display_page_number = display_start + len(section_pages)
            if len(pages) > options.resource_limits.max_pages:
                raise PageLimitExceededError.for_limit(
                    actual=len(pages), maximum=options.resource_limits.max_pages
                )
        numbered_pages = tuple(
            page.model_copy(update={"number": index + 1}) for index, page in enumerate(pages)
        )
        resolved_pages = self._resolve_page_fields(numbered_pages)
        return LayoutDocument(pages=resolved_pages, warnings=tuple(warnings))

    @staticmethod
    def _needs_parity_page(section_break: str, *, next_physical_page: int) -> bool:
        if section_break == "even_page":
            return next_physical_page % 2 != 0
        if section_break == "odd_page":
            return next_physical_page % 2 != 1
        return False

    @staticmethod
    def _blank_page(
        section: ResolvedSectionModel,
        *,
        number: int,
        section_index: int,
        display_number: int,
    ) -> PageModel:
        return PageModel(
            number=number,
            width=section.page_width,
            height=section.page_height,
            section_index=section_index,
            section_page_number=display_number,
            body_region=PageRegion(
                x=section.margin_left,
                y=section.margin_top,
                width=section.page_width - section.margin_left - section.margin_right,
                height=section.page_height - section.margin_top - section.margin_bottom,
            ),
        )

    @staticmethod
    def _same_page_geometry(page: PageModel, section: ResolvedSectionModel) -> bool:
        body = page.body_region
        return (
            body is not None
            and page.width == section.page_width
            and page.height == section.page_height
            and body.x == section.margin_left
            and body.y == section.margin_top
            and body.width == section.page_width - section.margin_left - section.margin_right
            and body.height == section.page_height - section.margin_top - section.margin_bottom
        )

    def _merge_continuous_first_page(
        self,
        pages: list[PageModel],
        section_pages: tuple[PageModel, ...],
    ) -> bool:
        if not section_pages or not pages:
            return False
        previous = pages[-1]
        current = section_pages[0]
        body = previous.body_region
        current_body = current.body_region
        if body is None or current_body is None:
            return False
        cursor_y = max(
            (block.y + block.height for block in previous.body),
            default=body.y,
        )
        current_bottom = max(
            (block.y + block.height for block in current.body),
            default=current_body.y,
        )
        content_height = current_bottom - current_body.y
        if cursor_y + content_height > body.y + body.height:
            return False
        delta = cursor_y - current_body.y
        moved = tuple(self._move_page_block(block, delta=delta) for block in current.body)
        pages[-1] = previous.model_copy(
            update={
                "body": (*previous.body, *moved),
                "source_end": max(previous.source_end, current.source_end),
            }
        )
        return True

    def _move_page_block(self, block: BlockBox, *, delta: float) -> BlockBox:
        if isinstance(block, ParagraphBox):
            return self._move_paragraph(block, delta=delta)
        if isinstance(block, TableBox):
            return block.model_copy(
                update={
                    "y": block.y + delta,
                    "cells": tuple(self._move_cell(cell, delta=delta) for cell in block.cells),
                }
            )
        return block.model_copy(update={"y": block.y + delta})

    def _layout_mixed_section(
        self,
        section: ResolvedSectionModel,
        *,
        section_index: int,
        first_page_number: int,
        warnings: list[ConversionWarning],
    ) -> tuple[PageModel, ...]:
        body = PageRegion(
            x=section.margin_left,
            y=section.margin_top,
            width=section.page_width - section.margin_left - section.margin_right,
            height=section.page_height - section.margin_top - section.margin_bottom,
        )
        doc_grid_line_pitch = self._doc_grid_line_pitch(section)
        blocks: list[ResolvedParagraphModel | TableModel] = []
        for block in section.blocks:
            if isinstance(block, ResolvedParagraphModel):
                blocks.extend(self._split_page_breaks((block,)))
            else:
                blocks.append(block)
        blocks = self._collapse_paragraph_spacing(blocks)

        page_boxes: list[list[ParagraphBox | TableBox]] = [[]]
        cursor_y = body.y
        bottom = body.y + body.height
        table_index = 0
        for block_index, block in enumerate(blocks):
            if isinstance(block, TableModel):
                table_box = self._layout_table(
                    block,
                    body=body,
                    y=cursor_y,
                    table_index=table_index,
                    doc_grid_line_pitch=doc_grid_line_pitch,
                )
                table_index += 1
                if table_box.y + table_box.height <= bottom:
                    page_boxes[-1].append(table_box)
                    cursor_y = table_box.y + table_box.height
                    continue
                # The table does not fit in what's left of the current page.
                # Rather than moving it whole to a fresh page (which would
                # waste the remaining room here), let _split_table_box fill
                # the rest of this page with as many rows as fit and only
                # start a fresh page for what doesn't. Vertically merged
                # cells are the exception: starting mid-page shrinks the
                # first chunk and makes it more likely a merged cell's rows
                # get split across chunks, which _split_table_box cannot
                # represent and must reject. Keep those on the old, safe
                # path of moving the whole table to a fresh page first.
                has_vertical_merge = any(
                    cell.vertical_merge is not None for row in block.rows for cell in row.cells
                )
                if has_vertical_merge and page_boxes[-1]:
                    page_boxes.append([])
                    cursor_y = body.y
                    table_box = self._layout_table(
                        block,
                        body=body,
                        y=cursor_y,
                        table_index=table_index - 1,
                        doc_grid_line_pitch=doc_grid_line_pitch,
                    )
                header_count = 0
                for row in block.rows:
                    if not row.repeat_header:
                        break
                    header_count += 1
                chunks = self._split_table_box(
                    table_box,
                    body=body,
                    header_count=header_count,
                    warnings=warnings,
                )
                for chunk_index, chunk in enumerate(chunks):
                    # The leading chunk continues the current page only if
                    # it is actually anchored at the table's original
                    # (possibly mid-page) starting position. It might not
                    # be: _split_table_box abandons a too-narrow starting
                    # sliver and anchors to a fresh page instead when
                    # nothing fits there. Every other chunk always starts a
                    # fresh page.
                    if not (chunk_index == 0 and chunk.y == table_box.y):
                        page_boxes.append([])
                    page_boxes[-1].append(chunk)
                cursor_y = chunks[-1].y + chunks[-1].height
                continue

            paragraph = block
            if page_boxes[-1] and paragraph.page_break_before:
                page_boxes.append([])
                cursor_y = body.y
            available = body.model_copy(update={"y": cursor_y})
            paragraph_box = self._single_line_paragraph(
                paragraph, body=available, doc_grid_line_pitch=doc_grid_line_pitch
            )
            next_block = blocks[block_index + 1] if block_index + 1 < len(blocks) else None
            if (
                paragraph.keep_next
                and isinstance(next_block, ResolvedParagraphModel)
                and page_boxes[-1]
                and not next_block.page_break_before
            ):
                following_region = body.model_copy(
                    update={"y": paragraph_box.y + paragraph_box.height}
                )
                following_box = self._single_line_paragraph(
                    next_block, body=following_region, doc_grid_line_pitch=doc_grid_line_pitch
                )
                pair_height = paragraph_box.height + following_box.height
                if following_box.y + following_box.height > bottom and pair_height <= body.height:
                    page_boxes.append([])
                    cursor_y = body.y
                    available = body.model_copy(update={"y": cursor_y})
                    paragraph_box = self._single_line_paragraph(
                        paragraph, body=available, doc_grid_line_pitch=doc_grid_line_pitch
                    )
            if paragraph_box.y + paragraph_box.height <= bottom:
                page_boxes[-1].append(paragraph_box)
                cursor_y = paragraph_box.y + paragraph_box.height
                continue
            if len(paragraph_box.lines) > 1 and not paragraph.keep_lines:
                cursor_y = self._place_split_paragraph(
                    paragraph_box,
                    paragraph=paragraph,
                    pages=page_boxes,
                    body=body,
                    cursor_y=cursor_y,
                )
                continue
            if page_boxes[-1]:
                page_boxes.append([])
                cursor_y = body.y
                available = body.model_copy(update={"y": cursor_y})
                paragraph_box = self._single_line_paragraph(
                    paragraph, body=available, doc_grid_line_pitch=doc_grid_line_pitch
                )
            if paragraph_box.height > body.height:
                raise LayoutError(
                    "paragraph cannot fit in the available page body",
                    location=f"paragraph[{paragraph.source_index}]",
                )
            page_boxes[-1].append(paragraph_box)
            cursor_y = paragraph_box.y + paragraph_box.height

        return tuple(
            PageModel(
                number=first_page_number + index,
                width=section.page_width,
                height=section.page_height,
                section_index=section_index,
                body_region=body,
                section_page_number=(
                    section.page_number_start + index
                    if section.page_number_start is not None
                    else None
                ),
                body=tuple(boxes),
                source_start=boxes[0].source_start if boxes else 0,
                source_end=boxes[-1].source_end if boxes else 0,
            )
            for index, boxes in enumerate(page_boxes)
        )

    def _add_header_footer(
        self,
        pages: tuple[PageModel, ...],
        *,
        section: ResolvedSectionModel,
    ) -> tuple[PageModel, ...]:
        result: list[PageModel] = []
        header_region = PageRegion(
            x=section.margin_left,
            y=section.header_distance,
            width=section.page_width - section.margin_left - section.margin_right,
            height=max(0.0, section.margin_top - section.header_distance),
        )
        footer_region = PageRegion(
            x=section.margin_left,
            y=max(section.margin_top, section.page_height - section.footer_distance - 12.0),
            width=section.page_width - section.margin_left - section.margin_right,
            height=max(0.0, section.footer_distance),
        )
        for section_page_index, page in enumerate(pages):
            header_model = self._select_header_footer(
                section.headers,
                page_index=section_page_index,
                page_number=page.section_page_number or page.number,
            )
            footer_model = self._select_header_footer(
                section.footers,
                page_index=section_page_index,
                page_number=page.section_page_number or page.number,
            )
            result.append(
                page.model_copy(
                    update={
                        "header_region": header_region,
                        "footer_region": footer_region,
                        "header": self._layout_header_footer(header_model, region=header_region),
                        "footer": self._layout_header_footer(footer_model, region=footer_region),
                    }
                )
            )
        return tuple(result)

    @staticmethod
    def _select_header_footer(
        models: tuple[HeaderFooterModel, ...],
        *,
        page_index: int,
        page_number: int,
    ) -> HeaderFooterModel | None:
        by_kind = {model.kind: model for model in models}
        if page_index == 0 and "first" in by_kind:
            return by_kind["first"]
        if page_number % 2 == 0 and "even" in by_kind:
            return by_kind["even"]
        return by_kind.get("default") or (models[0] if models else None)

    def _layout_header_footer(
        self,
        model: HeaderFooterModel | None,
        *,
        region: PageRegion,
    ) -> tuple[ParagraphBox | TableBox | ImageBox, ...]:
        if model is None:
            return ()
        boxes: list[ParagraphBox | TableBox | ImageBox] = []
        cursor_y = region.y
        for block_index, block in enumerate(model.blocks):
            if isinstance(block, ParagraphModel):
                paragraph = self._resolve_table_paragraph(block)
                available = region.model_copy(update={"y": cursor_y})
                header_box: ParagraphBox | TableBox | ImageBox = self._single_line_paragraph(
                    paragraph,
                    body=available,
                )
            else:
                header_box = self._layout_table(
                    block,
                    body=region,
                    y=cursor_y,
                    table_index=block_index,
                )
            boxes.append(header_box)
            cursor_y += header_box.height
        return tuple(boxes)

    def _resolve_page_fields(self, pages: tuple[PageModel, ...]) -> tuple[PageModel, ...]:
        total_pages = len(pages)
        result: list[PageModel] = []
        for page in pages:
            page_value = page.section_page_number
            if page_value is None:
                page_value = page.number
            replacements = {
                "{{PAGE}}": str(page_value),
                "{{NUMPAGES}}": str(total_pages),
            }
            result.append(
                page.model_copy(
                    update={
                        "header": tuple(
                            self._replace_fields_in_block(block, replacements)
                            for block in page.header
                        ),
                        "body": tuple(
                            self._replace_fields_in_block(block, replacements)
                            for block in page.body
                        ),
                        "footer": tuple(
                            self._replace_fields_in_block(block, replacements)
                            for block in page.footer
                        ),
                    }
                )
            )
        return tuple(result)

    def _replace_fields_in_block(
        self,
        block: ParagraphBox | TableBox | ImageBox | object,
        replacements: dict[str, str],
    ) -> object:
        if isinstance(block, ParagraphBox):
            lines = tuple(self._replace_fields_in_line(line, replacements) for line in block.lines)
            return block.model_copy(update={"lines": lines})
        if isinstance(block, TableBox):
            cells = tuple(
                cell.model_copy(
                    update={
                        "blocks": tuple(
                            self._replace_fields_in_block(child, replacements)
                            for child in cell.blocks
                        )
                    }
                )
                for cell in block.cells
            )
            return block.model_copy(update={"cells": cells})
        return block

    def _replace_fields_in_line(
        self,
        line: LineBox,
        replacements: dict[str, str],
    ) -> LineBox:
        if not any(
            marker in fragment.text
            for fragment in line.fragments
            if isinstance(fragment, TextFragment)
            for marker in replacements
        ):
            return line
        fragments: list[TextFragment | ImageBox | PlaceholderBox] = []
        cursor_x = line.x
        for fragment in line.fragments:
            if not isinstance(fragment, TextFragment):
                fragments.append(fragment.model_copy(update={"x": cursor_x}))
                cursor_x += fragment.width
                continue
            text = fragment.text
            for marker, value in replacements.items():
                text = text.replace(marker, value)
            font = ResolvedFont(
                family=fragment.font_name,
                path=fragment.font_path,
                postscript_name=fragment.font_name,
                source="layout-field",
            )
            measurement = self._text_measurer.measure(text, font, fragment.font_size)
            fragments.append(
                fragment.model_copy(
                    update={
                        "x": cursor_x,
                        "text": text,
                        "width": measurement.width,
                        "height": measurement.height,
                    }
                )
            )
            cursor_x += measurement.width
        return line.model_copy(
            update={"fragments": tuple(fragments), "used_width": cursor_x - line.x}
        )

    def _layout_tables_only(
        self,
        section: ResolvedSectionModel,
        *,
        tables: tuple[TableModel, ...],
        section_index: int,
        first_page_number: int,
        warnings: list[ConversionWarning],
    ) -> tuple[PageModel, ...]:
        body = PageRegion(
            x=section.margin_left,
            y=section.margin_top,
            width=section.page_width - section.margin_left - section.margin_right,
            height=section.page_height - section.margin_top - section.margin_bottom,
        )
        doc_grid_line_pitch = self._doc_grid_line_pitch(section)
        page_tables: list[TableBox] = []
        for table_index, table in enumerate(tables):
            box = self._layout_table(
                table,
                body=body,
                y=body.y,
                table_index=table_index,
                doc_grid_line_pitch=doc_grid_line_pitch,
            )
            header_count = 0
            for row in table.rows:
                if not row.repeat_header:
                    break
                header_count += 1
            page_tables.extend(
                self._split_table_box(
                    box,
                    body=body,
                    header_count=header_count,
                    warnings=warnings,
                )
            )
        return tuple(
            PageModel(
                number=first_page_number + index,
                width=section.page_width,
                height=section.page_height,
                section_index=section_index,
                section_page_number=(
                    section.page_number_start + index
                    if section.page_number_start is not None
                    else None
                ),
                body_region=body,
                body=(table_box,),
                source_start=table_box.source_start,
                source_end=table_box.source_end,
            )
            for index, table_box in enumerate(page_tables)
        )

    def _split_table_box(
        self,
        box: TableBox,
        *,
        body: PageRegion,
        header_count: int,
        warnings: list[ConversionWarning],
    ) -> tuple[TableBox, ...]:
        """Paginate a table box, splitting rows across pages when necessary.

        ``box.y`` may already sit partway down the current page (a table
        that starts after some preceding paragraphs): the first chunk only
        gets the remaining room down to the bottom of the body, while every
        following chunk gets a fresh, full page starting at ``body.y``. Rows
        are placed whole whenever possible (mirroring Word's default of
        moving an oversized row to the next page). A row is only split at a
        page boundary once even a fresh page body cannot hold it -- at that
        point ``w:cantSplit`` can no longer be honoured, so we fall back to
        splitting deterministically rather than raising or looping forever.
        Vertically merged cells that span more than one row are kept out of
        scope: if pagination would cut through one, the previous hard error
        is preserved rather than guessing how to divide the merged content.
        """
        first_chunk_available = body.y + body.height - box.y
        if box.height <= first_chunk_available:
            return (box,)

        cells_by_row: dict[int, list[CellBox]] = {}
        for cell in box.cells:
            cells_by_row.setdefault(cell.row_index, []).append(cell)
        vmerge_rows = {cell.row_index for cell in box.cells if cell.row_span > 1}

        header_height = sum(box.row_heights[:header_count])
        if header_count > 0 and header_height > body.height:
            # The header itself cannot be repeated within the body; give up on
            # repeating it instead of failing the whole table.
            header_count = 0
            header_height = 0.0

        queue: deque[_RowUnit] = deque(
            _RowUnit(
                row_index=row_index,
                row_top=(cells_by_row[row_index][0].y if row_index in cells_by_row else body.y),
                height=box.row_heights[row_index],
                cells=tuple(cells_by_row.get(row_index, ())),
                cant_split=(
                    box.row_cant_split[row_index] if row_index < len(box.row_cant_split) else False
                ),
            )
            for row_index in range(len(box.row_heights))
        )

        chunk_specs: list[tuple[list[CellBox], float, bool, float]] = []
        row_boundaries: set[int] = set()
        current_cells: list[CellBox] = []
        current_height = 0.0
        current_header_applied = False
        # is_first_chunk controls whether header rows repeat: it stays True
        # until the first chunk is flushed, since that leading chunk already
        # contains the header rows themselves (nothing to repeat yet).
        is_first_chunk = True
        # chunk_anchor is the y the *current* chunk is being built from. It
        # starts at box.y, which may already sit partway down the page (a
        # table that starts after preceding content); if nothing fits in
        # that cramped remainder it is abandoned in favour of a fresh page
        # at body.y, independently of is_first_chunk/header semantics above.
        chunk_anchor = box.y

        def flush() -> None:
            nonlocal current_cells, current_height, is_first_chunk, current_header_applied
            nonlocal chunk_anchor
            chunk_specs.append(
                (current_cells, current_height, current_header_applied, chunk_anchor)
            )
            current_cells = []
            current_height = 0.0
            current_header_applied = False
            is_first_chunk = False
            chunk_anchor = body.y

        def place(unit: _RowUnit, *, height: float, header_applied: bool) -> None:
            nonlocal current_height, current_header_applied
            header_offset = header_height if header_applied else 0.0
            target_y = chunk_anchor + header_offset + current_height
            delta = target_y - unit.row_top
            for cell in unit.cells:
                moved = self._move_cell(cell, delta=delta)
                # A split-off remainder still carries the original row's
                # height; pin every cell to the height actually placed here.
                current_cells.append(moved.model_copy(update={"height": height}))
            current_height += height
            # Every row placed in a chunk must agree on whether the header was
            # repeated; the branches below only ever place one kind per chunk.
            current_header_applied = header_applied

        def try_split(unit: _RowUnit, *, budget: float, header_applied: bool) -> bool:
            """Place up to ``budget`` points of ``unit`` in the current
            chunk, pushing any remainder back to the front of the queue and
            flushing. Returns False, leaving the queue and current chunk
            untouched, if not even one line or atomic block (image,
            placeholder) fits within ``budget`` -- the caller falls back to
            trying a fresh page in that case.
            """
            budget = max(0.0, min(unit.height, budget))
            split_y = unit.row_top + budget
            top_cells, bottom_cells = self._slice_row_at(
                unit.cells, row_top=unit.row_top, split_y=split_y
            )
            if not any(cell.blocks for cell in top_cells):
                return False
            place(
                _RowUnit(
                    row_index=unit.row_index,
                    row_top=unit.row_top,
                    height=budget,
                    cells=tuple(top_cells),
                ),
                height=budget,
                header_applied=header_applied,
            )
            queue.popleft()
            warnings.append(
                ConversionWarning(
                    code="table_row_split",
                    message=(
                        "A table row does not fit within a single page body and was "
                        "split across pages."
                    ),
                    part="word/document.xml",
                    location=f"table[{box.table_index}]/row[{unit.row_index}]",
                )
            )
            if bottom_cells:
                queue.appendleft(
                    _RowUnit(
                        row_index=unit.row_index,
                        row_top=split_y,
                        height=unit.height - budget,
                        cells=tuple(bottom_cells),
                    )
                )
            flush()
            return True

        while queue:
            unit = queue[0]
            header_applied = header_count > 0 and not is_first_chunk
            chunk_limit = body.y + body.height - chunk_anchor
            available = chunk_limit - (header_height if header_applied else 0.0) - current_height
            if unit.height <= available + _LAYOUT_EPSILON:
                place(unit, height=unit.height, header_applied=header_applied)
                queue.popleft()
                continue

            # Word's default lets a row break across pages: a splittable row
            # (not cantSplit, not part of a vertical merge) is cut at the
            # room remaining in the current chunk before anything else is
            # tried. Only when not even the row's first line/atomic block
            # fits there does this fall through to moving the row whole to
            # a fresh page below.
            can_split = not unit.cant_split and unit.row_index not in vmerge_rows
            if can_split and try_split(unit, budget=available, header_applied=header_applied):
                continue

            if current_cells:
                # Try the row again at the top of a fresh page first.
                flush()
                row_boundaries.add(unit.row_index)
                continue
            if chunk_anchor != body.y:
                # Nothing fit in the cramped remainder left over from
                # preceding content on this page. Abandon that starting
                # position and re-evaluate this row against a genuine fresh
                # page instead -- is_first_chunk (and therefore whether a
                # header row repeats) is untouched by this.
                chunk_anchor = body.y
                continue
            fresh_header_applied = header_count > 0 and not is_first_chunk
            fresh_available = body.height - (header_height if fresh_header_applied else 0.0)
            if unit.height <= fresh_available + _LAYOUT_EPSILON:
                place(unit, height=unit.height, header_applied=fresh_header_applied)
                queue.popleft()
                continue
            if fresh_header_applied and unit.height <= body.height + _LAYOUT_EPSILON:
                # The row would fit a fresh page on its own; the repeated
                # header is what's crowding it out. Drop the repeat for this
                # one page rather than splitting the row unnecessarily.
                place(unit, height=unit.height, header_applied=False)
                queue.popleft()
                flush()
                continue
            # The row is taller than a full, header-free page body: it must
            # be split (or, for a vmerge span, treated as unsupported).
            # w:cantSplit can no longer be honoured at this point either --
            # Word itself ends up splitting (or overflowing) a row that
            # tall, so we fall back to splitting deterministically rather
            # than raising or looping forever.
            if unit.row_index in vmerge_rows:
                raise LayoutError(
                    "vertically merged cell cannot cross a page boundary",
                    location=f"table[{box.table_index}]/row[{unit.row_index}]",
                )
            # Keep repeating the header on split segments too, unless doing
            # so would leave no room at all for content.
            split_header_applied = fresh_header_applied
            split_budget = body.height - (header_height if split_header_applied else 0.0)
            if split_budget <= _LAYOUT_EPSILON:
                split_header_applied = False
                split_budget = body.height
            if try_split(unit, budget=split_budget, header_applied=split_header_applied):
                continue
            # Not even a single line/image fits a fresh page body (for
            # example a single image taller than the page). Place the row
            # whole and let it overflow instead of looping forever.
            place(unit, height=unit.height, header_applied=False)
            queue.popleft()
            warnings.append(
                ConversionWarning(
                    code="table_row_overflow",
                    message=(
                        "A table row is taller than the page body and could not be "
                        "split further; it overflows the page."
                    ),
                    part="word/document.xml",
                    location=f"table[{box.table_index}]/row[{unit.row_index}]",
                )
            )
            flush()

        if current_cells or not chunk_specs:
            flush()

        for cell in box.cells:
            if cell.row_span > 1 and any(
                cell.row_index < boundary < cell.row_index + cell.row_span
                for boundary in row_boundaries
            ):
                raise LayoutError(
                    "vertically merged cell cannot cross a page boundary",
                    location=(
                        f"table[{box.table_index}]/row[{cell.row_index}]"
                        f"/column[{cell.column_index}]"
                    ),
                )

        result: list[TableBox] = []
        for chunk_index, (data_cells, data_height, header_repeated, anchor_y) in enumerate(
            chunk_specs
        ):
            if header_repeated:
                # header_repeated only ever happens once is_first_chunk is
                # False, at which point chunk_anchor has already settled to
                # body.y -- so the header rows always belong at body.y too.
                header_delta = body.y - box.y
                header_cells = tuple(
                    self._move_cell(cell, delta=header_delta)
                    for cell in box.cells
                    if cell.row_index < header_count
                )
                cells = (*header_cells, *data_cells)
                total_height = header_height + data_height
            else:
                cells = tuple(data_cells)
                total_height = data_height
            result.append(
                box.model_copy(
                    update={
                        "y": anchor_y,
                        "height": total_height,
                        "cells": cells,
                        "row_heights": (total_height,),
                        "row_cant_split": (False,),
                        "continued_from_previous_page": chunk_index > 0,
                        "continues_on_next_page": chunk_index + 1 < len(chunk_specs),
                    }
                )
            )
        return tuple(result)

    def _slice_row_at(
        self,
        cells: tuple[CellBox, ...],
        *,
        row_top: float,
        split_y: float,
    ) -> tuple[list[CellBox], list[CellBox]]:
        """Split a row's cells at an absolute y position.

        Every cell is cut at the same ``split_y`` so the row breaks along one
        straight line, matching how Word visually splits a row across pages.
        Cells that end up with no content on one side simply keep their
        content on the other side; borders and background colours are copied
        onto both parts unchanged. The bottom half's cells are re-anchored to
        ``split_y`` so a later split (a row spanning three or more pages) can
        keep cutting relative to the correct top.
        """
        top_height = max(0.0, split_y - row_top)
        top_cells: list[CellBox] = []
        bottom_cells: list[CellBox] = []
        for cell in cells:
            top_blocks, bottom_blocks = self._slice_cell_blocks_at(cell.blocks, split_y=split_y)
            top_cells.append(cell.model_copy(update={"blocks": top_blocks, "height": top_height}))
            if bottom_blocks:
                bottom_cells.append(cell.model_copy(update={"blocks": bottom_blocks, "y": split_y}))
        return top_cells, bottom_cells

    @staticmethod
    def _slice_cell_blocks_at(
        blocks: tuple[ParagraphBox | ImageBox | PlaceholderBox, ...],
        *,
        split_y: float,
    ) -> tuple[
        tuple[ParagraphBox | ImageBox | PlaceholderBox, ...],
        tuple[ParagraphBox | ImageBox | PlaceholderBox, ...],
    ]:
        """Split a cell's content blocks at an absolute y position.

        Paragraphs are split line-by-line so that only whole lines cross the
        boundary; non-paragraph blocks (images, placeholders) are atomic and
        are placed entirely above or entirely below the split.
        """
        top: list[ParagraphBox | ImageBox | PlaceholderBox] = []
        bottom: list[ParagraphBox | ImageBox | PlaceholderBox] = []
        for block in blocks:
            if bottom:
                # Once content has spilled below the split, keep the
                # remainder together so paragraph order is preserved.
                bottom.append(block)
                continue
            if block.y + block.height <= split_y:
                top.append(block)
                continue
            if isinstance(block, ParagraphBox) and block.y < split_y:
                top_lines = tuple(line for line in block.lines if line.y + line.height <= split_y)
                bottom_lines = tuple(line for line in block.lines if line.y + line.height > split_y)
                if top_lines:
                    last = top_lines[-1]
                    top.append(
                        block.model_copy(
                            update={
                                "lines": top_lines,
                                "height": (last.y + last.height) - block.y,
                                "continues_on_next_page": True,
                            }
                        )
                    )
                if bottom_lines:
                    first = bottom_lines[0]
                    last = bottom_lines[-1]
                    bottom.append(
                        block.model_copy(
                            update={
                                "lines": bottom_lines,
                                "y": first.y,
                                "height": (last.y + last.height) - first.y,
                                "continued_from_previous_page": True,
                            }
                        )
                    )
                continue
            bottom.append(block)
        return tuple(top), tuple(bottom)

    def _move_cell(self, cell: CellBox, *, delta: float) -> CellBox:
        blocks: list[ParagraphBox | ImageBox | PlaceholderBox] = []
        for block in cell.blocks:
            if isinstance(block, ParagraphBox):
                blocks.append(
                    block.model_copy(
                        update={
                            "y": block.y + delta,
                            "lines": tuple(
                                self._move_line(line, y=line.y + delta) for line in block.lines
                            ),
                        }
                    )
                )
            else:
                blocks.append(block.model_copy(update={"y": block.y + delta}))
        return cell.model_copy(update={"y": cell.y + delta, "blocks": blocks})

    def _layout_table(
        self,
        table: TableModel,
        *,
        body: PageRegion,
        y: float,
        table_index: int,
        doc_grid_line_pitch: float | None = None,
    ) -> TableBox:
        column_count = max(
            len(table.grid_widths),
            max((sum(cell.grid_span for cell in row.cells) for row in table.rows), default=1),
        )
        if table.grid_widths:
            widths = list(table.grid_widths)
            if len(widths) < column_count:
                widths.extend([body.width / column_count] * (column_count - len(widths)))
            if table.width is not None:
                scale = table.width / sum(widths)
                widths = [width * scale for width in widths]
        else:
            preferred = [0.0] * column_count
            for row in table.rows:
                column_index = 0
                for cell in row.cells:
                    span = min(cell.grid_span, column_count - column_index)
                    if cell.width is not None and span > 0:
                        portion = cell.width / span
                        for index in range(column_index, column_index + span):
                            preferred[index] = max(preferred[index], portion)
                    column_index += span
            target_width = table.width or body.width
            if any(preferred):
                missing = [index for index, width in enumerate(preferred) if width == 0]
                remaining = max(0.0, target_width - sum(preferred))
                fill = remaining / len(missing) if missing else 0.0
                widths = [fill if width == 0 else width for width in preferred]
                if table.width is not None and sum(widths) > 0:
                    scale = table.width / sum(widths)
                    widths = [width * scale for width in widths]
            else:
                widths = [target_width / column_count] * column_count
        total_width = sum(widths)
        available_table_width = max(0.01, body.width - max(0.0, table.left_indent))
        if total_width > available_table_width:
            scale = available_table_width / total_width
            widths = [width * scale for width in widths]
            total_width = available_table_width
        x = body.x + table.left_indent
        if table.alignment == "center":
            x += max(0.0, body.width - total_width) / 2
        elif table.alignment == "right":
            x += max(0.0, body.width - total_width)
        cells: list[CellBox] = []
        row_heights: list[float] = []
        row_cant_split: list[bool] = []
        row_y = y
        for row_index, row in enumerate(table.rows):
            pending: list[
                tuple[
                    int,
                    float,
                    float,
                    float,
                    tuple[ParagraphBox, ...],
                    TableCellModel,
                ]
            ] = []
            column_index = 0
            calculated_height = 0.0
            for cell in row.cells:
                span = min(cell.grid_span, column_count - column_index)
                cell_width = sum(widths[column_index : column_index + span])
                margin_left = cell.margins.left or table.cell_margins.left
                margin_right = cell.margins.right or table.cell_margins.right
                margin_top = cell.margins.top or table.cell_margins.top
                margin_bottom = cell.margins.bottom or table.cell_margins.bottom
                paragraph_region = PageRegion(
                    x=x + sum(widths[:column_index]) + margin_left,
                    y=row_y + margin_top,
                    width=max(0.01, cell_width - margin_left - margin_right),
                    height=body.height,
                )
                paragraph_boxes: list[ParagraphBox] = []
                paragraph_y = paragraph_region.y
                previous_paragraph_after = 0.0
                # A vertically merged cell's height feeds directly into
                # whether its row-span can still fit in one page-body chunk
                # (see the vmerge boundary checks in _split_table_box).
                # Snapping its lines to the grid would change that height
                # and could turn a previously fine layout into an
                # unsplittable merge spanning a page boundary, so merged
                # cells keep their unsnapped natural line heights.
                cell_doc_grid_pitch = (
                    None if cell.vertical_merge is not None else doc_grid_line_pitch
                )
                for raw_paragraph in cell.paragraphs:
                    resolved = self._resolve_table_paragraph(raw_paragraph)
                    if cell.text_alignment is not None:
                        resolved = resolved.model_copy(update={"alignment": cell.text_alignment})
                    effective_before = max(0.0, resolved.space_before - previous_paragraph_after)
                    if effective_before != resolved.space_before:
                        resolved = resolved.model_copy(update={"space_before": effective_before})
                    previous_paragraph_after = resolved.space_after
                    region = paragraph_region.model_copy(update={"y": paragraph_y})
                    paragraph_box = self._single_line_paragraph(
                        resolved, body=region, doc_grid_line_pitch=cell_doc_grid_pitch
                    )
                    paragraph_boxes.append(paragraph_box)
                    paragraph_y += paragraph_box.height
                content_height = paragraph_y - row_y + margin_bottom
                calculated_height = max(calculated_height, content_height)
                pending.append(
                    (
                        column_index,
                        cell_width,
                        margin_top,
                        margin_bottom,
                        tuple(paragraph_boxes),
                        cell,
                    )
                )
                column_index += span
            if row.height_rule == "exact" and row.height is not None:
                row_height = row.height
            else:
                row_height = max(calculated_height, row.height or 0.0, 11.0)
            row_heights.append(row_height)
            row_cant_split.append(row.cant_split)
            for (
                column_index,
                cell_width,
                margin_top,
                margin_bottom,
                blocks,
                raw_cell,
            ) in pending:
                content_height = sum(block.height for block in blocks)
                available_height = max(0.0, row_height - margin_top - margin_bottom)
                if raw_cell.vertical_alignment == "bottom":
                    vertical_offset = max(0.0, available_height - content_height)
                elif raw_cell.vertical_alignment == "center":
                    vertical_offset = max(0.0, available_height - content_height) / 2
                else:
                    vertical_offset = 0.0
                aligned_blocks = tuple(
                    self._move_paragraph(block, delta=vertical_offset) for block in blocks
                )
                cells.append(
                    CellBox(
                        x=x + sum(widths[:column_index]),
                        y=row_y,
                        width=cell_width,
                        height=row_height,
                        row_index=row_index,
                        column_index=column_index,
                        column_span=raw_cell.grid_span,
                        blocks=aligned_blocks,
                        background_color=raw_cell.background_color,
                        borders=raw_cell.borders,
                        vertical_alignment=raw_cell.vertical_alignment,
                        source_start=raw_cell.source_index,
                        source_end=raw_cell.source_index + 1,
                    )
                )
            row_y += row_height
        cells = self._apply_vertical_merges(cells, table=table, row_heights=row_heights)
        return TableBox(
            x=x,
            y=y,
            width=total_width,
            height=sum(row_heights),
            cells=tuple(cells),
            column_widths=tuple(widths),
            row_heights=tuple(row_heights),
            row_cant_split=tuple(row_cant_split),
            table_index=table_index,
            borders=table.borders,
            source_start=table.source_index,
            source_end=table.source_index + 1,
        )

    def _move_paragraph(self, block: ParagraphBox, *, delta: float) -> ParagraphBox:
        return block.model_copy(
            update={
                "y": block.y + delta,
                "lines": tuple(self._move_line(line, y=line.y + delta) for line in block.lines),
            }
        )

    @staticmethod
    def _apply_vertical_merges(
        cells: list[CellBox],
        *,
        table: TableModel,
        row_heights: list[float],
    ) -> list[CellBox]:
        positions = {(cell.row_index, cell.column_index): index for index, cell in enumerate(cells)}
        active: dict[int, int] = {}
        removed: set[int] = set()
        for row_index, row in enumerate(table.rows):
            column_index = 0
            for raw_cell in row.cells:
                current_index = positions[(row_index, column_index)]
                if raw_cell.vertical_merge == "restart":
                    active[column_index] = current_index
                elif raw_cell.vertical_merge == "continue" and column_index in active:
                    start_index = active[column_index]
                    start_cell = cells[start_index]
                    cells[start_index] = start_cell.model_copy(
                        update={
                            "height": start_cell.height + row_heights[row_index],
                            "row_span": start_cell.row_span + 1,
                        }
                    )
                    removed.add(current_index)
                else:
                    active.pop(column_index, None)
                column_index += raw_cell.grid_span
        return [cell for index, cell in enumerate(cells) if index not in removed]

    @staticmethod
    def _resolve_table_paragraph(paragraph: ParagraphModel) -> ResolvedParagraphModel:
        raw = paragraph
        runs = tuple(
            ResolvedRunModel(
                text=run.text,
                font_name=NativeLayoutEngine._font_name(run.properties),
                font_path=run.properties.font_path,
                font_size=run.properties.font_size or 11.0,
                bold=bool(run.properties.bold),
                italic=bool(run.properties.italic),
                underline=run.properties.underline or False,
                strike=bool(run.properties.strike),
                color=run.properties.color or "000000",
                highlight=run.properties.highlight,
                character_spacing=run.properties.character_spacing or 0.0,
                break_type=run.break_type,
                tab=run.tab,
                image=run.image,
                hidden=run.hidden or bool(run.properties.hidden),
                vertical_align=run.properties.vertical_align or "baseline",
                source_index=run.source_index,
            )
            for run in raw.runs
        )
        properties = raw.properties
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
            source_index=raw.source_index,
        )

    @staticmethod
    def _font_name(properties: RunProperties) -> str:
        return (
            properties.east_asia_font
            or properties.font_family
            or properties.ascii_font
            or properties.high_ansi_font
            or "Helvetica"
        )

    def _layout_section(
        self,
        section: ResolvedSectionModel,
        *,
        paragraphs: tuple[ResolvedParagraphModel, ...],
        section_index: int,
        first_page_number: int,
    ) -> tuple[PageModel, ...]:
        body = PageRegion(
            x=section.margin_left,
            y=section.margin_top,
            width=section.page_width - section.margin_left - section.margin_right,
            height=section.page_height - section.margin_top - section.margin_bottom,
        )
        doc_grid_line_pitch = self._doc_grid_line_pitch(section)
        page_boxes: list[list[ParagraphBox | TableBox]] = [[]]
        cursor_y = body.y
        bottom = body.y + body.height
        expanded_paragraphs = cast(
            "tuple[ResolvedParagraphModel, ...]",
            tuple(self._collapse_paragraph_spacing(list(self._split_page_breaks(paragraphs)))),
        )
        for paragraph_index, paragraph in enumerate(expanded_paragraphs):
            if page_boxes[-1] and paragraph.page_break_before:
                page_boxes.append([])
                cursor_y = body.y
            available = body.model_copy(update={"y": cursor_y})
            box = self._single_line_paragraph(
                paragraph, body=available, doc_grid_line_pitch=doc_grid_line_pitch
            )
            if (
                paragraph.keep_next
                and paragraph_index + 1 < len(expanded_paragraphs)
                and page_boxes[-1]
                and not expanded_paragraphs[paragraph_index + 1].page_break_before
            ):
                following_region = body.model_copy(update={"y": box.y + box.height})
                following_box = self._single_line_paragraph(
                    expanded_paragraphs[paragraph_index + 1],
                    body=following_region,
                    doc_grid_line_pitch=doc_grid_line_pitch,
                )
                pair_height = box.height + following_box.height
                if following_box.y + following_box.height > bottom and pair_height <= body.height:
                    page_boxes.append([])
                    cursor_y = body.y
                    available = body.model_copy(update={"y": cursor_y})
                    box = self._single_line_paragraph(
                        paragraph, body=available, doc_grid_line_pitch=doc_grid_line_pitch
                    )
            if box.y + box.height <= bottom:
                page_boxes[-1].append(box)
                cursor_y = box.y + box.height
                continue
            if len(box.lines) > 1 and not paragraph.keep_lines:
                cursor_y = self._place_split_paragraph(
                    box,
                    paragraph=paragraph,
                    pages=page_boxes,
                    body=body,
                    cursor_y=cursor_y,
                )
                continue
            if page_boxes[-1]:
                page_boxes.append([])
                cursor_y = body.y
                available = body.model_copy(update={"y": cursor_y})
                box = self._single_line_paragraph(
                    paragraph, body=available, doc_grid_line_pitch=doc_grid_line_pitch
                )
            if box.height > body.height:
                raise LayoutError(
                    "paragraph cannot fit in the available page body",
                    location=f"paragraph[{paragraph.source_index}]",
                )
            page_boxes[-1].append(box)
            cursor_y = box.y + box.height
        pages = tuple(
            PageModel(
                number=first_page_number + index,
                width=section.page_width,
                height=section.page_height,
                section_index=section_index,
                body_region=body,
                section_page_number=(
                    section.page_number_start + index
                    if section.page_number_start is not None
                    else None
                ),
                body=tuple(boxes),
                source_start=(
                    boxes[0].paragraph_index if boxes and isinstance(boxes[0], ParagraphBox) else 0
                ),
                source_end=(
                    boxes[-1].paragraph_index + 1
                    if boxes and isinstance(boxes[-1], ParagraphBox)
                    else 0
                ),
            )
            for index, boxes in enumerate(page_boxes)
        )
        return pages

    def _place_split_paragraph(
        self,
        box: ParagraphBox,
        *,
        paragraph: ResolvedParagraphModel,
        pages: list[list[ParagraphBox | TableBox]],
        body: PageRegion,
        cursor_y: float,
    ) -> float:
        remaining = list(box.lines)
        first_segment = True
        while remaining:
            top_space = paragraph.space_before if first_segment else 0.0
            available_height = body.y + body.height - cursor_y - top_space
            fit = 0
            used_height = 0.0
            for line in remaining:
                if used_height + line.height > available_height:
                    break
                used_height += line.height
                fit += 1
            if (
                paragraph.widow_control
                and first_segment
                and pages[-1]
                and fit == 1
                and len(remaining) > 1
            ):
                pages.append([])
                cursor_y = body.y
                continue
            if paragraph.widow_control and len(remaining) - fit == 1 and fit > 1:
                fit -= 1
                used_height = sum(line.height for line in remaining[:fit])
            if fit == 0:
                if pages[-1]:
                    pages.append([])
                    cursor_y = body.y
                    continue
                raise LayoutError(
                    "a line cannot fit in the available page body",
                    location=f"paragraph[{paragraph.source_index}]",
                )
            selected = remaining[:fit]
            remaining = remaining[fit:]
            target_y = cursor_y + top_space
            positioned: list[LineBox] = []
            for line in selected:
                positioned.append(self._move_line(line, y=target_y))
                target_y += line.height
            bottom_space = paragraph.space_after if not remaining else 0.0
            segment = ParagraphBox(
                x=box.x,
                y=cursor_y,
                width=box.width,
                height=top_space + used_height + bottom_space,
                lines=tuple(positioned),
                paragraph_index=box.paragraph_index,
                continued_from_previous_page=not first_segment,
                continues_on_next_page=bool(remaining),
                source_start=positioned[0].source_start,
                source_end=positioned[-1].source_end,
            )
            pages[-1].append(segment)
            cursor_y += segment.height
            first_segment = False
            if remaining:
                pages.append([])
                cursor_y = body.y
        return cursor_y

    @staticmethod
    def _move_line(line: LineBox, *, y: float) -> LineBox:
        delta = y - line.y
        fragments = tuple(
            fragment.model_copy(update={"y": fragment.y + delta}) for fragment in line.fragments
        )
        return line.model_copy(update={"y": y, "fragments": fragments})

    @staticmethod
    def _split_page_breaks(
        paragraphs: tuple[ResolvedParagraphModel, ...],
    ) -> tuple[ResolvedParagraphModel, ...]:
        expanded: list[ResolvedParagraphModel] = []
        for paragraph in paragraphs:
            current: list[ResolvedRunModel] = []
            force_before = paragraph.page_break_before
            found_break = False
            for run in paragraph.runs:
                if run.break_type != "page":
                    current.append(run)
                    continue
                expanded.append(
                    paragraph.model_copy(
                        update={"runs": tuple(current), "page_break_before": force_before}
                    )
                )
                current = []
                force_before = True
                found_break = True
            if current or not found_break:
                expanded.append(
                    paragraph.model_copy(
                        update={"runs": tuple(current), "page_break_before": force_before}
                    )
                )
            elif found_break:
                expanded.append(
                    paragraph.model_copy(update={"runs": (), "page_break_before": True})
                )
        return tuple(expanded)

    @staticmethod
    def _collapse_paragraph_spacing(
        blocks: list[ResolvedParagraphModel | TableModel],
    ) -> list[ResolvedParagraphModel | TableModel]:
        """Apply Word's paragraph-spacing collapse between adjacent paragraphs.

        Word does not add a paragraph's ``space_after`` to the next
        paragraph's ``space_before``; the gap between them is the larger of
        the two. Since a paragraph's own box height already accounts for its
        ``space_after``, the collapse is applied by shrinking the following
        paragraph's ``space_before`` down to only the amount not already
        covered by the previous paragraph's ``space_after``. Non-paragraph
        blocks (tables) and explicit page breaks reset the previous
        paragraph's contribution, since nothing collapses across them.
        """
        result: list[ResolvedParagraphModel | TableModel] = []
        previous_after = 0.0
        for block in blocks:
            if not isinstance(block, ResolvedParagraphModel):
                result.append(block)
                previous_after = 0.0
                continue
            if block.page_break_before:
                previous_after = 0.0
            effective_before = max(0.0, block.space_before - previous_after)
            if effective_before != block.space_before:
                block = block.model_copy(update={"space_before": effective_before})
            result.append(block)
            previous_after = block.space_after
        return result

    @staticmethod
    def _snap_line_height_to_doc_grid(
        line_height: float,
        *,
        line_spacing_rule: str,
        doc_grid_line_pitch: float | None,
    ) -> float:
        """Snap a line's height up to the section's ``w:docGrid`` line pitch.

        Word sections with ``w:docGrid w:type="lines"`` (or
        ``"linesAndChars"``) round every line's natural height up to the
        nearest multiple of ``w:linePitch``, so a 10.5pt font whose natural
        line is well under an 18pt grid still consumes a full grid line.
        Skipping this makes native text denser than Word's and understates
        page counts for documents that rely on the default Japanese
        template, which sets this grid. A line whose spacing rule is
        ``exact`` is left untouched -- Word always honours an explicit fixed
        line height over the grid.
        """
        if doc_grid_line_pitch is None or line_spacing_rule == "exact":
            return line_height
        grid_lines = math.ceil(line_height / doc_grid_line_pitch - _LAYOUT_EPSILON)
        return max(1, grid_lines) * doc_grid_line_pitch

    @staticmethod
    def _doc_grid_line_pitch(section: ResolvedSectionModel) -> float | None:
        """Return the section's line-grid pitch, or ``None`` if it has none.

        Only ``w:docGrid`` types that actually snap line heights to a grid
        (``lines`` and ``linesAndChars``) apply; ``default`` and
        ``snapToChars`` leave line heights alone.
        """
        if section.doc_grid_type not in {"lines", "linesAndChars"}:
            return None
        return section.doc_grid_line_pitch

    def _single_line_paragraph(
        self,
        paragraph: ResolvedParagraphModel,
        *,
        body: PageRegion,
        doc_grid_line_pitch: float | None = None,
    ) -> ParagraphBox:
        if paragraph.numbering_label is not None:
            runs = list(paragraph.runs)
            if runs and runs[0].text:
                runs[0] = runs[0].model_copy(
                    update={"text": f"{paragraph.numbering_label} {runs[0].text}"}
                )
            else:
                runs.insert(
                    0,
                    ResolvedRunModel(text=f"{paragraph.numbering_label} "),
                )
            return self._single_line_paragraph(
                paragraph.model_copy(update={"runs": tuple(runs), "numbering_label": None}),
                body=body,
                doc_grid_line_pitch=doc_grid_line_pitch,
            )
        if any(run.break_type == "line" for run in paragraph.runs):
            return self._explicit_line_paragraph(
                paragraph, body=body, doc_grid_line_pitch=doc_grid_line_pitch
            )
        visible_runs = tuple(
            run for run in paragraph.runs if not run.hidden and run.text and run.image is None
        )
        plain_text_runs = all(
            run.break_type is None and not run.tab and run.image is None and run.placeholder is None
            for run in paragraph.runs
        )
        if len(visible_runs) == 1 and plain_text_runs:
            return self._wrapped_text_paragraph(
                paragraph, run=visible_runs[0], body=body, doc_grid_line_pitch=doc_grid_line_pitch
            )
        if len(visible_runs) > 1 and plain_text_runs:
            return self._wrapped_runs_paragraph(
                paragraph, body=body, doc_grid_line_pitch=doc_grid_line_pitch
            )

        x = body.x + paragraph.left_indent + paragraph.first_line_indent - paragraph.hanging_indent
        y = body.y + paragraph.space_before
        fragments: list[TextFragment | ImageBox | PlaceholderBox] = []
        used_width = 0.0
        ascent = 0.0
        descent = 0.0
        source_offset = 0
        for run in paragraph.runs:
            if run.hidden:
                continue
            if run.tab:
                explicit_stops = tuple(
                    stop.position for stop in paragraph.tabs if stop.position > used_width
                )
                if explicit_stops:
                    used_width = min(explicit_stops)
                else:
                    used_width = (math.floor(used_width / 36.0) + 1) * 36.0
                source_offset += 1
                continue
            if run.image is not None:
                available_width = max(
                    0.01,
                    body.width - paragraph.left_indent - paragraph.right_indent - used_width,
                )
                scale = min(1.0, available_width / run.image.width)
                width = run.image.width * scale
                height = run.image.height * scale
                fragments.append(
                    ImageBox(
                        x=x + used_width,
                        y=y,
                        width=width,
                        height=height,
                        image_data=run.image.data,
                        content_type=run.image.content_type,
                        part_name=run.image.part_name,
                        relationship_id=run.image.relationship_id,
                        source_start=source_offset,
                        source_end=source_offset + 1,
                    )
                )
                used_width += width
                ascent = max(ascent, height)
                source_offset += 1
                continue
            if run.placeholder is not None:
                available_width = max(
                    0.01,
                    body.width - paragraph.left_indent - paragraph.right_indent - used_width,
                )
                scale = min(1.0, available_width / run.placeholder.width)
                width = run.placeholder.width * scale
                height = run.placeholder.height * scale
                fragments.append(
                    PlaceholderBox(
                        x=x + used_width,
                        y=y,
                        width=width,
                        height=height,
                        label=run.placeholder.label,
                        source_start=source_offset,
                        source_end=source_offset + 1,
                    )
                )
                used_width += width
                ascent = max(ascent, height)
                source_offset += 1
                continue
            if not run.text:
                continue
            font = ResolvedFont(
                family=run.font_name,
                path=run.font_path,
                postscript_name=run.font_name,
                source="resolved-run",
            )
            measurement = self._text_measurer.measure(
                run.text,
                font,
                run.font_size,
                character_spacing=run.character_spacing,
            )
            fragments.append(
                TextFragment(
                    x=x + used_width,
                    y=y,
                    width=measurement.width,
                    height=measurement.height,
                    text=run.text,
                    font_name=run.font_name,
                    font_path=run.font_path,
                    font_size=run.font_size,
                    bold=run.bold,
                    italic=run.italic,
                    underline=run.underline,
                    strike=run.strike,
                    color=run.color,
                    highlight=run.highlight,
                    character_spacing=run.character_spacing,
                    baseline_shift=self._baseline_shift(run),
                    source_start=source_offset,
                    source_end=source_offset + len(run.text),
                )
            )
            used_width += measurement.width
            ascent = max(ascent, measurement.ascent)
            descent = max(descent, measurement.descent)
            source_offset += len(run.text)
        if not fragments:
            default_font = ResolvedFont(
                family="Helvetica",
                postscript_name="Helvetica",
                source="empty-paragraph",
            )
            metric = self._text_measurer.measure("Ag", default_font, 11.0)
            ascent = metric.ascent
            descent = metric.descent
        height = self._snap_line_height_to_doc_grid(
            ascent + descent,
            line_spacing_rule=paragraph.line_spacing_rule,
            doc_grid_line_pitch=doc_grid_line_pitch,
        )
        line = LineBox(
            x=x,
            y=y,
            width=max(0, body.width - paragraph.left_indent - paragraph.right_indent),
            height=height,
            used_width=used_width,
            ascent=ascent,
            descent=descent,
            baseline=ascent,
            fragments=tuple(fragments),
            source_start=0,
            source_end=source_offset,
        )
        return ParagraphBox(
            x=x,
            y=y,
            width=line.width,
            height=height + paragraph.space_before + paragraph.space_after,
            lines=(line,),
            paragraph_index=paragraph.source_index,
            source_start=0,
            source_end=source_offset,
        )

    def _explicit_line_paragraph(
        self,
        paragraph: ResolvedParagraphModel,
        *,
        body: PageRegion,
        doc_grid_line_pitch: float | None = None,
    ) -> ParagraphBox:
        groups: list[list[ResolvedRunModel]] = [[]]
        for run in paragraph.runs:
            if run.break_type == "line":
                groups.append([])
            elif not run.hidden:
                groups[-1].append(run)
        lines: list[LineBox] = []
        y = body.y + paragraph.space_before
        source_offset = 0
        for line_index, group in enumerate(groups):
            x = body.x + paragraph.left_indent
            if line_index == 0:
                x += paragraph.first_line_indent - paragraph.hanging_indent
            fragments: list[TextFragment] = []
            used_width = 0.0
            ascent = 0.0
            descent = 0.0
            for run in group:
                if not run.text:
                    continue
                font = ResolvedFont(
                    family=run.font_name,
                    path=run.font_path,
                    postscript_name=run.font_name,
                    source="resolved-run",
                )
                measurement = self._text_measurer.measure(
                    run.text,
                    font,
                    run.font_size,
                    character_spacing=run.character_spacing,
                )
                fragments.append(
                    TextFragment(
                        x=x + used_width,
                        y=y,
                        width=measurement.width,
                        height=measurement.height,
                        text=run.text,
                        font_name=run.font_name,
                        font_path=run.font_path,
                        font_size=run.font_size,
                        bold=run.bold,
                        italic=run.italic,
                        underline=run.underline,
                        strike=run.strike,
                        color=run.color,
                        highlight=run.highlight,
                        character_spacing=run.character_spacing,
                        baseline_shift=self._baseline_shift(run),
                        source_start=source_offset,
                        source_end=source_offset + len(run.text),
                    )
                )
                source_offset += len(run.text)
                used_width += measurement.width
                ascent = max(ascent, measurement.ascent)
                descent = max(descent, measurement.descent)
            if not group:
                source_offset += 1
                ascent = 8.8
                descent = 2.2
            line_height = self._snap_line_height_to_doc_grid(
                ascent + descent,
                line_spacing_rule=paragraph.line_spacing_rule,
                doc_grid_line_pitch=doc_grid_line_pitch,
            )
            lines.append(
                LineBox(
                    x=x,
                    y=y,
                    width=max(0, body.width - paragraph.left_indent - paragraph.right_indent),
                    height=line_height,
                    used_width=used_width,
                    ascent=ascent,
                    descent=descent,
                    baseline=ascent,
                    fragments=tuple(fragments),
                    source_start=fragments[0].source_start if fragments else source_offset - 1,
                    source_end=fragments[-1].source_end if fragments else source_offset,
                )
            )
            y += line_height
        return ParagraphBox(
            x=body.x + paragraph.left_indent,
            y=body.y,
            width=max(0, body.width - paragraph.left_indent - paragraph.right_indent),
            height=(
                paragraph.space_before + sum(line.height for line in lines) + paragraph.space_after
            ),
            lines=tuple(lines),
            paragraph_index=paragraph.source_index,
            source_start=0,
            source_end=source_offset,
        )

    def _wrapped_runs_paragraph(
        self,
        paragraph: ResolvedParagraphModel,
        *,
        body: PageRegion,
        doc_grid_line_pitch: float | None = None,
    ) -> ParagraphBox:
        clusters = self._styled_clusters(paragraph)
        first_offset = paragraph.first_line_indent - paragraph.hanging_indent
        first_width = max(
            0.01,
            body.width - paragraph.left_indent - paragraph.right_indent - first_offset,
        )
        following_width = max(
            0.01,
            body.width - paragraph.left_indent - paragraph.right_indent,
        )
        broken_lines = self._break_styled_clusters(
            clusters,
            first_width=first_width,
            following_width=following_width,
        )
        lines: list[LineBox] = []
        y = body.y + paragraph.space_before
        for line_index, line_clusters in enumerate(broken_lines):
            line_width = first_width if line_index == 0 else following_width
            line_x = (
                body.x + paragraph.left_indent + first_offset
                if line_index == 0
                else body.x + paragraph.left_indent
            )
            natural_width = sum(cluster.width for cluster in line_clusters)
            if paragraph.alignment == "center":
                line_x += max(0.0, line_width - natural_width) / 2
            elif paragraph.alignment == "right":
                line_x += max(0.0, line_width - natural_width)

            slots: tuple[int, ...] = ()
            if paragraph.alignment == "justify" and line_index + 1 < len(broken_lines):
                slots = self._styled_justification_slots(line_clusters)
            extra = max(0.0, line_width - natural_width) / len(slots) if slots else 0.0
            ascent = 0.0
            descent = 0.0
            for cluster in line_clusters:
                font = ResolvedFont(
                    family=cluster.run.font_name,
                    path=cluster.run.font_path,
                    postscript_name=cluster.run.font_name,
                    source="resolved-run",
                )
                metric = self._text_measurer.measure("Ag", font, cluster.run.font_size)
                ascent = max(ascent, metric.ascent)
                descent = max(descent, metric.descent)
            natural_height = ascent + descent
            if paragraph.line_spacing_rule == "exact":
                line_height = paragraph.line_spacing
            elif paragraph.line_spacing_rule == "at_least":
                line_height = max(natural_height, paragraph.line_spacing)
            else:
                line_height = natural_height * paragraph.line_spacing
            line_height = self._snap_line_height_to_doc_grid(
                line_height,
                line_spacing_rule=paragraph.line_spacing_rule,
                doc_grid_line_pitch=doc_grid_line_pitch,
            )

            fragments: list[TextFragment] = []
            cursor_x = line_x
            group: list[_StyledCluster] = []
            for cluster_index, cluster in enumerate(line_clusters):
                if group and (
                    group[-1].run_index != cluster.run_index or cluster_index - 1 in slots
                ):
                    cursor_x = self._append_styled_fragment(
                        fragments,
                        group,
                        x=cursor_x,
                        y=y,
                        height=line_height,
                        extra_after=extra if cluster_index - 1 in slots else 0.0,
                    )
                    group = []
                group.append(cluster)
            if group:
                self._append_styled_fragment(
                    fragments,
                    group,
                    x=cursor_x,
                    y=y,
                    height=line_height,
                    extra_after=extra if len(line_clusters) - 1 in slots else 0.0,
                )
            used_width = line_width if slots else natural_width
            lines.append(
                LineBox(
                    x=line_x,
                    y=y,
                    width=line_width,
                    height=line_height,
                    used_width=used_width,
                    ascent=ascent,
                    descent=descent,
                    baseline=ascent,
                    fragments=tuple(fragments),
                    source_start=line_clusters[0].start,
                    source_end=line_clusters[-1].end,
                )
            )
            y += line_height
        source_end = sum(len(run.text) for run in paragraph.runs)
        return ParagraphBox(
            x=body.x + paragraph.left_indent,
            y=body.y,
            width=max(0.0, body.width - paragraph.left_indent - paragraph.right_indent),
            height=paragraph.space_before
            + sum(line.height for line in lines)
            + paragraph.space_after,
            lines=tuple(lines),
            paragraph_index=paragraph.source_index,
            source_start=0,
            source_end=source_end,
        )

    def _styled_clusters(
        self,
        paragraph: ResolvedParagraphModel,
    ) -> tuple[_StyledCluster, ...]:
        result: list[_StyledCluster] = []
        source_offset = 0
        for run_index, run in enumerate(paragraph.runs):
            run_start = source_offset
            source_offset += len(run.text)
            if run.hidden or not run.text:
                continue
            font = ResolvedFont(
                family=run.font_name,
                path=run.font_path,
                postscript_name=run.font_name,
                source="resolved-run",
            )
            run_clusters = UnicodeText.indexed_grapheme_clusters(run.text)
            for cluster_index, cluster in enumerate(run_clusters):
                width = self._text_measurer.measure(
                    cluster.text,
                    font,
                    run.font_size,
                    character_spacing=run.character_spacing,
                ).width
                if cluster_index + 1 < len(run_clusters):
                    width += run.character_spacing
                result.append(
                    _StyledCluster(
                        text=cluster.text,
                        start=run_start + cluster.start,
                        end=run_start + cluster.end,
                        width=width,
                        run_index=run_index,
                        run=run,
                    )
                )
        return tuple(result)

    @staticmethod
    def _break_styled_clusters(
        clusters: tuple[_StyledCluster, ...],
        *,
        first_width: float,
        following_width: float,
    ) -> tuple[tuple[_StyledCluster, ...], ...]:
        lines: list[tuple[_StyledCluster, ...]] = []
        current: list[_StyledCluster] = []
        for cluster in clusters:
            current.append(cluster)
            while len(current) > 1:
                limit = first_width if not lines else following_width
                if sum(item.width for item in current) <= limit:
                    break
                split: int | None = None
                for index in range(1, len(current)):
                    if not JapaneseLineBreakingRules.can_break_between(
                        current[index - 1].text,
                        current[index].text,
                    ):
                        continue
                    if sum(item.width for item in current[:index]) <= limit:
                        split = index
                if split is None:
                    split = len(current) - 1
                lines.append(tuple(current[:split]))
                current = current[split:]
        if current:
            lines.append(tuple(current))
        return tuple(lines)

    @staticmethod
    def _styled_justification_slots(
        clusters: tuple[_StyledCluster, ...],
    ) -> tuple[int, ...]:
        spaces = tuple(
            index for index, cluster in enumerate(clusters[:-1]) if cluster.text.isspace()
        )
        if spaces:
            return spaces
        return tuple(
            index
            for index in range(len(clusters) - 1)
            if JapaneseLineBreakingRules.can_break_between(
                clusters[index].text,
                clusters[index + 1].text,
            )
        )

    @staticmethod
    def _append_styled_fragment(
        fragments: list[TextFragment],
        clusters: list[_StyledCluster],
        *,
        x: float,
        y: float,
        height: float,
        extra_after: float,
    ) -> float:
        run = clusters[0].run
        width = sum(cluster.width for cluster in clusters) + extra_after
        fragments.append(
            TextFragment(
                x=x,
                y=y,
                width=width,
                height=height,
                text="".join(cluster.text for cluster in clusters),
                font_name=run.font_name,
                font_path=run.font_path,
                font_size=run.font_size,
                bold=run.bold,
                italic=run.italic,
                underline=run.underline,
                strike=run.strike,
                color=run.color,
                highlight=run.highlight,
                character_spacing=run.character_spacing,
                baseline_shift=NativeLayoutEngine._baseline_shift(run),
                source_start=clusters[0].start,
                source_end=clusters[-1].end,
            )
        )
        return x + width

    def _wrapped_text_paragraph(
        self,
        paragraph: ResolvedParagraphModel,
        *,
        run: ResolvedRunModel,
        body: PageRegion,
        doc_grid_line_pitch: float | None = None,
    ) -> ParagraphBox:
        resolved_run = run
        font = ResolvedFont(
            family=resolved_run.font_name,
            path=resolved_run.font_path,
            postscript_name=resolved_run.font_name,
            source="resolved-run",
        )

        def measure(text: str) -> float:
            return self._text_measurer.measure(
                text,
                font,
                resolved_run.font_size,
                character_spacing=resolved_run.character_spacing,
            ).width

        first_offset = paragraph.first_line_indent - paragraph.hanging_indent
        first_x = body.x + paragraph.left_indent + first_offset
        available_width = max(
            0.01,
            body.width - paragraph.left_indent - paragraph.right_indent - first_offset,
        )
        broken_lines = LineBreaker().break_text(
            resolved_run.text,
            max_width=available_width,
            measure=measure,
        )
        metric = self._text_measurer.measure("Ag", font, resolved_run.font_size)
        natural_height = metric.height
        if paragraph.line_spacing_rule == "exact":
            line_height = paragraph.line_spacing
        elif paragraph.line_spacing_rule == "at_least":
            line_height = max(natural_height, paragraph.line_spacing)
        else:
            line_height = natural_height * paragraph.line_spacing
        line_height = self._snap_line_height_to_doc_grid(
            line_height,
            line_spacing_rule=paragraph.line_spacing_rule,
            doc_grid_line_pitch=doc_grid_line_pitch,
        )
        lines: list[LineBox] = []
        y = body.y + paragraph.space_before
        for index, broken in enumerate(broken_lines):
            line_x = first_x if index == 0 else body.x + paragraph.left_indent
            line_width = (
                available_width
                if index == 0
                else max(0.01, body.width - paragraph.left_indent - paragraph.right_indent)
            )
            if paragraph.alignment == "center":
                line_x += max(0, line_width - broken.width) / 2
            elif paragraph.alignment == "right":
                line_x += max(0, line_width - broken.width)
            justify = (
                paragraph.alignment == "justify"
                and index < len(broken_lines) - 1
                and "\t" not in broken.text
            )
            fragments = self._text_fragments(
                text=broken.text,
                source_start=broken.start,
                x=line_x,
                y=y,
                line_height=line_height,
                line_width=line_width,
                natural_width=broken.width,
                run=resolved_run,
                measure=measure,
                justify=justify,
            )
            used_width = line_width if justify and len(fragments) > 1 else broken.width
            lines.append(
                LineBox(
                    x=line_x,
                    y=y,
                    width=line_width,
                    height=line_height,
                    used_width=used_width,
                    ascent=metric.ascent,
                    descent=metric.descent,
                    baseline=metric.ascent,
                    fragments=fragments,
                    source_start=broken.start,
                    source_end=broken.end,
                )
            )
            y += line_height
        return ParagraphBox(
            x=body.x + paragraph.left_indent,
            y=body.y,
            width=max(0, body.width - paragraph.left_indent - paragraph.right_indent),
            height=paragraph.space_before + len(lines) * line_height + paragraph.space_after,
            lines=tuple(lines),
            paragraph_index=paragraph.source_index,
            source_start=0,
            source_end=len(resolved_run.text),
        )

    @staticmethod
    def _text_fragments(
        *,
        text: str,
        source_start: int,
        x: float,
        y: float,
        line_height: float,
        line_width: float,
        natural_width: float,
        run: ResolvedRunModel,
        measure: Callable[[str], float],
        justify: bool,
    ) -> tuple[TextFragment, ...]:
        if not justify:
            return (
                TextFragment(
                    x=x,
                    y=y,
                    width=natural_width,
                    height=line_height,
                    text=text,
                    font_name=run.font_name,
                    font_path=run.font_path,
                    font_size=run.font_size,
                    bold=run.bold,
                    italic=run.italic,
                    underline=run.underline,
                    strike=run.strike,
                    color=run.color,
                    highlight=run.highlight,
                    character_spacing=run.character_spacing,
                    baseline_shift=NativeLayoutEngine._baseline_shift(run),
                    source_start=source_start,
                    source_end=source_start + len(text),
                ),
            )

        clusters = UnicodeText.indexed_grapheme_clusters(text)
        space_slots = tuple(
            index for index, cluster in enumerate(clusters[:-1]) if cluster.text.isspace()
        )
        slots = space_slots or tuple(
            index
            for index in range(len(clusters) - 1)
            if JapaneseLineBreakingRules.can_break_between(
                clusters[index].text, clusters[index + 1].text
            )
        )
        if not slots:
            return NativeLayoutEngine._text_fragments(
                text=text,
                source_start=source_start,
                x=x,
                y=y,
                line_height=line_height,
                line_width=line_width,
                natural_width=natural_width,
                run=run,
                measure=measure,
                justify=False,
            )
        extra = max(0.0, line_width - natural_width) / len(slots)
        fragments: list[TextFragment] = []
        cursor = x
        for index, cluster in enumerate(clusters):
            width = float(measure(cluster.text))
            if index in slots:
                width += extra
            fragments.append(
                TextFragment(
                    x=cursor,
                    y=y,
                    width=width,
                    height=line_height,
                    text=cluster.text,
                    font_name=run.font_name,
                    font_path=run.font_path,
                    font_size=run.font_size,
                    bold=run.bold,
                    italic=run.italic,
                    underline=run.underline,
                    strike=run.strike,
                    color=run.color,
                    highlight=run.highlight,
                    character_spacing=run.character_spacing,
                    baseline_shift=NativeLayoutEngine._baseline_shift(run),
                    source_start=source_start + cluster.start,
                    source_end=source_start + cluster.end,
                )
            )
            cursor += width
        return tuple(fragments)

    @staticmethod
    def _baseline_shift(run: ResolvedRunModel) -> float:
        if run.vertical_align == "superscript":
            return run.font_size * 0.33
        if run.vertical_align == "subscript":
            return -run.font_size * 0.2
        return 0.0
