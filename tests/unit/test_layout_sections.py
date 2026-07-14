from __future__ import annotations

from docxpdf_native.abstractions import TextMeasurer
from docxpdf_native.layout.engine import NativeLayoutEngine
from docxpdf_native.models import (
    ConversionOptions,
    ResolvedDocumentModel,
    ResolvedFont,
    ResolvedParagraphModel,
    ResolvedRunModel,
    ResolvedSectionModel,
    SectionModel,
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
        return TextMeasurement(
            width=len(text) * font_size * 0.5,
            ascent=font_size * 0.8,
            descent=font_size * 0.2,
        )


def _paragraph(text: str) -> ResolvedParagraphModel:
    return ResolvedParagraphModel(runs=(ResolvedRunModel(text=text),))


def test_section_page_number_start_is_unspecified_by_default() -> None:
    assert SectionModel().page_number_start is None
    assert ResolvedSectionModel().page_number_start is None


def test_section_page_numbers_continue_when_next_section_has_no_start() -> None:
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(blocks=(_paragraph("first"),), page_number_start=7),
            ResolvedSectionModel(blocks=(_paragraph("second"),), page_number_start=None),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert [page.section_page_number for page in layout.pages] == [7, 8]


def test_continuous_section_uses_remaining_space_on_current_page() -> None:
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(blocks=(_paragraph("first"),)),
            ResolvedSectionModel(
                blocks=(_paragraph("second"),),
                section_break="continuous",
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 1
    assert [block.kind for block in layout.pages[0].body] == ["paragraph", "paragraph"]


def test_continuous_section_that_does_not_fit_continues_page_numbering() -> None:
    geometry = {
        "page_width": 100,
        "page_height": 40,
        "margin_left": 10,
        "margin_right": 10,
        "margin_top": 10,
        "margin_bottom": 10,
    }
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(blocks=(_paragraph("first"),), **geometry),
            ResolvedSectionModel(
                blocks=(_paragraph("second"),),
                section_break="continuous",
                **geometry,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 2
    assert [page.section_page_number for page in layout.pages] == [1, 2]


def test_odd_page_section_inserts_blank_page_when_needed() -> None:
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(blocks=(_paragraph("first"),)),
            ResolvedSectionModel(
                blocks=(_paragraph("third"),),
                section_break="odd_page",
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    assert layout.page_count == 3
    assert layout.pages[1].body == ()
    assert layout.pages[2].number == 3
