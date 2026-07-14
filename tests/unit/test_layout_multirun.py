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


def test_multiple_styled_runs_wrap_and_preserve_run_formatting() -> None:
    paragraph = ResolvedParagraphModel(
        runs=(
            ResolvedRunModel(
                text="one ",
                font_name="Helvetica",
                font_size=12,
                bold=True,
            ),
            ResolvedRunModel(
                text="two three",
                font_name="Helvetica",
                font_size=12,
                italic=True,
            ),
        ),
    )
    document = ResolvedDocumentModel(
        sections=(
            ResolvedSectionModel(
                blocks=(paragraph,),
                page_width=80,
                page_height=200,
                margin_left=10,
                margin_right=10,
                margin_top=10,
                margin_bottom=10,
            ),
        ),
    )

    layout = NativeLayoutEngine(_FixedTextMeasurer()).layout(document, options=ConversionOptions())

    lines = layout.pages[0].body[0].lines  # type: ignore[union-attr]
    assert ["".join(fragment.text for fragment in line.fragments) for line in lines] == [
        "one two ",
        "three",
    ]
    assert lines[0].fragments[0].bold is True  # type: ignore[union-attr]
    assert lines[0].fragments[-1].italic is True  # type: ignore[union-attr]
