from __future__ import annotations

from docxpdf_native.abstractions import FontResolver
from docxpdf_native.models import (
    DocumentDefaults,
    DocumentModel,
    NumberingDefinition,
    NumberingLevel,
    ParagraphModel,
    ParagraphProperties,
    ResolvedFont,
    RunModel,
    RunProperties,
    SectionModel,
    StyleModel,
    TableCellModel,
    TableModel,
    TableRowModel,
)
from docxpdf_native.resolver import NativeStyleResolver


class _RecordingFontResolver(FontResolver):
    def __init__(self) -> None:
        self.calls: list[tuple[str, bool]] = []

    def resolve(
        self,
        font_name: str,
        *,
        east_asia: bool = False,
        paragraph_index: int | None = None,
        run_index: int | None = None,
    ) -> ResolvedFont:
        self.calls.append((font_name, east_asia))
        return ResolvedFont(
            family=font_name,
            postscript_name=font_name,
            source="test",
        )


def test_native_resolver_splits_mixed_script_run_between_declared_fonts() -> None:
    document = DocumentModel(
        defaults=DocumentDefaults(
            run=RunProperties(
                ascii_font="Latin Face",
                high_ansi_font="Latin Face",
                east_asia_font="Japanese Face",
            ),
        ),
        sections=(
            SectionModel(
                blocks=(ParagraphModel(runs=(RunModel(text="ABC日本語123"),)),),
            ),
        ),
    )
    font_resolver = _RecordingFontResolver()

    resolved = NativeStyleResolver(font_resolver).resolve(document)

    runs = resolved.paragraphs[0].runs
    assert [(run.text, run.font_name) for run in runs] == [
        ("ABC", "Latin Face"),
        ("日本語", "Japanese Face"),
        ("123", "Latin Face"),
    ]
    assert font_resolver.calls == [
        ("Latin Face", False),
        ("Japanese Face", True),
        ("Latin Face", False),
    ]


def test_native_resolver_inherits_indent_from_numbering_level() -> None:
    document = DocumentModel(
        numbering=(
            NumberingDefinition(
                numbering_id=1,
                levels=(
                    NumberingLevel(
                        level=0,
                        text="%1.",
                        paragraph=ParagraphProperties(left_indent=36.0, hanging_indent=18.0),
                    ),
                ),
            ),
        ),
        sections=(
            SectionModel(
                blocks=(
                    ParagraphModel(
                        properties=ParagraphProperties(numbering_id=1, numbering_level=0),
                        runs=(RunModel(text="item"),),
                    ),
                ),
            ),
        ),
    )

    resolved = NativeStyleResolver(_RecordingFontResolver()).resolve(document)

    paragraph = resolved.paragraphs[0]
    assert paragraph.left_indent == 36.0
    assert paragraph.hanging_indent == 18.0


def test_native_resolver_direct_indent_overrides_numbering_level() -> None:
    document = DocumentModel(
        numbering=(
            NumberingDefinition(
                numbering_id=1,
                levels=(
                    NumberingLevel(
                        level=0,
                        text="%1.",
                        paragraph=ParagraphProperties(left_indent=36.0, hanging_indent=18.0),
                    ),
                ),
            ),
        ),
        sections=(
            SectionModel(
                blocks=(
                    ParagraphModel(
                        properties=ParagraphProperties(
                            numbering_id=1,
                            numbering_level=0,
                            left_indent=72.0,
                        ),
                        runs=(RunModel(text="item"),),
                    ),
                ),
            ),
        ),
    )

    resolved = NativeStyleResolver(_RecordingFontResolver()).resolve(document)

    paragraph = resolved.paragraphs[0]
    # Direct formatting (left_indent=72.0) wins over the numbering level.
    assert paragraph.left_indent == 72.0
    # hanging_indent was never set directly, so the numbering level's own
    # value still shows through.
    assert paragraph.hanging_indent == 18.0


def test_native_resolver_numbering_level_indent_overrides_paragraph_style() -> None:
    document = DocumentModel(
        styles=(
            StyleModel(
                style_id="ListParagraph",
                style_type="paragraph",
                paragraph=ParagraphProperties(left_indent=10.0, hanging_indent=5.0),
            ),
        ),
        numbering=(
            NumberingDefinition(
                numbering_id=1,
                levels=(
                    NumberingLevel(
                        level=0,
                        text="%1.",
                        paragraph=ParagraphProperties(left_indent=36.0, hanging_indent=18.0),
                    ),
                ),
            ),
        ),
        sections=(
            SectionModel(
                blocks=(
                    ParagraphModel(
                        style_id="ListParagraph",
                        properties=ParagraphProperties(numbering_id=1, numbering_level=0),
                        runs=(RunModel(text="item"),),
                    ),
                ),
            ),
        ),
    )

    resolved = NativeStyleResolver(_RecordingFontResolver()).resolve(document)

    paragraph = resolved.paragraphs[0]
    # The numbering level's indent overrides the paragraph style's, even
    # though the style is "closer" in the traditional inheritance sense.
    assert paragraph.left_indent == 36.0
    assert paragraph.hanging_indent == 18.0


def test_native_resolver_paragraph_without_numbering_is_unaffected() -> None:
    document = DocumentModel(
        numbering=(
            NumberingDefinition(
                numbering_id=1,
                levels=(
                    NumberingLevel(
                        level=0,
                        text="%1.",
                        paragraph=ParagraphProperties(left_indent=36.0, hanging_indent=18.0),
                    ),
                ),
            ),
        ),
        sections=(
            SectionModel(
                blocks=(ParagraphModel(runs=(RunModel(text="plain"),)),),
            ),
        ),
    )

    resolved = NativeStyleResolver(_RecordingFontResolver()).resolve(document)

    paragraph = resolved.paragraphs[0]
    assert paragraph.left_indent == 0.0
    assert paragraph.hanging_indent == 0.0


def test_native_resolver_bakes_numbering_label_into_table_cell_paragraph_text() -> None:
    # Table cells are resolved once here (to compute style inheritance, the
    # numbering label included) and then converted back to a raw
    # ParagraphModel via _resolved_to_raw, because NativeLayoutEngine
    # re-resolves table cell paragraphs later without access to the
    # document's numbering definitions. If the already-resolved label isn't
    # baked into the run text at this point, it is silently dropped and a
    # bulleted/numbered cell renders with no marker at all.
    document = DocumentModel(
        numbering=(
            NumberingDefinition(
                numbering_id=1,
                levels=(NumberingLevel(level=0, number_format="bullet", text=""),),
            ),
        ),
        sections=(
            SectionModel(
                blocks=(
                    TableModel(
                        rows=(
                            TableRowModel(
                                cells=(
                                    TableCellModel(
                                        paragraphs=(
                                            ParagraphModel(
                                                properties=ParagraphProperties(
                                                    numbering_id=1, numbering_level=0
                                                ),
                                                runs=(RunModel(text="item"),),
                                            ),
                                        ),
                                    ),
                                ),
                            ),
                        ),
                    ),
                ),
            ),
        ),
    )

    resolved = NativeStyleResolver(_RecordingFontResolver()).resolve(document)

    table = resolved.sections[0].blocks[0]
    cell_paragraph = table.rows[0].cells[0].paragraphs[0]
    assert cell_paragraph.runs[0].text == " item"
