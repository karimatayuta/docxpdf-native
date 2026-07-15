from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from pydantic import ValidationError

from docxpdf_native.models import (
    BlockBox,
    CellBox,
    ConversionOptions,
    ConversionResult,
    ConversionWarning,
    DocumentModel,
    FontConfiguration,
    FontSubstitution,
    ImageBox,
    ImageModel,
    LayoutDocument,
    LineBox,
    PageModel,
    ParagraphBox,
    ParagraphModel,
    ParagraphProperties,
    ResolvedDocumentModel,
    ResolvedParagraphModel,
    ResolvedRunModel,
    ResourceLimits,
    RunModel,
    RunProperties,
    SectionModel,
    StyleModel,
    TableBox,
    TableCellModel,
    TableModel,
    TableRowModel,
    TextFragment,
    UnsupportedFeature,
)


def test_document_models_represent_mixed_document_content() -> None:
    run = RunModel(text="日本語", properties=RunProperties(east_asia_font="Noto Sans CJK JP"))
    paragraph = ParagraphModel(
        runs=(run,),
        properties=ParagraphProperties(alignment="center"),
        source_index=2,
    )
    cell = TableCellModel(paragraphs=(paragraph,), grid_span=2)
    table = TableModel(rows=(TableRowModel(cells=(cell,), repeat_header=True),))
    image = ImageModel(
        relationship_id="rId1",
        part_name="word/media/image.png",
        content_type="image/png",
        data=b"png",
        width=72,
        height=36,
    )
    image_paragraph = ParagraphModel(runs=(RunModel(image=image),))
    section = SectionModel(blocks=(paragraph, table, image_paragraph), orientation="landscape")

    document = DocumentModel(sections=(section,))

    assert document.sections[0].blocks[0].runs[0].text == "日本語"
    assert document.sections[0].blocks[1].rows[0].cells[0].grid_span == 2
    assert document.sections[0].blocks[2].runs[0].image.width == 72


def test_document_models_are_frozen() -> None:
    run = RunModel(text="fixed")

    with pytest.raises(ValidationError):
        run.text = "changed"  # type: ignore[misc]


def test_style_model_supports_inheritance_and_direct_properties() -> None:
    style = StyleModel(
        style_id="Heading1",
        style_type="paragraph",
        name="Heading 1",
        based_on="Normal",
        next_style="Normal",
        link="Heading1Char",
        paragraph=ParagraphProperties(keep_next=True, space_before=12),
        run=RunProperties(bold=True, font_size=16),
    )

    assert style.based_on == "Normal"
    assert style.paragraph.keep_next is True
    assert style.run.font_size == 16


def test_resource_limits_reject_non_positive_values() -> None:
    with pytest.raises(ValidationError):
        ResourceLimits(max_pages=0)


def test_conversion_options_default_to_lenient_and_accept_font_paths() -> None:
    options = ConversionOptions(
        font_configuration=FontConfiguration(
            font_directories=(Path("fonts"),),
            registered_fonts={"Example": Path("fonts/example.ttf")},
            substitutions={"Missing": "Example"},
        )
    )

    assert options.strict is False
    assert options.deterministic is True
    assert options.font_configuration.font_directories == (Path("fonts"),)
    assert options.resource_limits.max_pages > 0


def test_resolved_models_store_effective_formatting() -> None:
    run = ResolvedRunModel(
        text="abc",
        font_name="Example",
        font_path=Path("example.ttf"),
        font_size=12,
        bold=True,
    )
    paragraph = ResolvedParagraphModel(
        runs=(run,),
        alignment="right",
        left_indent=18,
        source_index=4,
    )
    document = ResolvedDocumentModel(sections=(SectionModel(),), paragraphs=(paragraph,))

    assert document.paragraphs[0].runs[0].font_path == Path("example.ttf")
    assert document.paragraphs[0].alignment == "right"


def test_layout_models_serialize_independently_from_pdf() -> None:
    fragment = TextFragment(
        x=72,
        y=72,
        width=24,
        height=12,
        text="test",
        font_name="Example",
        font_size=10,
        source_start=0,
        source_end=4,
    )
    line = LineBox(
        x=72,
        y=72,
        width=100,
        height=14,
        used_width=24,
        ascent=9,
        descent=3,
        baseline=9,
        fragments=(fragment,),
        source_start=0,
        source_end=4,
    )
    paragraph = ParagraphBox(
        x=72,
        y=72,
        width=451,
        height=14,
        lines=(line,),
        source_start=0,
        source_end=4,
    )
    page = PageModel(number=1, width=595.28, height=841.89, body=(paragraph,))
    layout = LayoutDocument(pages=(page,))

    payload = layout.model_dump_json()

    assert '"text":"test"' in payload
    assert layout.page_count == 1


def test_layout_box_types_cover_images_tables_and_cells() -> None:
    image = ImageBox(
        x=10,
        y=20,
        width=30,
        height=40,
        image_data=b"image",
        content_type="image/png",
    )
    cell = CellBox(x=0, y=0, width=50, height=20, row_index=0, column_index=0)
    table = TableBox(x=0, y=0, width=50, height=20, cells=(cell,))
    generic = BlockBox(x=0, y=0, width=1, height=1)

    assert image.width == 30
    assert table.cells[0].column_index == 0
    assert generic.source_start == 0


def test_result_models_capture_warnings_features_and_substitutions() -> None:
    feature = UnsupportedFeature(
        name="text_box",
        part="word/document.xml",
        element="w:txbxContent",
        location="/w:document/w:body/w:p[1]",
        status="unsupported",
        workaround="Use an ordinary paragraph.",
    )
    warning = ConversionWarning(
        code="unsupported_feature",
        message="Text box was omitted.",
        feature=feature,
    )
    substitution = FontSubstitution(
        requested_font="Missing",
        selected_font="Fallback",
        reason="configured substitution",
        paragraph_index=1,
        run_index=2,
        east_asia=True,
    )
    result = ConversionResult(
        page_count=2,
        warnings=(warning,),
        unsupported_features=(feature,),
        font_substitutions=(substitution,),
        pdf_bytes=b"%PDF",
    )

    assert result.warnings[0].feature.name == "text_box"
    assert result.font_substitutions[0].east_asia is True
    assert result.pdf_bytes == b"%PDF"


def test_result_json_excludes_binary_pdf_payload() -> None:
    result = ConversionResult(page_count=1, pdf_bytes=BytesIO(b"%PDF").getvalue())

    assert "pdf_bytes" not in result.model_dump_json()
