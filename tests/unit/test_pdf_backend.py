from __future__ import annotations

from base64 import b64decode
from io import BytesIO
from pathlib import Path

import pytest
import reportlab
from pypdf import PdfReader
from pypdf.generic import ContentStream
from reportlab.lib import utils as reportlab_utils

from docxpdf_native.exceptions import PdfGenerationError
from docxpdf_native.models import (
    BorderModel,
    CellBox,
    ImageBox,
    LayoutDocument,
    LineBox,
    PageModel,
    ParagraphBox,
    PlaceholderBox,
    TableBorders,
    TableBox,
    TextFragment,
)
from docxpdf_native.pdf import ReportLabPdfBackend

PNG_1X1 = b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC"
)
JPEG_1X1 = b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRof"
    "Hh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwh"
    "MjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAAR"
    "CAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAA"
    "AgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkK"
    "FhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWG"
    "h4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl"
    "5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREA"
    "AgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYk"
    "NOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOE"
    "hYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk"
    "5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDi6KKK+ZP3E//Z"
)


def _fragment_layout(fragment: TextFragment) -> LayoutDocument:
    line = LineBox(
        x=72,
        y=72,
        width=451,
        height=14,
        used_width=fragment.width,
        ascent=10,
        descent=2,
        baseline=10,
        fragments=(fragment,),
        source_start=0,
        source_end=fragment.source_end,
    )
    paragraph = ParagraphBox(
        x=72,
        y=72,
        width=451,
        height=14,
        lines=(line,),
        source_start=0,
        source_end=fragment.source_end,
    )
    return LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(paragraph,)),)
    )


def _text_layout(text: str = "Hello PDF") -> LayoutDocument:
    return _fragment_layout(
        TextFragment(
            x=72,
            y=72,
            width=55,
            height=12,
            text=text,
            font_name="Helvetica",
            font_size=12,
            source_start=0,
            source_end=len(text),
        )
    )


def _image_layout(image: ImageBox) -> LayoutDocument:
    line = LineBox(
        x=image.x,
        y=image.y,
        width=image.width,
        height=image.height,
        used_width=image.width,
        ascent=image.height,
        descent=0,
        baseline=image.height,
        fragments=(image,),
    )
    paragraph = ParagraphBox(
        x=image.x,
        y=image.y,
        width=image.width,
        height=image.height,
        lines=(line,),
    )
    return LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(paragraph,)),)
    )


def test_reportlab_backend_emits_extractable_text() -> None:
    payload = ReportLabPdfBackend(deterministic=True).render(_text_layout())

    extracted = PdfReader(BytesIO(payload)).pages[0].extract_text()

    assert extracted.strip() == "Hello PDF"


def test_reportlab_backend_emits_character_spacing_operator() -> None:
    fragment = (
        _text_layout()
        .pages[0]
        .body[0]
        .lines[0]
        .fragments[0]
        .model_copy(update={"character_spacing": 1.5})
    )

    rendered = ReportLabPdfBackend(deterministic=True).render(_fragment_layout(fragment))

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    spacing = [float(values[0]) for values, operator in operations if operator == b"Tc"]
    assert spacing == [1.5]


def test_reportlab_backend_emits_identical_bytes_in_deterministic_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    timestamps = iter((1_000_000_000.0, 2_000_000_000.0))
    monkeypatch.setattr(reportlab_utils.time, "time", lambda: next(timestamps))
    backend = ReportLabPdfBackend(deterministic=True)

    first = backend.render(_text_layout())
    second = backend.render(_text_layout())

    assert first == second


def test_reportlab_backend_uses_layout_page_size() -> None:
    layout = _text_layout()
    page = layout.pages[0].model_copy(update={"width": 300.0, "height": 500.0})

    rendered = ReportLabPdfBackend().render(layout.model_copy(update={"pages": (page,)}))

    media_box = PdfReader(BytesIO(rendered)).pages[0].mediabox
    assert (float(media_box.width), float(media_box.height)) == (300.0, 500.0)


def test_reportlab_backend_renders_each_page_with_its_own_size() -> None:
    first_page = _text_layout("Page one").pages[0]
    second_page = (
        _text_layout("Page two")
        .pages[0]
        .model_copy(update={"number": 2, "width": 400.0, "height": 600.0})
    )
    layout = LayoutDocument(pages=(first_page, second_page))

    rendered = ReportLabPdfBackend().render(layout)

    pages = PdfReader(BytesIO(rendered)).pages
    observed = [
        (
            page.extract_text().strip(),
            float(page.mediabox.width),
            float(page.mediabox.height),
        )
        for page in pages
    ]
    assert observed == [
        ("Page one", 595.28, 841.89),
        ("Page two", 400.0, 600.0),
    ]


def test_reportlab_backend_draws_text_highlight_from_top_left_coordinates() -> None:
    fragment = (
        _text_layout()
        .pages[0]
        .body[0]
        .lines[0]
        .fragments[0]
        .model_copy(update={"highlight": "FFFF00"})
    )

    rendered = ReportLabPdfBackend().render(_fragment_layout(fragment))

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    fill_colors = [
        tuple(float(value) for value in values) for values, op in operations if op == b"rg"
    ]
    rectangles = [
        tuple(round(float(value), 2) for value in values)
        for values, op in operations
        if op == b"re"
    ]
    assert ((1.0, 1.0, 0.0) in fill_colors, rectangles) == (
        True,
        [(72.0, 757.89, 55.0, 12.0)],
    )


def test_reportlab_backend_draws_underline_at_fragment_baseline() -> None:
    fragment = (
        _text_layout().pages[0].body[0].lines[0].fragments[0].model_copy(update={"underline": True})
    )

    rendered = ReportLabPdfBackend().render(_fragment_layout(fragment))

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    moves = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"m"
    ]
    lines = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"l"
    ]
    assert list(zip(moves, lines, strict=True)) == [
        ((72.0, 758.69), (127.0, 758.69)),
    ]


def test_reportlab_backend_draws_strikethrough_above_fragment_baseline() -> None:
    fragment = (
        _text_layout().pages[0].body[0].lines[0].fragments[0].model_copy(update={"strike": True})
    )

    rendered = ReportLabPdfBackend().render(_fragment_layout(fragment))

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    moves = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"m"
    ]
    lines = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"l"
    ]
    assert list(zip(moves, lines, strict=True)) == [
        ((72.0, 763.49), (127.0, 763.49)),
    ]


def test_reportlab_backend_draws_table_cell_background_without_implicit_border() -> None:
    cell = CellBox(
        x=72,
        y=72,
        width=200,
        height=40,
        row_index=0,
        column_index=0,
        background_color="E6E6E6",
    )
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=40,
        cells=(cell,),
        column_widths=(200,),
        row_heights=(40,),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    fill_colors = [
        tuple(round(float(value), 4) for value in values)
        for values, op in operations
        if op == b"rg"
    ]
    stroke_colors = [
        tuple(float(value) for value in values) for values, op in operations if op == b"RG"
    ]
    rectangles = [
        tuple(round(float(value), 2) for value in values)
        for values, op in operations
        if op == b"re"
    ]
    assert (fill_colors, stroke_colors, rectangles) == (
        [(0.902, 0.902, 0.902)],
        [],
        [(72.0, 729.89, 200.0, 40.0)],
    )


def test_reportlab_backend_draws_configured_cell_border_styles() -> None:
    cell = CellBox(
        x=72,
        y=72,
        width=200,
        height=40,
        row_index=0,
        column_index=0,
        borders=TableBorders(
            top=BorderModel(style="dashed", width=2, color="FF0000"),
            right=BorderModel(style="single", width=1, color="00FF00"),
            bottom=BorderModel(style="dotted", width=3, color="0000FF"),
            left=BorderModel(style="nil", width=4, color="000000"),
        ),
    )
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=40,
        cells=(cell,),
        column_widths=(200,),
        row_heights=(40,),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    colors = [tuple(float(value) for value in values) for values, op in operations if op == b"RG"]
    widths = [float(values[0]) for values, op in operations if op == b"w"]
    dashes = [
        (tuple(float(value) for value in values[0]), float(values[1]))
        for values, op in operations
        if op == b"d"
    ]
    moves = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"m"
    ]
    lines = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"l"
    ]
    assert (colors, widths, dashes, list(zip(moves, lines, strict=True))) == (
        [(1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0)],
        [2.0, 1.0, 3.0],
        [((6.0, 4.0), 0.0), ((), 0.0), ((3.0, 6.0), 0.0)],
        [
            ((72.0, 769.89), (272.0, 769.89)),
            ((272.0, 769.89), (272.0, 729.89)),
            ((72.0, 729.89), (272.0, 729.89)),
        ],
    )


def test_reportlab_backend_draws_table_outer_and_inside_borders() -> None:
    cells = (
        CellBox(
            x=72,
            y=72,
            width=100,
            height=40,
            row_index=0,
            column_index=0,
        ),
        CellBox(
            x=172,
            y=72,
            width=100,
            height=40,
            row_index=0,
            column_index=1,
        ),
    )
    outer = BorderModel(style="single", width=2, color="FF0000")
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=40,
        cells=cells,
        column_widths=(100, 100),
        row_heights=(40,),
        borders=TableBorders(
            top=outer,
            right=outer,
            bottom=outer,
            left=outer,
            inside_vertical=BorderModel(style="dashed", width=1, color="0000FF"),
        ),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    colors = [tuple(float(value) for value in values) for values, op in operations if op == b"RG"]
    moves = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"m"
    ]
    lines = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"l"
    ]
    assert (colors, list(zip(moves, lines, strict=True))) == (
        [
            (1.0, 0.0, 0.0),
            (1.0, 0.0, 0.0),
            (1.0, 0.0, 0.0),
            (1.0, 0.0, 0.0),
            (0.0, 0.0, 1.0),
        ],
        [
            ((72.0, 769.89), (272.0, 769.89)),
            ((272.0, 769.89), (272.0, 729.89)),
            ((72.0, 729.89), (272.0, 729.89)),
            ((72.0, 769.89), (72.0, 729.89)),
            ((172.0, 769.89), (172.0, 729.89)),
        ],
    )


def test_reportlab_backend_draws_inside_horizontal_table_border() -> None:
    cells = (
        CellBox(
            x=72,
            y=72,
            width=200,
            height=40,
            row_index=0,
            column_index=0,
        ),
        CellBox(
            x=72,
            y=112,
            width=200,
            height=40,
            row_index=1,
            column_index=0,
        ),
    )
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=80,
        cells=cells,
        column_widths=(200,),
        row_heights=(40, 40),
        borders=TableBorders(
            inside_horizontal=BorderModel(style="single", width=1, color="00FF00")
        ),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    moves = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"m"
    ]
    lines = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"l"
    ]
    assert list(zip(moves, lines, strict=True)) == [
        ((72.0, 729.89), (272.0, 729.89)),
    ]


def test_reportlab_backend_draws_double_border_as_two_parallel_lines() -> None:
    cell = CellBox(
        x=72,
        y=72,
        width=200,
        height=40,
        row_index=0,
        column_index=0,
        borders=TableBorders(top=BorderModel(style="double", width=3, color="000000")),
    )
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=40,
        cells=(cell,),
        column_widths=(200,),
        row_heights=(40,),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    moves = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"m"
    ]
    lines = [
        tuple(round(float(value), 2) for value in values) for values, op in operations if op == b"l"
    ]
    assert list(zip(moves, lines, strict=True)) == [
        ((72.0, 768.89), (272.0, 768.89)),
        ((72.0, 770.89), (272.0, 770.89)),
    ]


def test_reportlab_backend_draws_paragraphs_inside_table_cells() -> None:
    paragraph = _text_layout("Cell text").pages[0].body[0]
    cell = CellBox(
        x=72,
        y=72,
        width=200,
        height=40,
        row_index=0,
        column_index=0,
        blocks=(paragraph,),
    )
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=40,
        cells=(cell,),
        column_widths=(200,),
        row_heights=(40,),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    extracted = PdfReader(BytesIO(rendered)).pages[0].extract_text()
    assert extracted.strip() == "Cell text"


def test_reportlab_backend_draws_images_inside_table_cells() -> None:
    image = ImageBox(
        x=80,
        y=80,
        width=20,
        height=10,
        image_data=PNG_1X1,
        content_type="image/png",
    )
    cell = CellBox(
        x=72,
        y=72,
        width=200,
        height=40,
        row_index=0,
        column_index=0,
        blocks=(image,),
    )
    table = TableBox(
        x=72,
        y=72,
        width=200,
        height=40,
        cells=(cell,),
        column_widths=(200,),
        row_heights=(40,),
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(table,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    assert len([values for values, op in operations if op == b"Do"]) == 1


def test_reportlab_backend_draws_png_image_box_at_top_left_coordinates() -> None:
    image = ImageBox(
        x=90,
        y=100,
        width=20,
        height=10,
        image_data=PNG_1X1,
        content_type="image/png",
        part_name="word/media/image1.png",
    )

    rendered = ReportLabPdfBackend().render(_image_layout(image))

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    image_draws = [values for values, op in operations if op == b"Do"]
    transforms = [
        tuple(round(float(value), 2) for value in values)
        for values, op in operations
        if op == b"cm" and tuple(values) != (1, 0, 0, 1, 0, 0)
    ]
    assert (len(image_draws), transforms) == (
        1,
        [(20.0, 0.0, 0.0, 10.0, 90.0, 731.89)],
    )


def test_reportlab_backend_draws_jpeg_image_box() -> None:
    image = ImageBox(
        x=90,
        y=100,
        width=20,
        height=10,
        image_data=JPEG_1X1,
        content_type="image/jpeg",
        part_name="word/media/image1.jpeg",
    )

    rendered = ReportLabPdfBackend().render(_image_layout(image))

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    image_draws = [values for values, op in operations if op == b"Do"]
    assert len(image_draws) == 1


def test_reportlab_backend_draws_page_level_image_box() -> None:
    image = ImageBox(
        x=90,
        y=100,
        width=20,
        height=10,
        image_data=PNG_1X1,
        content_type="image/png",
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(image,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    assert len([values for values, op in operations if op == b"Do"]) == 1


def test_reportlab_backend_draws_page_level_placeholder_box_without_raising() -> None:
    placeholder = PlaceholderBox(
        x=90,
        y=100,
        width=72,
        height=36,
        label="[Embedded Object]",
    )
    layout = LayoutDocument(
        pages=(PageModel(number=1, width=595.28, height=841.89, body=(placeholder,)),)
    )

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    operations = ContentStream(reader.pages[0].get_contents(), reader).operations
    # The placeholder is drawn as a filled/stroked rectangle plus a short
    # label, never as an embedded image.
    assert any(op == b"re" for _, op in operations)
    assert not any(op == b"Do" for _, op in operations)
    assert "[Embedded Object]" in (reader.pages[0].extract_text() or "")


def test_reportlab_backend_draws_placeholder_box_inside_table_cells() -> None:
    placeholder = PlaceholderBox(x=8, y=8, width=20, height=10, label="[Image]")
    cell = CellBox(
        x=0, y=0, width=40, height=20, row_index=0, column_index=0, blocks=(placeholder,)
    )
    table = TableBox(
        x=0, y=0, width=40, height=20, cells=(cell,), column_widths=(40,), row_heights=(20,)
    )
    layout = LayoutDocument(pages=(PageModel(number=1, width=200, height=200, body=(table,)),))

    rendered = ReportLabPdfBackend().render(layout)

    reader = PdfReader(BytesIO(rendered))
    assert "[Image]" in (reader.pages[0].extract_text() or "")


def test_reportlab_backend_writes_returned_bytes_to_path(tmp_path: Path) -> None:
    destination = tmp_path / "rendered.pdf"

    rendered = ReportLabPdfBackend().render(_text_layout(), destination)

    assert destination.read_bytes() == rendered


def test_reportlab_backend_writes_returned_bytes_to_binary_stream() -> None:
    destination = BytesIO()

    rendered = ReportLabPdfBackend().render(_text_layout(), destination)

    assert destination.getvalue() == rendered


def test_reportlab_backend_wraps_rendering_errors() -> None:
    image = ImageBox(
        x=90,
        y=100,
        width=20,
        height=10,
        image_data=b"not a PNG",
        content_type="image/png",
        part_name="word/media/broken.png",
    )

    with pytest.raises(PdfGenerationError, match="PDF generation failed") as caught:
        ReportLabPdfBackend().render(_image_layout(image))

    assert caught.value.cause is not None


def test_reportlab_backend_registers_injected_ttf_font() -> None:
    vera_path = Path(reportlab.__file__).parent / "fonts" / "Vera.ttf"
    fragment = _text_layout("Embedded font").pages[0].body[0].lines[0].fragments[0]
    fragment = fragment.model_copy(update={"font_name": "DocxPdfNativeTestVera"})

    rendered = ReportLabPdfBackend(font_paths={"DocxPdfNativeTestVera": vera_path}).render(
        _fragment_layout(fragment)
    )

    extracted = PdfReader(BytesIO(rendered)).pages[0].extract_text()
    assert extracted.strip() == "Embedded font"


@pytest.mark.parametrize(
    ("formatting", "expected_font"),
    [
        ({"bold": True}, "/Helvetica-Bold"),
        ({"italic": True}, "/Helvetica-Oblique"),
        ({"bold": True, "italic": True}, "/Helvetica-BoldOblique"),
        ({"font_name": "Times-Roman", "bold": True}, "/Times-Bold"),
        ({"font_name": "Times-Roman", "italic": True}, "/Times-Italic"),
        (
            {"font_name": "Times-Roman", "bold": True, "italic": True},
            "/Times-BoldItalic",
        ),
        ({"font_name": "Courier", "bold": True}, "/Courier-Bold"),
        ({"font_name": "Courier", "italic": True}, "/Courier-Oblique"),
        (
            {"font_name": "Courier", "bold": True, "italic": True},
            "/Courier-BoldOblique",
        ),
    ],
)
def test_reportlab_backend_selects_standard_font_face(
    formatting: dict[str, bool | str], expected_font: str
) -> None:
    fragment = _text_layout().pages[0].body[0].lines[0].fragments[0]
    fragment = fragment.model_copy(update=formatting)

    rendered = ReportLabPdfBackend().render(_fragment_layout(fragment))

    page = PdfReader(BytesIO(rendered)).pages[0]
    fonts = page["/Resources"]["/Font"]
    base_fonts = {str(font.get_object()["/BaseFont"]) for font in fonts.values()}
    assert expected_font in base_fonts
