from __future__ import annotations

from base64 import b64decode
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pytest
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from pypdf import PdfReader

from docxpdf_native import ConversionOptions, Converter, FontConfiguration
from docxpdf_native.exceptions import InvalidDocxError, PdfGenerationError, UnsupportedFeatureError
from docxpdf_native.models import ImageBox, ParagraphBox, TableBox

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
VML_NS = "urn:schemas-microsoft-com:vml"

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


def _minimal_docx(
    text: str = "Hello native PDF",
    *,
    extra_body: str = "",
    body_xml: str | None = None,
    ascii_font: str = "Helvetica",
    east_asia_font: str = "Helvetica",
    document_relationships: bytes | None = None,
    extra_parts: dict[str, bytes] | None = None,
    final_section_properties: str | None = None,
) -> bytes:
    content_types = b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
    root_rels = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    styles = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="{ascii_font}" w:hAnsi="{ascii_font}" w:eastAsia="{east_asia_font}"/>
    <w:sz w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>""".encode()
    body = (
        body_xml
        or f"""
    <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
    {extra_body}
    """
    )
    section_properties = (
        final_section_properties
        or """
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    """
    )
    document = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:wp="{WP_NS}"
  xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}">
  <w:body>
    {body}
    <w:sectPr>
      {section_properties}
    </w:sectPr>
  </w:body>
</w:document>""".encode()

    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "word/document.xml": document,
        "word/styles.xml": styles,
    }
    if document_relationships is not None:
        parts["word/_rels/document.xml.rels"] = document_relationships
    parts.update(extra_parts or {})
    output = BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    return output.getvalue()


def _table_xml(*rows: tuple[str, str]) -> str:
    row_xml = "".join(
        f"""<w:tr>
          <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
            <w:p><w:r><w:t>{left}</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
            <w:p><w:r><w:t>{right}</w:t></w:r></w:p></w:tc>
        </w:tr>"""
        for left, right in rows
    )
    return f"""<w:tbl>
      <w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="000000"/>
          <w:left w:val="single" w:sz="8" w:color="000000"/>
          <w:bottom w:val="single" w:sz="8" w:color="000000"/>
          <w:right w:val="single" w:sz="8" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="808080"/>
          <w:insideV w:val="single" w:sz="4" w:color="808080"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
      {row_xml}
    </w:tbl>"""


def _paginated_table_xml(*, data_rows: int) -> str:
    rows = [
        """<w:tr><w:trPr><w:trHeight w:val="480" w:hRule="exact"/><w:tblHeader/></w:trPr>
          <w:tc><w:p><w:r><w:t>repeated header</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>header value</w:t></w:r></w:p></w:tc>
        </w:tr>"""
    ]
    rows.extend(
        f"""<w:tr><w:trPr><w:trHeight w:val="480" w:hRule="exact"/><w:cantSplit/></w:trPr>
          <w:tc><w:p><w:r><w:t>row {index:02d}</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>value {index:02d}</w:t></w:r></w:p></w:tc>
        </w:tr>"""
        for index in range(1, data_rows + 1)
    )
    return f"""<w:tbl>
      <w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
      {"".join(rows)}
    </w:tbl>"""


def _merged_table_xml() -> str:
    return """<w:tbl>
      <w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
          <w:p><w:r><w:t>horizontal merge</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>vertical merge</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>right first</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>right second</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"""


def _inline_image_xml(relationship_id: str, *, width_emu: int, height_emu: int) -> str:
    return f"""<w:r><w:drawing><wp:inline>
      <wp:extent cx="{width_emu}" cy="{height_emu}"/>
      <wp:docPr id="1" name="Fixture image"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill>
        <a:blip r:embed="{relationship_id}"/>
      </pic:blipFill></pic:pic></a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r>"""


def _build_test_japanese_font(destination: Path, text: str) -> None:
    characters = tuple(dict.fromkeys(text))
    glyph_order = [".notdef", "space", *(f"uni{ord(char):04X}" for char in characters)]
    glyphs = {}
    metrics = {}
    for glyph_name in glyph_order:
        pen = TTGlyphPen(None)
        if glyph_name not in {"space"}:
            pen.moveTo((80, 80))
            pen.lineTo((920, 80))
            pen.lineTo((920, 920))
            pen.lineTo((80, 920))
            pen.closePath()
        glyphs[glyph_name] = pen.glyph()
        metrics[glyph_name] = (1000, 0)
    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder(glyph_order)
    builder.setupCharacterMap(
        {ord(char): f"uni{ord(char):04X}" for char in characters} | {32: "space"}
    )
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics(metrics)
    builder.setupHorizontalHeader(ascent=900, descent=-100)
    builder.setupNameTable(
        {
            "familyName": "DocxPdf Test Japanese",
            "styleName": "Regular",
            "uniqueFontIdentifier": "DocxPdfTestJapanese-Regular",
            "fullName": "DocxPdf Test Japanese Regular",
            "psName": "DocxPdfTestJapanese-Regular",
        }
    )
    builder.setupOS2(
        sTypoAscender=900,
        sTypoDescender=-100,
        usWinAscent=900,
        usWinDescent=100,
    )
    builder.setupPost()
    builder.setupMaxp()
    builder.save(destination)


def test_converter_returns_one_layout_page_for_minimal_docx() -> None:
    destination = BytesIO()

    result = Converter(ConversionOptions(strict=False)).convert(
        _minimal_docx(), destination=destination
    )

    assert result.page_count == 1


def test_converter_emits_readable_pdf_text() -> None:
    result = Converter(ConversionOptions(strict=False)).convert(_minimal_docx())

    extracted = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0].extract_text()

    assert extracted.strip() == "Hello native PDF"


def test_converter_honors_explicit_page_break() -> None:
    source = _minimal_docx(
        "before",
        extra_body="""
        <w:p>
          <w:r><w:br w:type="page"/></w:r>
          <w:r><w:t>after</w:t></w:r>
        </w:p>
        """,
    )

    result = Converter(ConversionOptions(strict=False)).convert(source)

    assert result.page_count == 2


def test_converter_rejects_text_box_in_strict_mode() -> None:
    source = _minimal_docx(
        "kept",
        extra_body="<w:p><w:txbxContent><w:p/></w:txbxContent></w:p>",
    )

    with pytest.raises(UnsupportedFeatureError) as caught:
        Converter(ConversionOptions(strict=True)).convert(source)

    assert caught.value.feature.name == "text_box"


def test_converter_records_text_box_warning_in_lenient_mode() -> None:
    source = _minimal_docx(
        "kept",
        extra_body="<w:p><w:txbxContent><w:p/></w:txbxContent></w:p>",
    )

    result = Converter(ConversionOptions(strict=False)).convert(source)

    assert result.warnings[0].feature.name == "text_box"  # type: ignore[union-attr]


def test_converter_is_binary_deterministic() -> None:
    source = _minimal_docx()
    converter = Converter(ConversionOptions(strict=False, deterministic=True))

    first = converter.convert_bytes(source)
    second = converter.convert_bytes(source)

    assert first == second


def test_converter_writes_path_and_rejects_same_input_output(tmp_path: Path) -> None:
    source = tmp_path / "input.docx"
    source.write_bytes(_minimal_docx())
    destination = tmp_path / "output.pdf"
    converter = Converter(ConversionOptions(strict=False))

    converter.convert(source, destination)

    assert destination.read_bytes().startswith(b"%PDF")
    with pytest.raises(InvalidDocxError):
        converter.convert(source, source)


def test_converter_rejects_symbolic_link_output(tmp_path: Path) -> None:
    actual = tmp_path / "actual.pdf"
    actual.write_bytes(b"keep")
    destination = tmp_path / "output.pdf"
    destination.symlink_to(actual)

    with pytest.raises(PdfGenerationError, match="symbolic link"):
        Converter(ConversionOptions(strict=False)).convert(_minimal_docx(), destination)


def test_converter_embeds_registered_japanese_font(tmp_path: Path) -> None:
    text = "日本語の文書"
    font_path = tmp_path / "test-japanese.ttf"
    _build_test_japanese_font(font_path, text)
    source = _minimal_docx(text, east_asia_font="Test Japanese")
    options = ConversionOptions(
        strict=True,
        font_configuration=FontConfiguration(
            registered_fonts={"Test Japanese": font_path},
        ),
    )

    result = Converter(options).convert(source)

    extracted = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0].extract_text()
    assert text in extracted


def test_converter_preserves_paragraph_table_paragraph_block_order() -> None:
    source = _minimal_docx(
        body_xml=f"""
        <w:p><w:r><w:t>before table</w:t></w:r></w:p>
        {_table_xml(("left cell", "right cell"))}
        <w:p><w:r><w:t>after table</w:t></w:r></w:p>
        """
    )
    converter = Converter(ConversionOptions(strict=False))

    converter.convert(source)

    layout = converter.last_layout
    assert layout is not None
    assert tuple(type(block) for block in layout.pages[0].body) == (
        ParagraphBox,
        TableBox,
        ParagraphBox,
    )


def test_converter_renders_fixed_table_text_into_readable_pdf() -> None:
    source = _minimal_docx(
        body_xml=_table_xml(
            ("row one left", "row one right"),
            ("row two left", "row two right"),
        )
    )
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    pdf = PdfReader(BytesIO(result.pdf_bytes or b""))
    extracted = tuple(line for line in pdf.pages[0].extract_text().splitlines() if line)
    layout = converter.last_layout
    assert layout is not None
    assert (
        result.page_count,
        tuple(type(block) for block in layout.pages[0].body),
        extracted,
    ) == (
        1,
        (TableBox,),
        ("row one left", "row one right", "row two left", "row two right"),
    )


def test_converter_renders_inline_png_and_jpeg_into_readable_pdf() -> None:
    relationships = f"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="{PACKAGE_REL_NS}">
      <Relationship Id="rIdPng" Type="{R_NS}/image" Target="media/image1.png"/>
      <Relationship Id="rIdJpeg" Type="{R_NS}/image" Target="media/image2.jpg"/>
    </Relationships>""".encode()
    source = _minimal_docx(
        body_xml=f"""
        <w:p>{_inline_image_xml("rIdPng", width_emu=914400, height_emu=914400)}</w:p>
        <w:p>{_inline_image_xml("rIdJpeg", width_emu=457200, height_emu=914400)}</w:p>
        """,
        document_relationships=relationships,
        extra_parts={
            "word/media/image1.png": PNG_1X1,
            "word/media/image2.jpg": JPEG_1X1,
        },
    )
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    pdf_page = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0]
    xobjects = pdf_page["/Resources"]["/XObject"].get_object()
    pdf_image_count = sum(
        item.get_object().get("/Subtype") == "/Image" for item in xobjects.values()
    )
    layout = converter.last_layout
    assert layout is not None
    layout_images = tuple(
        fragment
        for block in layout.pages[0].body
        if isinstance(block, ParagraphBox)
        for line in block.lines
        for fragment in line.fragments
        if isinstance(fragment, ImageBox)
    )
    assert (
        tuple(image.content_type for image in layout_images),
        tuple((image.width, image.height) for image in layout_images),
        pdf_image_count,
    ) == (
        ("image/png", "image/jpeg"),
        ((72.0, 72.0), (36.0, 72.0)),
        2,
    )


def test_converter_emits_a4_landscape_page_geometry() -> None:
    source = _minimal_docx(
        "landscape body",
        final_section_properties="""
          <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        """,
    )
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    page = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0]
    layout = converter.last_layout
    assert layout is not None
    assert (float(page.mediabox.width), float(page.mediabox.height)) == pytest.approx(
        (841.9, 595.3), abs=0.1
    )
    assert (layout.pages[0].width, layout.pages[0].height) == pytest.approx((841.9, 595.3), abs=0.1)


def test_converter_changes_pdf_page_geometry_at_section_boundary() -> None:
    source = _minimal_docx(
        body_xml="""
          <w:p><w:pPr><w:sectPr>
            <w:type w:val="nextPage"/>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          </w:sectPr></w:pPr><w:r><w:t>portrait section</w:t></w:r></w:p>
          <w:p><w:r><w:t>landscape section</w:t></w:r></w:p>
        """,
        final_section_properties="""
          <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        """,
    )
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    pdf = PdfReader(BytesIO(result.pdf_bytes or b""))
    sizes = tuple((float(page.mediabox.width), float(page.mediabox.height)) for page in pdf.pages)
    texts = tuple(page.extract_text().strip() for page in pdf.pages)
    layout = converter.last_layout
    assert layout is not None
    assert sizes[0] == pytest.approx((595.3, 841.9), abs=0.1)
    assert sizes[1] == pytest.approx((841.9, 595.3), abs=0.1)
    assert texts == ("portrait section", "landscape section")
    assert tuple(page.section_index for page in layout.pages) == (0, 1)


def test_converter_renders_header_footer_page_and_total_page_fields() -> None:
    relationships = f"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="{PACKAGE_REL_NS}">
      <Relationship Id="rIdHeader" Type="{R_NS}/header" Target="header1.xml"/>
      <Relationship Id="rIdFooter" Type="{R_NS}/footer" Target="footer1.xml"/>
    </Relationships>""".encode()
    header = f"""<w:hdr xmlns:w="{W_NS}">
      <w:p><w:r><w:t>fixture header</w:t></w:r></w:p>
    </w:hdr>""".encode()
    footer = f"""<w:ftr xmlns:w="{W_NS}">
      <w:p><w:r><w:t>Page </w:t></w:r>
        <w:fldSimple w:instr=" PAGE "><w:r><w:t>0</w:t></w:r></w:fldSimple>
        <w:r><w:t> / </w:t></w:r>
        <w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>0</w:t></w:r></w:fldSimple>
      </w:p>
    </w:ftr>""".encode()
    source = _minimal_docx(
        body_xml="""
          <w:p><w:r><w:t>first body</w:t></w:r></w:p>
          <w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>second body</w:t></w:r></w:p>
        """,
        document_relationships=relationships,
        extra_parts={"word/header1.xml": header, "word/footer1.xml": footer},
        final_section_properties="""
          <w:headerReference w:type="default" r:id="rIdHeader"/>
          <w:footerReference w:type="default" r:id="rIdFooter"/>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
        """,
    )
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    texts = tuple(page.extract_text() for page in PdfReader(BytesIO(result.pdf_bytes or b"")).pages)
    layout = converter.last_layout
    assert layout is not None
    assert result.page_count == 2
    assert all("fixture header" in text for text in texts)
    normalized_texts = tuple(" ".join(text.split()) for text in texts)
    assert "Page 1 / 2" in normalized_texts[0]
    assert "Page 2 / 2" in normalized_texts[1]
    assert all(page.header and page.footer for page in layout.pages)


def test_converter_emits_two_japanese_pages_at_explicit_break(tmp_path: Path) -> None:
    first_text = "日本語の文書"
    second_text = "文書の日本語"
    font_path = tmp_path / "test-japanese-pages.ttf"
    _build_test_japanese_font(font_path, first_text + second_text)
    source = _minimal_docx(
        body_xml=f"""
          <w:p><w:r><w:t>{first_text}</w:t></w:r></w:p>
          <w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>{second_text}</w:t></w:r></w:p>
        """,
        east_asia_font="Test Japanese Pages",
    )
    converter = Converter(
        ConversionOptions(
            strict=True,
            font_configuration=FontConfiguration(
                registered_fonts={"Test Japanese Pages": font_path},
                include_system_fonts=False,
            ),
        )
    )

    result = converter.convert(source)

    pdf_texts = tuple(
        page.extract_text().strip() for page in PdfReader(BytesIO(result.pdf_bytes or b"")).pages
    )
    layout = converter.last_layout
    assert layout is not None
    assert result.page_count == layout.page_count == 2
    assert pdf_texts == (first_text, second_text)


def test_converter_records_configured_east_asia_font_substitution(tmp_path: Path) -> None:
    text = "代替書体"
    font_path = tmp_path / "substitute.ttf"
    _build_test_japanese_font(font_path, text)
    source = _minimal_docx(text, east_asia_font="Missing Japanese Face")
    converter = Converter(
        ConversionOptions(
            strict=True,
            font_configuration=FontConfiguration(
                registered_fonts={"Fixture Substitute": font_path},
                substitutions={"Missing Japanese Face": "Fixture Substitute"},
                include_system_fonts=False,
            ),
        )
    )

    result = converter.convert(source)

    assert len(result.font_substitutions) == 1
    substitution = result.font_substitutions[0]
    assert substitution.requested_font == "Missing Japanese Face"
    assert substitution.selected_font == "DocxPdf Test Japanese"
    assert substitution.reason == "configured"
    assert (substitution.paragraph_index, substitution.run_index) == (0, 0)
    assert substitution.east_asia is True


def test_converter_splits_table_across_pages_and_repeats_header() -> None:
    source = _minimal_docx(body_xml=_paginated_table_xml(data_rows=40))
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    pdf = PdfReader(BytesIO(result.pdf_bytes or b""))
    layout = converter.last_layout
    assert layout is not None
    table_pages = tuple(page.body[0] for page in layout.pages)
    assert result.page_count == len(pdf.pages) == 2
    assert all(isinstance(table, TableBox) for table in table_pages)
    assert table_pages[0].continued_from_previous_page is False
    assert table_pages[0].continues_on_next_page is True
    assert table_pages[1].continued_from_previous_page is True
    assert table_pages[1].continues_on_next_page is False
    assert all("repeated header" in page.extract_text() for page in pdf.pages)
    assert "row 01" in pdf.pages[0].extract_text()
    assert "row 40" in pdf.pages[1].extract_text()


def test_converter_preserves_horizontal_and_vertical_table_merges() -> None:
    source = _minimal_docx(body_xml=_merged_table_xml())
    converter = Converter(ConversionOptions(strict=False))

    result = converter.convert(source)

    layout = converter.last_layout
    assert layout is not None
    table = layout.pages[0].body[0]
    assert isinstance(table, TableBox)
    horizontal = next(cell for cell in table.cells if cell.row_index == 0)
    vertical = next(cell for cell in table.cells if cell.row_index == 1 and cell.column_index == 0)
    assert (horizontal.column_span, horizontal.row_span) == (2, 1)
    assert (vertical.column_span, vertical.row_span) == (1, 2)
    assert not any(cell.row_index == 2 and cell.column_index == 0 for cell in table.cells)
    extracted = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0].extract_text()
    for text in ("horizontal merge", "vertical merge", "right first", "right second"):
        assert text in extracted


def test_converter_handles_legacy_constructs_with_default_options() -> None:
    # Default options are lenient: a content control, tracked changes, a
    # nested-looking complex field, a VML image, and an OLE preview that
    # cannot be decoded must all convert without raising, with content
    # preserved and the undecodable object reserved as a placeholder.
    body_xml = f"""
      <w:sdt><w:sdtPr/><w:sdtContent>
        <w:p><w:r><w:t>content control text</w:t></w:r></w:p>
      </w:sdtContent></w:sdt>
      <w:p>
        <w:ins w:id="1" w:author="A"><w:r><w:t>inserted </w:t></w:r></w:ins>
        <w:del w:id="2" w:author="A"><w:r><w:delText>deleted </w:delText></w:r></w:del>
        <w:r><w:t>kept</w:t></w:r>
      </w:p>
      <w:p>
        <w:r><w:t>Page </w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:p><w:r>
        <w:pict xmlns:v="{VML_NS}"><v:shape style="width:20pt;height:10pt">
          <v:imagedata r:id="rIdVml"/></v:shape></w:pict>
      </w:r></w:p>
      <w:p><w:r>
        <w:object xmlns:v="{VML_NS}">
          <v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdOle"/></v:shape>
        </w:object>
      </w:r></w:p>
    """
    document_relationships = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{PACKAGE_REL_NS}">
  <Relationship Id="rIdVml" Type="{R_NS}/image" Target="media/vml-image.png"/>
  <Relationship Id="rIdOle" Type="{R_NS}/image" Target="media/ole-preview.wmf"/>
</Relationships>""".encode()
    source = _minimal_docx(
        body_xml=body_xml,
        document_relationships=document_relationships,
        extra_parts={
            "word/media/vml-image.png": PNG_1X1,
            "word/media/ole-preview.wmf": b"not a real wmf",
        },
    )

    result = Converter().convert(source)

    assert result.page_count >= 1
    extracted = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0].extract_text()
    assert "content control text" in extracted
    assert "inserted" in extracted
    assert "deleted" not in extracted
    assert "kept" in extracted
    placeholder_warnings = [
        warning for warning in result.warnings if warning.code == "content_placeholder"
    ]
    assert len(placeholder_warnings) == 1
    assert placeholder_warnings[0].feature is not None
    assert placeholder_warnings[0].feature.element == "[Embedded Object]"
