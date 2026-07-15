from __future__ import annotations

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from docxpdf_native.exceptions import (
    InvalidOoxmlError,
    ResourceLimitError,
    UnsupportedFeatureError,
)
from docxpdf_native.models.document import ParagraphModel, TableModel
from docxpdf_native.models.options import ConversionOptions, ResourceLimits
from docxpdf_native.ooxml.package import OoxmlPackage
from docxpdf_native.ooxml.parser import OoxmlDocumentParser

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V_NS = "urn:schemas-microsoft-com:vml"


def make_relationships(*items: tuple[str, str, str]) -> bytes:
    relationships = "".join(
        f'<Relationship Id="{relation_id}" Type="{relation_type}" Target="{target}"/>'
        for relation_id, relation_type, target in items
    )
    return f'<Relationships xmlns="{PR_NS}">{relationships}</Relationships>'.encode()


def build_docx(document: bytes, extra_parts: dict[str, bytes] | None = None) -> bytes:
    parts = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT_NS}">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="png" ContentType="image/png"/>'
            '<Default Extension="jpg" ContentType="image/jpeg"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>"
        ).encode(),
        "_rels/.rels": make_relationships(
            (
                "rId1",
                f"{R_NS}/officeDocument",
                "word/document.xml",
            )
        ),
        "word/document.xml": document,
    }
    parts.update(extra_parts or {})
    target = BytesIO()
    with ZipFile(target, "w", compression=ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return target.getvalue()


def test_parser_reads_paragraph_runs_properties_and_section_geometry() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}">
      <w:body>
        <w:p>
          <w:pPr>
            <w:pStyle w:val="BodyText"/><w:jc w:val="center"/>
            <w:ind w:left="720" w:right="360" w:firstLine="240"/>
            <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="exact"/>
            <w:keepNext/><w:keepLines/><w:widowControl w:val="0"/>
          </w:pPr>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Helvetica" w:eastAsia="Noto Sans CJK JP"/>
              <w:sz w:val="24"/><w:b/><w:i/><w:u w:val="single"/>
              <w:color w:val="123abc"/><w:spacing w:val="20"/>
            </w:rPr>
            <w:t xml:space="preserve"> Hello </w:t>
          </w:r>
          <w:r><w:tab/></w:r><w:r><w:br/></w:r><w:r><w:br w:type="page"/></w:r>
        </w:p>
        <w:sectPr>
          <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
          <w:pgMar w:top="720" w:right="900" w:bottom="1080" w:left="1440"
                   w:header="360" w:footer="480"/>
          <w:pgNumType w:start="7"/>
        </w:sectPr>
      </w:body>
    </w:document>""".encode()

    model = OoxmlDocumentParser().parse(build_docx(document))

    section = model.sections[0]
    paragraph = section.blocks[0]
    assert isinstance(paragraph, ParagraphModel)
    assert (section.page_width, section.page_height, section.orientation) == (
        792.0,
        612.0,
        "landscape",
    )
    assert (
        section.margin_top,
        section.margin_right,
        section.margin_bottom,
        section.margin_left,
    ) == (36.0, 45.0, 54.0, 72.0)
    assert (section.header_distance, section.footer_distance, section.page_number_start) == (
        18.0,
        24.0,
        7,
    )
    assert paragraph.style_id == "BodyText"
    assert paragraph.properties.alignment == "center"
    assert paragraph.properties.left_indent == 36.0
    assert paragraph.properties.first_line_indent == 12.0
    assert paragraph.properties.line_spacing == 18.0
    assert paragraph.properties.keep_next is True
    assert paragraph.properties.widow_control is False
    assert paragraph.runs[0].text == " Hello "
    assert paragraph.runs[0].preserve_space is True
    assert paragraph.runs[0].properties.east_asia_font == "Noto Sans CJK JP"
    assert paragraph.runs[0].properties.font_size == 12.0
    assert paragraph.runs[0].properties.color == "123ABC"
    assert paragraph.runs[1].tab is True
    assert paragraph.runs[2].break_type == "line"
    assert paragraph.runs[3].break_type == "page"


def test_parser_reads_doc_grid_lines_type_and_line_pitch() -> None:
    # 360 twips == 18pt; w:type="lines" is the default Japanese template's
    # setting and is what makes Word snap line heights to the grid.
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:docGrid w:type="lines" w:linePitch="360"/>
      </w:sectPr>
    </w:body></w:document>""".encode()

    model = OoxmlDocumentParser().parse(build_docx(document))

    section = model.sections[0]
    assert section.doc_grid_type == "lines"
    assert section.doc_grid_line_pitch == 18.0


def test_parser_defaults_doc_grid_when_absent() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr/>
    </w:body></w:document>""".encode()

    model = OoxmlDocumentParser().parse(build_docx(document))

    section = model.sections[0]
    assert section.doc_grid_type == "default"
    assert section.doc_grid_line_pitch is None


def test_parser_splits_sections_at_paragraph_section_properties() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:p><w:r><w:t>first</w:t></w:r><w:pPr><w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/><w:type w:val="continuous"/>
      </w:sectPr></w:pPr></w:p>
      <w:p><w:r><w:t>second</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>
    </w:body></w:document>""".encode()

    sections = OoxmlDocumentParser().parse(build_docx(document)).sections

    assert len(sections) == 2
    assert sections[0].blocks[0].text == "first"  # type: ignore[union-attr]
    assert sections[0].section_break == "continuous"
    assert sections[1].blocks[0].text == "second"  # type: ignore[union-attr]
    assert sections[1].orientation == "landscape"


def test_parser_reads_fixed_table_geometry_and_cell_properties() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:tbl>
        <w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>
          <w:jc w:val="center"/><w:tblInd w:w="240" w:type="dxa"/>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr><w:trPr><w:trHeight w:val="400" w:hRule="exact"/><w:cantSplit/><w:tblHeader/></w:trPr>
          <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:gridSpan w:val="2"/>
            <w:vMerge w:val="restart"/><w:shd w:fill="ffeeaa"/><w:vAlign w:val="center"/>
          </w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:sectPr/>
    </w:body></w:document>""".encode()

    table = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(table, TableModel)
    assert table.grid_widths == (150.0, 150.0)
    assert (table.width, table.autofit, table.alignment, table.left_indent) == (
        300.0,
        False,
        "center",
        12.0,
    )
    assert (table.rows[0].height, table.rows[0].height_rule) == (20.0, "exact")
    assert table.rows[0].cant_split is True
    assert table.rows[0].repeat_header is True
    cell = table.rows[0].cells[0]
    assert (cell.width, cell.grid_span, cell.vertical_merge) == (150.0, 2, "restart")
    assert (cell.background_color, cell.vertical_alignment) == ("FFEEAA", "center")
    assert cell.paragraphs[0].text == "A"


def test_parser_reads_inline_png_through_relationship() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body><w:p><w:r>
      <w:drawing><wp:inline><wp:extent cx="12700" cy="25400"/>
        <wp:docPr id="1" name="Picture" descr="sample"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId5"/>
        </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    image = b"\x89PNG\r\n\x1a\ncontent"
    document_rels = make_relationships(("rId5", f"{R_NS}/image", "media/image1.png"))

    parsed = OoxmlDocumentParser().parse(
        build_docx(
            document,
            {
                "word/_rels/document.xml.rels": document_rels,
                "word/media/image1.png": image,
            },
        )
    )
    paragraph = parsed.sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    parsed_image = paragraph.runs[0].image
    assert parsed_image is not None
    assert (parsed_image.part_name, parsed_image.content_type) == (
        "word/media/image1.png",
        "image/png",
    )
    assert (parsed_image.width, parsed_image.height, parsed_image.description) == (
        1.0,
        2.0,
        "sample",
    )
    assert parsed_image.data == image


def test_parser_reads_header_and_footer_references() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>
      <w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr>
      <w:headerReference w:type="default" r:id="rId2"/>
      <w:footerReference w:type="even" r:id="rId3"/>
      </w:sectPr></w:body></w:document>""".encode()
    document_rels = make_relationships(
        ("rId2", f"{R_NS}/header", "header1.xml"),
        ("rId3", f"{R_NS}/footer", "footer1.xml"),
    )
    header = f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>head</w:t></w:r></w:p></w:hdr>'.encode()
    footer = f'<w:ftr xmlns:w="{W_NS}"><w:p><w:r><w:t>foot</w:t></w:r></w:p></w:ftr>'.encode()

    section = (
        OoxmlDocumentParser()
        .parse(
            build_docx(
                document,
                {
                    "word/_rels/document.xml.rels": document_rels,
                    "word/header1.xml": header,
                    "word/footer1.xml": footer,
                },
            )
        )
        .sections[0]
    )

    assert section.headers[0].kind == "default"
    assert section.headers[0].blocks[0].text == "head"  # type: ignore[union-attr]
    assert section.footers[0].kind == "even"
    assert section.footers[0].blocks[0].text == "foot"  # type: ignore[union-attr]


def test_parser_rejects_malformed_document_xml() -> None:
    with pytest.raises(InvalidOoxmlError, match=r"word/document\.xml"):
        OoxmlDocumentParser().parse(build_docx(b"<broken>"))


def test_parser_enforces_paragraph_limit() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:p><w:r><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>
      <w:p><w:r><w:t>three</w:t></w:r></w:p><w:sectPr/>
    </w:body></w:document>""".encode()
    parser = OoxmlDocumentParser(limits=ResourceLimits(max_paragraphs=1))

    with pytest.raises(ResourceLimitError, match="paragraph count"):
        parser.parse(build_docx(document))


def test_parser_uses_lenient_unsupported_feature_policy_by_default() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:t>kept</w:t></w:r><w:txbxContent><w:p/></w:txbxContent>
      </w:p><w:sectPr/></w:body></w:document>""".encode()
    parser = OoxmlDocumentParser()

    model = parser.parse(build_docx(document))

    assert model.sections[0].blocks[0].text == "kept"  # type: ignore[union-attr]
    assert parser.unsupported_features[0].name == "text_box"
    assert parser.warnings[0].code == "unsupported_feature"


def test_parser_raises_in_strict_mode_for_genuinely_unsupported_content() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:t>kept</w:t></w:r><w:txbxContent><w:p/></w:txbxContent>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    with pytest.raises(UnsupportedFeatureError) as caught:
        OoxmlDocumentParser().parse(build_docx(document), options=ConversionOptions(strict=True))

    assert caught.value.feature.name == "text_box"


def test_parser_records_lenient_warning_and_keeps_supported_content() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:t>kept</w:t></w:r><w:txbxContent><w:p/></w:txbxContent>
      </w:p><w:sectPr/></w:body></w:document>""".encode()
    parser = OoxmlDocumentParser()

    model = parser.parse(build_docx(document), options=ConversionOptions(strict=False))

    assert model.sections[0].blocks[0].text == "kept"  # type: ignore[union-attr]
    assert parser.unsupported_features[0].name == "text_box"
    assert parser.warnings[0].code == "unsupported_feature"


def test_parser_ignores_vml_shape_defaults_in_settings() -> None:
    document = f'<w:document xmlns:w="{W_NS}"><w:body><w:sectPr/></w:body></w:document>'.encode()
    settings = f"""<w:settings xmlns:w="{W_NS}" xmlns:v="urn:schemas-microsoft-com:vml">
      <w:shapeDefaults><v:shape id="default-shape"/></w:shapeDefaults>
      <w:hdrShapeDefaults><v:shape id="default-header-shape"/></w:hdrShapeDefaults>
    </w:settings>""".encode()

    model = OoxmlDocumentParser().parse(build_docx(document, {"word/settings.xml": settings}))

    assert len(model.sections) == 1


def test_parser_reads_styles_defaults_theme_numbering_and_metadata() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Body"/><w:numPr><w:ilvl w:val="0"/>
      <w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>
      <w:sectPr/>
    </w:body></w:document>""".encode()
    styles = f"""<w:styles xmlns:w="{W_NS}">
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif"
        w:eastAsia="Noto Serif CJK JP"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
        <w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/></w:style>
      <w:style w:type="paragraph" w:styleId="Body" w:default="1">
        <w:name w:val="Body text"/><w:basedOn w:val="Base"/><w:next w:val="Body"/>
        <w:link w:val="BodyChar"/><w:pPr><w:jc w:val="both"/></w:pPr>
        <w:rPr><w:b/><w:color w:val="445566"/></w:rPr>
      </w:style>
    </w:styles>""".encode()
    theme = f"""<a:theme xmlns:a="{A_NS}"><a:themeElements><a:fontScheme name="Theme">
      <a:majorFont><a:latin typeface="Major Latin"/><a:ea typeface="Major EA"/>
        <a:cs typeface="Major CS"/></a:majorFont>
      <a:minorFont><a:latin typeface="Minor Latin"/><a:ea typeface="Minor EA"/>
        <a:cs typeface="Minor CS"/></a:minorFont>
    </a:fontScheme></a:themeElements></a:theme>""".encode()
    numbering = f"""<w:numbering xmlns:w="{W_NS}">
      <w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="3"/>
        <w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
        <w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:b/></w:rPr>
      </w:lvl></w:abstractNum>
      <w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>
    </w:numbering>""".encode()
    core = b"""<cp:coreProperties
      xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns:dc="http://purl.org/dc/elements/1.1/">
      <dc:title>Sample</dc:title><dc:creator>Author</dc:creator>
      <cp:keywords>one,two</cp:keywords>
    </cp:coreProperties>"""
    app = b"""<Properties
      xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
      <Application>Fixture Generator</Application><Pages>999</Pages>
    </Properties>"""

    model = OoxmlDocumentParser().parse(
        build_docx(
            document,
            {
                "word/styles.xml": styles,
                "word/theme/theme1.xml": theme,
                "word/numbering.xml": numbering,
                "docProps/core.xml": core,
                "docProps/app.xml": app,
            },
        )
    )

    assert model.defaults.run.font_family == "Liberation Serif"
    assert model.defaults.run.east_asia_font == "Noto Serif CJK JP"
    assert model.defaults.run.font_size == 11.0
    assert model.defaults.paragraph.space_after == 6.0
    body_style = next(style for style in model.styles if style.style_id == "Body")
    assert (body_style.based_on, body_style.next_style, body_style.link) == (
        "Base",
        "Body",
        "BodyChar",
    )
    assert body_style.is_default is True
    assert body_style.paragraph.alignment == "justify"
    assert (body_style.run.bold, body_style.run.color) == (True, "445566")
    assert (model.theme_fonts.major_latin, model.theme_fonts.major_east_asia) == (
        "Major Latin",
        "Major EA",
    )
    assert (model.theme_fonts.minor_latin, model.theme_fonts.minor_complex_script) == (
        "Minor Latin",
        "Minor CS",
    )
    definition = model.numbering[0]
    assert (definition.numbering_id, definition.abstract_numbering_id) == (7, 4)
    assert (definition.levels[0].level, definition.levels[0].start) == (0, 3)
    assert definition.levels[0].paragraph.left_indent == 36.0
    assert definition.levels[0].run.bold is True
    assert (model.metadata.title, model.metadata.creator, model.metadata.keywords) == (
        "Sample",
        "Author",
        "one,two",
    )
    assert model.metadata.application == "Fixture Generator"


def test_parser_applies_unsupported_policy_to_binary_package_parts() -> None:
    document = f'<w:document xmlns:w="{W_NS}"><w:body><w:sectPr/></w:body></w:document>'.encode()
    data = build_docx(document, {"word/vbaProject.bin": b"macro"})

    with pytest.raises(UnsupportedFeatureError) as caught:
        OoxmlDocumentParser().parse(data, options=ConversionOptions(strict=True))

    assert caught.value.feature.name == "macro"


def test_parser_replaces_simple_page_fields_with_layout_sentinels() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:fldSimple w:instr=" PAGE "><w:r><w:t>4</w:t></w:r></w:fldSimple>
      <w:r><w:t>/</w:t></w:r>
      <w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>12</w:t></w:r></w:fldSimple>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert [run.text for run in paragraph.runs] == ["{{PAGE}}", "/", "{{NUMPAGES}}"]


def test_parser_reads_advanced_paragraph_tabs_numbering_and_run_formatting() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr>
      <w:jc w:val="right"/><w:ind w:right="480" w:hanging="240"/>
      <w:spacing w:before="100" w:after="200" w:line="480" w:lineRule="auto"/>
      <w:keepNext w:val="false"/><w:keepLines/><w:pageBreakBefore/>
      <w:widowControl/><w:tabs><w:tab w:val="decimal" w:pos="1440" w:leader="dot"/>
      <w:tab w:val="bar" w:pos="2880" w:leader="middleDot"/></w:tabs>
      <w:numPr><w:ilvl w:val="2"/><w:numId w:val="9"/></w:numPr>
      </w:pPr><w:r><w:rPr><w:rStyle w:val="Emphasis"/>
        <w:rFonts w:hAnsi="Latin Face" w:cs="Complex Face"/><w:sz w:val="18"/>
        <w:b w:val="0"/><w:i w:val="off"/><w:strike/><w:u w:val="none"/>
        <w:color w:val="auto"/><w:highlight w:val="yellow"/>
        <w:vertAlign w:val="superscript"/><w:spacing w:val="10"/><w:vanish/>
      </w:rPr><w:t>Café</w:t></w:r>
      <w:hyperlink><w:r><w:t> link</w:t></w:r></w:hyperlink><w:r/><w:r><w:cr/></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    properties = paragraph.properties
    assert (properties.alignment, properties.right_indent, properties.hanging_indent) == (
        "right",
        24.0,
        12.0,
    )
    assert (properties.space_before, properties.space_after, properties.line_spacing) == (
        5.0,
        10.0,
        2.0,
    )
    assert properties.line_spacing_rule == "auto"
    assert (properties.keep_next, properties.keep_lines, properties.page_break_before) == (
        False,
        True,
        True,
    )
    assert [(tab.position, tab.alignment, tab.leader) for tab in properties.tabs] == [
        (72.0, "decimal", "dot"),
        (144.0, "bar", "middle_dot"),
    ]
    assert (properties.numbering_id, properties.numbering_level) == (9, 2)
    run = paragraph.runs[0]
    assert run.text == "Café"
    assert (run.style_id, run.hidden) == ("Emphasis", True)
    assert (run.properties.font_family, run.properties.high_ansi_font) == (
        "Latin Face",
        "Latin Face",
    )
    assert run.properties.complex_script_font == "Complex Face"
    assert (run.properties.bold, run.properties.italic, run.properties.strike) == (
        False,
        False,
        True,
    )
    assert run.properties.underline is False
    assert (run.properties.highlight, run.properties.vertical_align) == (
        "FFFF00",
        "superscript",
    )
    assert run.properties.character_spacing == 0.5
    assert paragraph.runs[1].text == " link"
    assert paragraph.runs[2].text == ""
    assert paragraph.runs[3].break_type == "line"


def test_parser_accepts_numbering_tab_alignment() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr>
      <w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs>
      </w:pPr><w:r><w:t>numbered item</w:t></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.properties.tabs[0].alignment == "num"


def test_parser_reads_table_margins_borders_and_continued_vertical_merge() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:tbl><w:tblPr>
      <w:tblCellMar><w:top w:w="100"/><w:right w:w="120"/>
        <w:bottom w:w="140"/><w:left w:w="160"/></w:tblCellMar>
      <w:tblBorders><w:top w:val="double" w:sz="8" w:color="112233"/>
        <w:right w:val="none"/><w:insideH w:val="single" w:sz="4" w:color="auto"/>
      </w:tblBorders></w:tblPr><w:tr><w:trPr>
        <w:trHeight w:val="360" w:hRule="atLeast"/>
      </w:trPr><w:tc><w:tcPr><w:vMerge/>
        <w:tcMar><w:left w:w="200"/></w:tcMar>
        <w:tcBorders><w:bottom w:val="dashed" w:sz="12" w:color="AABBCC"/></w:tcBorders>
      </w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
      </w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>""".encode()

    table = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(table, TableModel)
    assert table.autofit is True
    assert table.cell_margins.model_dump() == {
        "top": 5.0,
        "right": 6.0,
        "bottom": 7.0,
        "left": 8.0,
    }
    assert table.borders.top is not None
    assert (table.borders.top.style, table.borders.top.width, table.borders.top.color) == (
        "double",
        1.0,
        "112233",
    )
    assert table.borders.right is None
    assert table.borders.inside_horizontal is not None
    row = table.rows[0]
    assert (row.height, row.height_rule) == (18.0, "at_least")
    cell = row.cells[0]
    assert cell.vertical_merge == "continue"
    assert cell.margins.left == 10.0
    assert cell.borders.bottom is not None
    assert cell.borders.bottom.width == 1.5
    assert cell.text_alignment == "right"


def test_parser_reads_inline_jpeg_content_type() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body><w:p><w:r>
      <w:drawing><wp:inline><wp:extent cx="25400" cy="12700"/>
        <wp:docPr id="2" name="Photo" title="title only"/>
        <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId8"/>
        </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(("rId8", f"{R_NS}/image", "media/photo.jpg"))

    paragraph = (
        OoxmlDocumentParser()
        .parse(
            build_docx(
                document,
                {
                    "word/_rels/document.xml.rels": relation,
                    "word/media/photo.jpg": b"jpeg",
                },
            )
        )
        .sections[0]
        .blocks[0]
    )

    assert isinstance(paragraph, ParagraphModel)
    image = paragraph.runs[0].image
    assert image is not None
    assert (image.content_type, image.width, image.height, image.description) == (
        "image/jpeg",
        2.0,
        1.0,
        "title only",
    )


def test_parser_enforces_run_limit() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    with pytest.raises(ResourceLimitError, match="run count"):
        OoxmlDocumentParser(limits=ResourceLimits(max_runs=1)).parse(build_docx(document))


def test_parser_enforces_table_cell_limit() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:tbl><w:tr>
      <w:tc><w:p/></w:tc><w:tc><w:p/></w:tc>
      </w:tr></w:tbl><w:sectPr/></w:body></w:document>""".encode()

    with pytest.raises(ResourceLimitError, match="table cell count"):
        OoxmlDocumentParser(limits=ResourceLimits(max_table_cells=1)).parse(build_docx(document))


@pytest.mark.parametrize(
    "document",
    [
        b"<root />",
        f'<w:document xmlns:w="{W_NS}"/>'.encode(),
    ],
)
def test_parser_rejects_wrong_document_shape(document: bytes) -> None:
    with pytest.raises(InvalidOoxmlError):
        OoxmlDocumentParser().parse(build_docx(document))


@pytest.mark.parametrize(
    ("part_name", "content"),
    [
        ("word/styles.xml", b"<root />"),
        ("word/theme/theme1.xml", b"<root />"),
        ("word/numbering.xml", b"<root />"),
        ("word/settings.xml", b"<root />"),
        ("word/fontTable.xml", b"<root />"),
        ("docProps/core.xml", b"<root />"),
        ("docProps/app.xml", b"<root />"),
        (
            "word/styles.xml",
            f'<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph"/></w:styles>'.encode(),
        ),
    ],
)
def test_parser_rejects_malformed_auxiliary_ooxml_parts(
    part_name: str,
    content: bytes,
) -> None:
    document = f'<w:document xmlns:w="{W_NS}"><w:body><w:sectPr/></w:body></w:document>'.encode()

    with pytest.raises(InvalidOoxmlError):
        OoxmlDocumentParser().parse(build_docx(document, {part_name: content}))


@pytest.mark.parametrize(
    "section_properties",
    [
        '<w:pgSz w:orient="sideways"/>',
        '<w:type w:val="unknown"/>',
        '<w:docGrid w:type="unknown"/>',
    ],
)
def test_parser_rejects_invalid_section_properties(section_properties: str) -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>{section_properties}</w:sectPr></w:body></w:document>""".encode()

    with pytest.raises(InvalidOoxmlError):
        OoxmlDocumentParser().parse(build_docx(document))


@pytest.mark.parametrize("allow_external_relationships", [False, True])
def test_parser_renders_hyperlink_text_regardless_of_external_relationship_policy(
    allow_external_relationships: bool,
) -> None:
    # External relationships (hyperlinks, in particular) are extremely common
    # in real-world DOCX files and are never fetched, so their mere presence
    # must never fail the conversion -- ``allow_external_relationships`` is
    # retained on ConversionOptions for backward compatibility only and no
    # longer changes this behavior.
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body><w:p>
      <w:hyperlink r:id="rId9"><w:r><w:t>link text</w:t></w:r></w:hyperlink>
      </w:p><w:sectPr/></w:body></w:document>""".encode()
    relationships = (
        f'<Relationships xmlns="{PR_NS}"><Relationship Id="rId9" Type="{R_NS}/hyperlink" '
        'Target="https://example.invalid/" TargetMode="External"/></Relationships>'
    ).encode()
    data = build_docx(document, {"word/_rels/document.xml.rels": relationships})

    paragraph = (
        OoxmlDocumentParser()
        .parse(
            data,
            options=ConversionOptions(allow_external_relationships=allow_external_relationships),
        )
        .sections[0]
        .blocks[0]
    )

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.text == "link text"


def test_parser_tolerates_external_relationship_owned_by_an_auxiliary_part() -> None:
    # A real-world failure mode: an external relationship (e.g. a hyperlink
    # inside footnotes.xml) owned by a part other than word/document.xml must
    # not fail the whole conversion either.
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p>
      <w:sectPr/></w:body></w:document>""".encode()
    footnotes_relationships = (
        f'<Relationships xmlns="{PR_NS}"><Relationship Id="rId1" Type="{R_NS}/hyperlink" '
        'Target="https://example.invalid/" TargetMode="External"/></Relationships>'
    ).encode()
    data = build_docx(
        document,
        {
            "word/footnotes.xml": f'<w:footnotes xmlns:w="{W_NS}"/>'.encode(),
            "word/_rels/footnotes.xml.rels": footnotes_relationships,
        },
    )

    model = OoxmlDocumentParser().parse(data)

    assert model.sections[0].blocks[0].text == "body"  # type: ignore[union-attr]


def test_parser_never_fetches_external_relationship_targets() -> None:
    # Presence is tolerated, but the target is still never resolved to a
    # package part -- there is nothing to "fetch" because it never had bytes
    # inside the DOCX to begin with.
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body><w:p>
      <w:hyperlink r:id="rId9"><w:r><w:t>link text</w:t></w:r></w:hyperlink>
      </w:p><w:sectPr/></w:body></w:document>""".encode()
    relationships = (
        f'<Relationships xmlns="{PR_NS}"><Relationship Id="rId9" Type="{R_NS}/hyperlink" '
        'Target="https://example.invalid/" TargetMode="External"/></Relationships>'
    ).encode()
    data = build_docx(document, {"word/_rels/document.xml.rels": relationships})

    OoxmlDocumentParser().parse(data)

    package = OoxmlPackage.open(data, allow_external_relationships=True)
    relation = package.relationships_for("word/document.xml").by_id("rId9")
    assert relation.is_external is True
    assert relation.resolved_target is None


@pytest.mark.parametrize("strict", [False, True])
def test_parser_replaces_external_image_with_same_size_placeholder(strict: bool) -> None:
    # External bytes are never fetched, but the wp:extent size is still known
    # from the markup, so the image's footprint is preserved as a
    # placeholder instead of being silently dropped -- in both modes, since
    # this degradation is always reported rather than ever being silent.
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body><w:p><w:r>
      <w:drawing><wp:inline><wp:extent cx="127000" cy="63500"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/>
      </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relationships = (
        f'<Relationships xmlns="{PR_NS}"><Relationship Id="rId7" Type="{R_NS}/image" '
        'Target="https://example.invalid/image.png" TargetMode="External"/></Relationships>'
    ).encode()
    data = build_docx(document, {"word/_rels/document.xml.rels": relationships})

    parser = OoxmlDocumentParser()
    model = parser.parse(data, options=ConversionOptions(strict=strict))
    paragraph = model.sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert placeholder.label == "[Image]"
    assert (placeholder.width, placeholder.height) == (10.0, 5.0)
    assert parser.warnings[0].code == "content_placeholder"
    assert parser.unsupported_features[0].name == "content_placeholder"


def test_parser_rejects_header_with_wrong_root_element() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>
      <w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr>
      </w:body></w:document>""".encode()
    relationship = make_relationships(("rId2", f"{R_NS}/header", "header1.xml"))

    with pytest.raises(InvalidOoxmlError, match=r"header1\.xml"):
        OoxmlDocumentParser().parse(
            build_docx(
                document,
                {
                    "word/_rels/document.xml.rels": relationship,
                    "word/header1.xml": b"<root />",
                },
            )
        )


def test_parser_preserves_theme_font_references_on_direct_run_formatting() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"
        w:eastAsiaTheme="minorEastAsia" w:csTheme="majorBidi"/>
      </w:rPr><w:t>theme</w:t></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    properties = paragraph.runs[0].properties
    assert (properties.font_family, properties.ascii_font, properties.high_ansi_font) == (
        "+majorHAnsi",
        "+majorHAnsi",
        "+majorHAnsi",
    )
    assert properties.east_asia_font == "+minorEastAsia"
    assert properties.complex_script_font == "+majorBidi"


@pytest.mark.parametrize("huge_length", ["9" * 1000, "20000020"])
def test_parser_rejects_ooxml_length_outside_supported_point_range(huge_length: str) -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p/>
      <w:sectPr><w:pgSz w:w="{huge_length}" w:h="16838"/></w:sectPr>
      </w:body></w:document>""".encode()

    with pytest.raises(InvalidOoxmlError, match="point range"):
        OoxmlDocumentParser().parse(build_docx(document))


def test_parser_wraps_invalid_model_value_as_ooxml_error() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p/>
      <w:sectPr><w:pgSz w:w="-1" w:h="16838"/></w:sectPr>
      </w:body></w:document>""".encode()

    with pytest.raises(InvalidOoxmlError, match="model validation"):
        OoxmlDocumentParser().parse(build_docx(document))


def test_parser_replaces_undecodable_image_format_with_placeholder() -> None:
    # SVG is neither a natively supported format nor Pillow-decodable, so it
    # degrades to a same-size placeholder instead of raising or vanishing.
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body><w:p><w:r>
      <w:drawing><wp:inline><wp:extent cx="12700" cy="25400"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId6"/>
      </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(("rId6", f"{R_NS}/image", "media/image.svg"))
    data = build_docx(
        document,
        {
            "word/_rels/document.xml.rels": relation,
            "word/media/image.svg": b"<svg />",
        },
    )

    parser = OoxmlDocumentParser()
    model = parser.parse(data, options=ConversionOptions(strict=True))
    paragraph = model.sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert (placeholder.label, placeholder.width, placeholder.height) == ("[Image]", 1.0, 2.0)
    assert parser.unsupported_features[0].name == "content_placeholder"


def test_parser_converts_gif_image_to_png_when_pillow_is_available() -> None:
    pytest.importorskip("PIL")
    from io import BytesIO

    from PIL import Image

    buffer = BytesIO()
    Image.new("RGB", (2, 3), color=(10, 20, 30)).save(buffer, format="GIF")
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body><w:p><w:r>
      <w:drawing><wp:inline><wp:extent cx="12700" cy="25400"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId6"/>
      </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(("rId6", f"{R_NS}/image", "media/image.gif"))
    data = build_docx(
        document,
        {
            "word/_rels/document.xml.rels": relation,
            "word/media/image.gif": buffer.getvalue(),
        },
    )

    parser = OoxmlDocumentParser()
    model = parser.parse(data, options=ConversionOptions(strict=True))
    paragraph = model.sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    image = paragraph.runs[0].image
    assert image is not None
    assert image.content_type == "image/png"
    assert image.data[:8] == b"\x89PNG\r\n\x1a\n"
    assert parser.unsupported_features == ()


def test_parser_flattens_block_level_content_control() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body>
      <w:sdt><w:sdtPr/><w:sdtContent>
        <w:p><w:r><w:t>controlled</w:t></w:r></w:p>
      </w:sdtContent></w:sdt>
      <w:sectPr/>
    </w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.text == "controlled"


def test_parser_flattens_nested_inline_content_controls() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:t>before </w:t></w:r>
      <w:sdt><w:sdtPr/><w:sdtContent>
        <w:sdt><w:sdtPr/><w:sdtContent>
          <w:r><w:t>nested</w:t></w:r>
        </w:sdtContent></w:sdt>
      </w:sdtContent></w:sdt>
      <w:r><w:t> after</w:t></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.text == "before nested after"


def test_parser_renders_inserted_text_and_drops_deleted_text() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:ins w:id="1" w:author="A"><w:r><w:t>inserted </w:t></w:r></w:ins>
      <w:del w:id="2" w:author="A"><w:r><w:delText>deleted </w:delText></w:r></w:del>
      <w:r><w:t>kept</w:t></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.text == "inserted kept"


def test_parser_replaces_complex_page_field_with_sentinel() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>3</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert [run.text for run in paragraph.runs] == ["{{PAGE}}"]


def test_parser_renders_only_outermost_complex_field_result() -> None:
    # { IF { PAGE } = 1 "First" "Other" } -- the inner PAGE field lives
    # entirely inside the outer IF field's instruction, so its own cached
    # result ("4") must never surface; only the outer's cached result does.
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>4</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:instrText xml:space="preserve"> = 1 "First" "Other" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>First</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert [run.text for run in paragraph.runs] == ["First"]


def test_parser_renders_other_complex_field_commands_as_cached_text() -> None:
    document = f"""<w:document xmlns:w="{W_NS}"><w:body><w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Jane Doe</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.text == "Jane Doe"


def test_parser_renders_vml_pict_image_from_style_dimensions() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:v="{V_NS}">
      <w:body><w:p><w:r>
      <w:pict><v:shape style="width:36pt;height:18pt"><v:imagedata r:id="rId5"/></v:shape></w:pict>
      </w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    image = b"\x89PNG\r\n\x1a\ncontent"
    relation = make_relationships(("rId5", f"{R_NS}/image", "media/image1.png"))

    paragraph = (
        OoxmlDocumentParser()
        .parse(
            build_docx(
                document,
                {
                    "word/_rels/document.xml.rels": relation,
                    "word/media/image1.png": image,
                },
            )
        )
        .sections[0]
        .blocks[0]
    )

    assert isinstance(paragraph, ParagraphModel)
    parsed_image = paragraph.runs[0].image
    assert parsed_image is not None
    assert (parsed_image.width, parsed_image.height) == (36.0, 18.0)
    assert parsed_image.data == image


@pytest.mark.parametrize(
    ("style", "expected"),
    [
        ("width:1in;height:0.5in", (72.0, 36.0)),
        ("width:2.54cm;height:1.27cm", (72.0, 36.0)),
        ("width:96px;height:48px", (72.0, 36.0)),
    ],
)
def test_parser_converts_vml_style_length_units_to_points(
    style: str,
    expected: tuple[float, float],
) -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:v="{V_NS}">
      <w:body><w:p><w:r>
      <w:pict><v:shape style="{style}"><v:imagedata r:id="rId5"/></v:shape></w:pict>
      </w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(("rId5", f"{R_NS}/image", "media/image1.png"))

    paragraph = (
        OoxmlDocumentParser()
        .parse(
            build_docx(
                document,
                {
                    "word/_rels/document.xml.rels": relation,
                    "word/media/image1.png": b"\x89PNG\r\n\x1a\ncontent",
                },
            )
        )
        .sections[0]
        .blocks[0]
    )

    assert isinstance(paragraph, ParagraphModel)
    parsed_image = paragraph.runs[0].image
    assert parsed_image is not None
    assert (parsed_image.width, parsed_image.height) == expected


def test_parser_replaces_ole_embedded_object_with_dimension_preserving_placeholder() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:v="{V_NS}">
      <w:body><w:p><w:r>
      <w:object>
        <v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rId7"/></v:shape>
      </w:object>
      </w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(("rId7", f"{R_NS}/image", "media/image1.emf"))

    parser = OoxmlDocumentParser()
    model = parser.parse(
        build_docx(
            document,
            {
                "word/_rels/document.xml.rels": relation,
                "word/media/image1.emf": b"not a real emf",
            },
        )
    )
    paragraph = model.sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert placeholder.label == "[Embedded Object]"
    assert (placeholder.width, placeholder.height) == (72.0, 36.0)
    assert parser.unsupported_features[0].name == "content_placeholder"
    assert parser.unsupported_features[0].status == "placeholder"


def test_parser_falls_back_to_legacy_object_size_without_vml_style() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:v="{V_NS}">
      <w:body><w:p><w:r>
      <w:object w:dxaOrig="1440" w:dyaOrig="720">
        <v:shape><v:imagedata r:id="rId7"/></v:shape>
      </w:object>
      </w:r></w:p><w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(("rId7", f"{R_NS}/image", "media/image1.emf"))

    model = OoxmlDocumentParser().parse(
        build_docx(
            document,
            {
                "word/_rels/document.xml.rels": relation,
                "word/media/image1.emf": b"not a real emf",
            },
        )
    )
    paragraph = model.sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert (placeholder.width, placeholder.height) == (72.0, 36.0)


def test_parser_does_not_stop_on_a_document_combining_legacy_constructs() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:v="{V_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body>
      <w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>
      <w:p>
        <w:ins w:id="1" w:author="A"><w:r><w:t>inserted </w:t></w:r></w:ins>
        <w:del w:id="2" w:author="A"><w:r><w:delText>deleted </w:delText></w:r></w:del>
        <w:r><w:t>kept</w:t></w:r>
      </w:p>
      <w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>1</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:p><w:r><w:pict><v:shape style="width:20pt;height:10pt">
        <v:imagedata r:id="rIdVml"/></v:shape></w:pict></w:r></w:p>
      <w:p><w:r><w:object><v:shape style="width:72pt;height:36pt">
        <v:imagedata r:id="rIdOle"/></v:shape></w:object></w:r></w:p>
      <w:sectPr/></w:body></w:document>""".encode()
    relation = make_relationships(
        ("rIdVml", f"{R_NS}/image", "media/vml-image.png"),
        ("rIdOle", f"{R_NS}/image", "media/ole-preview.wmf"),
    )
    data = build_docx(
        document,
        {
            "word/_rels/document.xml.rels": relation,
            "word/media/vml-image.png": b"\x89PNG\r\n\x1a\ncontent",
            "word/media/ole-preview.wmf": b"not a real wmf",
        },
    )

    # Nothing here raises with the default (lenient) options: every kind of
    # construct is either rendered natively or reserved as a placeholder.
    parser = OoxmlDocumentParser()
    model = parser.parse(data)
    blocks = model.sections[0].blocks

    assert [block.text for block in blocks[:3]] == ["controlled", "inserted kept", "{{PAGE}}"]  # type: ignore[union-attr]
    assert blocks[3].runs[0].image is not None  # type: ignore[union-attr]
    assert blocks[4].runs[0].placeholder is not None  # type: ignore[union-attr]
    assert len(parser.unsupported_features) == 1
    assert parser.unsupported_features[0].status == "placeholder"


_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_DIAGRAM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@pytest.mark.parametrize(
    ("inner", "label"),
    [
        (f'<c:chart xmlns:c="{_CHART_NS}"/>', "[Chart]"),
        (f'<dgm:relIds xmlns:dgm="{_DIAGRAM_NS}"/>', "[SmartArt]"),
        ("<w:txbxContent/>", "[Text Box]"),
        ("<w:noRecognizedMarker/>", "[Shape]"),
    ],
)
def test_parser_labels_anchored_drawing_placeholders_by_content(inner: str, label: str) -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}">
      <w:body><w:p><w:r>
      <w:drawing><wp:anchor><wp:extent cx="127000" cy="63500"/>{inner}</wp:anchor></w:drawing>
      </w:r></w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert placeholder.label == label


def test_parser_labels_vml_textpath_shape_as_wordart_placeholder() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:v="{V_NS}">
      <w:body><w:p><w:r>
      <w:pict><v:shape style="width:100pt;height:20pt">
        <v:textpath string="DRAFT"/></v:shape></w:pict>
      </w:r></w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert placeholder.label == "[WordArt]"


def test_parser_skips_vml_shape_when_no_usable_dimensions_are_found() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:v="{V_NS}">
      <w:body><w:p>
      <w:r><w:t>before</w:t></w:r>
      <w:r><w:pict><v:shape style="width:bogus;height:bogus"/></w:pict></w:r>
      <w:r><w:t>after</w:t></w:r>
      </w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    assert paragraph.text == "beforeafter"


def test_parser_falls_back_to_placeholder_when_image_relationship_is_missing() -> None:
    document = f"""<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"
      xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}"><w:body><w:p><w:r>
      <w:drawing><wp:inline><wp:extent cx="12700" cy="25400"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdMissing"/>
      </pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>""".encode()

    paragraph = OoxmlDocumentParser().parse(build_docx(document)).sections[0].blocks[0]

    assert isinstance(paragraph, ParagraphModel)
    placeholder = paragraph.runs[0].placeholder
    assert placeholder is not None
    assert placeholder.label == "[Image]"
