from __future__ import annotations

from xml.etree import ElementTree

import pytest

from docxpdf_native.models.styles import (
    ParagraphProperties,
    RunProperties,
    StyleModel,
    ThemeFonts,
)
from docxpdf_native.ooxml.styles import (
    OoxmlStyleResolver,
    OoxmlStylesParser,
    ThemeFontParser,
    ThemeFontResolver,
)


def test_style_resolver_applies_document_defaults_and_parent_chain() -> None:
    resolver = OoxmlStyleResolver(
        styles=(
            StyleModel(
                style_id="Base",
                name="Base",
                style_type="paragraph",
                run=RunProperties(font_size=10.0, bold=True),
                paragraph=ParagraphProperties(space_after=8.0),
            ),
            StyleModel(
                style_id="Child",
                name="Child",
                style_type="paragraph",
                based_on="Base",
                run=RunProperties(color="123456"),
                paragraph=ParagraphProperties(left_indent=12.0),
            ),
        ),
        document_run_defaults=RunProperties(font_family="Default Sans", italic=True),
        document_paragraph_defaults=ParagraphProperties(alignment="left", space_before=3.0),
    )

    resolved = resolver.resolve(paragraph_style_id="Child")

    assert resolved.run.font_family == "Default Sans"
    assert resolved.run.font_size == 10.0
    assert resolved.run.bold is True
    assert resolved.run.italic is True
    assert resolved.run.color == "123456"
    assert resolved.paragraph.alignment == "left"
    assert resolved.paragraph.space_before == 3.0
    assert resolved.paragraph.space_after == 8.0
    assert resolved.paragraph.left_indent == 12.0


def test_character_style_and_direct_formatting_have_highest_priority() -> None:
    resolver = OoxmlStyleResolver(
        styles=(
            StyleModel(
                style_id="Paragraph",
                name="Paragraph",
                style_type="paragraph",
                run=RunProperties(bold=True, italic=False, color="000000"),
            ),
            StyleModel(
                style_id="Emphasis",
                name="Emphasis",
                style_type="character",
                run=RunProperties(italic=True, color="FF0000"),
            ),
        )
    )

    resolved = resolver.resolve(
        paragraph_style_id="Paragraph",
        character_style_id="Emphasis",
        run_direct=RunProperties(bold=False, color="00FF00"),
    )

    assert resolved.run.bold is False
    assert resolved.run.italic is True
    assert resolved.run.color == "00FF00"


def test_style_resolver_detects_inheritance_cycle() -> None:
    resolver = OoxmlStyleResolver(
        styles=(
            StyleModel(style_id="A", name="A", style_type="paragraph", based_on="B"),
            StyleModel(style_id="B", name="B", style_type="paragraph", based_on="A"),
        )
    )

    with pytest.raises(ValueError, match=r"style inheritance cycle.*A.*B.*A"):
        resolver.resolve(paragraph_style_id="A")


def test_style_resolver_rejects_unknown_style() -> None:
    resolver = OoxmlStyleResolver(styles=())

    with pytest.raises(ValueError, match="Unknown"):
        resolver.resolve(paragraph_style_id="Unknown")


def test_theme_font_resolver_maps_latin_and_east_asia_placeholders() -> None:
    theme = ThemeFonts(
        major_latin="Major Latin",
        major_east_asia="Major Japanese",
        minor_latin="Minor Latin",
        minor_east_asia="Minor Japanese",
    )

    assert ThemeFontResolver.resolve("+majorHAnsi", theme, east_asia=False) == "Major Latin"
    assert ThemeFontResolver.resolve("+minorEastAsia", theme, east_asia=True) == "Minor Japanese"


def test_theme_font_resolver_falls_back_to_latin_when_complex_script_is_unset() -> None:
    # ``<a:cs typeface=""/>`` (or a missing <a:cs> element) parses to an
    # unset major_complex_script/minor_complex_script -- Word itself then
    # falls back to the scheme's Latin face rather than leaving the run
    # without a resolvable font.
    theme = ThemeFonts(major_latin="Cambria", minor_latin="Calibri")

    assert ThemeFontResolver.resolve("+majorBidi", theme, east_asia=False) == "Cambria"
    assert ThemeFontResolver.resolve("+minorBidi", theme, east_asia=False) == "Calibri"
    assert ThemeFontResolver.resolve("+majorEastAsia", theme, east_asia=True) == "Cambria"


def test_theme_font_resolver_never_returns_the_raw_placeholder_string() -> None:
    # Even in the pathological case where the theme has no Latin face
    # either, the unresolved "+major*"/"+minor*" placeholder must never be
    # handed back as if it were a real font name.
    theme = ThemeFonts()

    assert ThemeFontResolver.resolve("+majorBidi", theme, east_asia=False) is None
    assert ThemeFontResolver.resolve("+minorEastAsia", theme, east_asia=True) is None


def test_style_resolver_resolves_theme_complex_script_font_via_latin_fallback() -> None:
    resolver = OoxmlStyleResolver(
        styles=(
            StyleModel(
                style_id="Body",
                name="Body",
                style_type="paragraph",
                run=RunProperties(font_family="+majorBidi"),
            ),
        ),
        theme_fonts=ThemeFonts(major_latin="Cambria"),
    )

    resolved = resolver.resolve(paragraph_style_id="Body")

    assert resolved.run.font_family == "Cambria"


def test_style_resolver_resolves_theme_font_properties() -> None:
    resolver = OoxmlStyleResolver(
        styles=(
            StyleModel(
                style_id="Body",
                name="Body",
                style_type="paragraph",
                run=RunProperties(font_family="+minorHAnsi", east_asia_font="+minorEastAsia"),
            ),
        ),
        theme_fonts=ThemeFonts(
            minor_latin="Body Latin",
            minor_east_asia="Body Japanese",
        ),
    )

    resolved = resolver.resolve(paragraph_style_id="Body")

    assert resolved.run.font_family == "Body Latin"
    assert resolved.run.east_asia_font == "Body Japanese"


def test_styles_parser_reads_defaults_inheritance_and_formatting() -> None:
    root = ElementTree.fromstring(
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults>
            <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:styleId="Body" w:default="1">
            <w:name w:val="Body text"/>
            <w:basedOn w:val="Base"/>
            <w:next w:val="Body"/>
            <w:pPr>
              <w:jc w:val="center"/>
              <w:ind w:left="720" w:right="360" w:firstLine="240"/>
              <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="exact"/>
              <w:keepNext/>
              <w:tabs><w:tab w:val="right" w:pos="1440" w:leader="dot"/></w:tabs>
            </w:pPr>
            <w:rPr>
              <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>
              <w:sz w:val="24"/>
              <w:b/>
              <w:i w:val="0"/>
              <w:u w:val="single"/>
              <w:color w:val="ff00aa"/>
              <w:spacing w:val="20"/>
            </w:rPr>
          </w:style>
        </w:styles>
        """
    )

    sheet = OoxmlStylesParser.parse(root)

    assert sheet.defaults.run.font_size == 11.0
    assert sheet.defaults.paragraph.alignment == "left"
    assert len(sheet.styles) == 1
    style = sheet.styles[0]
    assert style.style_id == "Body"
    assert style.based_on == "Base"
    assert style.next_style == "Body"
    assert style.is_default is True
    assert style.paragraph.alignment == "center"
    assert style.paragraph.left_indent == 36.0
    assert style.paragraph.right_indent == 18.0
    assert style.paragraph.first_line_indent == 12.0
    assert style.paragraph.space_before == 6.0
    assert style.paragraph.space_after == 12.0
    assert style.paragraph.line_spacing == 18.0
    assert style.paragraph.line_spacing_rule == "exact"
    assert style.paragraph.keep_next is True
    assert style.paragraph.tabs[0].position == 72.0
    assert style.paragraph.tabs[0].alignment == "right"
    assert style.run.font_family == "+minorHAnsi"
    assert style.run.east_asia_font == "+minorEastAsia"
    assert style.run.font_size == 12.0
    assert style.run.bold is True
    assert style.run.italic is False
    assert style.run.underline == "single"
    assert style.run.color == "FF00AA"
    assert style.run.character_spacing == 1.0


def test_theme_font_parser_reads_major_and_minor_fonts() -> None:
    root = ElementTree.fromstring(
        """
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements>
            <a:fontScheme name="Office">
              <a:majorFont>
                <a:latin typeface="Major Latin"/><a:ea typeface="Major Japanese"/>
                <a:cs typeface="Major Complex"/>
              </a:majorFont>
              <a:minorFont>
                <a:latin typeface="Minor Latin"/><a:ea typeface="Minor Japanese"/>
                <a:cs typeface="Minor Complex"/>
              </a:minorFont>
            </a:fontScheme>
          </a:themeElements>
        </a:theme>
        """
    )

    theme = ThemeFontParser.parse(root)

    assert theme.major_latin == "Major Latin"
    assert theme.major_east_asia == "Major Japanese"
    assert theme.major_complex_script == "Major Complex"
    assert theme.minor_latin == "Minor Latin"
    assert theme.minor_east_asia == "Minor Japanese"
    assert theme.minor_complex_script == "Minor Complex"


def test_styles_parser_accepts_numbering_tab_alignment() -> None:
    root = ElementTree.fromstring(
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="Numbered">
            <w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs></w:pPr>
          </w:style>
        </w:styles>
        """
    )

    style = OoxmlStylesParser.parse(root).styles[0]

    assert style.paragraph.tabs[0].alignment == "num"


def test_styles_parser_treats_present_underline_without_value_as_single() -> None:
    root = ElementTree.fromstring(
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="character" w:styleId="Underline">
            <w:rPr><w:u/></w:rPr>
          </w:style>
        </w:styles>
        """
    )

    sheet = OoxmlStylesParser.parse(root)

    assert sheet.styles[0].run.underline == "single"


def test_parsed_absent_properties_do_not_clear_inherited_values() -> None:
    root = ElementTree.fromstring(
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="Base">
            <w:pPr><w:keepNext/></w:pPr>
            <w:rPr><w:b/><w:sz w:val="20"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Child">
            <w:basedOn w:val="Base"/>
            <w:pPr><w:jc w:val="right"/></w:pPr>
            <w:rPr><w:color w:val="123456"/></w:rPr>
          </w:style>
        </w:styles>
        """
    )
    sheet = OoxmlStylesParser.parse(root)

    resolved = OoxmlStyleResolver(styles=sheet.styles).resolve(paragraph_style_id="Child")

    assert resolved.run.bold is True
    assert resolved.run.font_size == 10.0
    assert resolved.run.color == "123456"
    assert resolved.paragraph.keep_next is True
    assert resolved.paragraph.alignment == "right"
