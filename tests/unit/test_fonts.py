from __future__ import annotations

import logging
import platform
import shutil
import tempfile
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pytest
import reportlab
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTCollection, TTFont
from pypdf import PdfReader
from reportlab.pdfbase import pdfmetrics

from docxpdf_native import ConversionOptions, Converter
from docxpdf_native.exceptions import FontNotFoundError
from docxpdf_native.fonts import cjk
from docxpdf_native.fonts.bundled import bundled_font_paths, metric_compatible_substitution
from docxpdf_native.fonts.cjk import (
    contains_cjk_characters,
    detect_system_cjk_fonts,
    known_japanese_font_category,
    system_cjk_candidates,
)
from docxpdf_native.fonts.metrics import FontToolsTextMeasurer
from docxpdf_native.fonts.registry import FontRegistry
from docxpdf_native.fonts.resolver import DefaultFontResolver
from docxpdf_native.models.fonts import FontConfiguration, ResolvedFont


def _vera_font() -> Path:
    return Path(reportlab.__file__).parent / "fonts" / "Vera.ttf"


def _vera_bold_font() -> Path:
    return Path(reportlab.__file__).parent / "fonts" / "VeraBd.ttf"


def _copy_font_with_family(source: Path, destination: Path, family: str) -> None:
    with TTFont(source) as font:
        name_table = font["name"]
        name_table.names = [name for name in name_table.names if name.nameID not in {1, 4, 6, 16}]
        postscript = family.replace(" ", "")
        for platform_id, encoding_id, language_id in ((1, 0, 0), (3, 1, 0x409)):
            name_table.setName(family, 1, platform_id, encoding_id, language_id)
            name_table.setName(family, 4, platform_id, encoding_id, language_id)
            name_table.setName(postscript, 6, platform_id, encoding_id, language_id)
            name_table.setName(family, 16, platform_id, encoding_id, language_id)
        font.save(destination)


def _build_ttc(destination: Path, faces: tuple[tuple[Path, str], ...]) -> None:
    """Combine standalone TrueType faces into a ``.ttc`` collection, in order.

    The renamed intermediate ``.ttf`` files are built in a scratch directory
    *outside* ``destination``'s directory tree so they cannot be picked up
    (and shadow the collection) by a registry scan of ``destination``'s
    parent directory.
    """
    with tempfile.TemporaryDirectory() as scratch:
        renamed_paths = []
        for index, (source, family) in enumerate(faces):
            renamed = Path(scratch) / f"face{index}.ttf"
            _copy_font_with_family(source, renamed, family)
            renamed_paths.append(renamed)
        collection = TTCollection()
        opened = [TTFont(path) for path in renamed_paths]
        collection.fonts = opened
        collection.save(str(destination))
        for font in opened:
            font.close()


def _build_cff_face(destination: Path, family: str) -> None:
    """Build a minimal CFF-flavored (PostScript outline) OpenType font.

    ReportLab's TTFont backend cannot embed these; used to verify that
    collection scanning skips CFF faces rather than indexing something it
    cannot later embed.
    """
    glyph_order = [".notdef", "space", "A"]
    builder = FontBuilder(1000, isTTF=False)
    builder.setupGlyphOrder(glyph_order)
    builder.setupCharacterMap({65: "A", 32: "space"})
    charstrings = {}
    for name in glyph_order:
        pen = T2CharStringPen(500, None)
        if name == "A":
            pen.moveTo((80, 80))
            pen.lineTo((920, 80))
            pen.lineTo((920, 920))
            pen.closePath()
        charstrings[name] = pen.getCharString()
    postscript_name = family.replace(" ", "") + "-Regular"
    builder.setupCFF(postscript_name, {"FullName": family}, charstrings, {})
    builder.setupHorizontalMetrics(dict.fromkeys(glyph_order, (500, 0)))
    builder.setupHorizontalHeader(ascent=900, descent=-100)
    builder.setupNameTable({"familyName": family, "styleName": "Regular"})
    builder.setupOS2()
    builder.setupPost()
    builder.setupMaxp()
    builder.save(str(destination))


def _build_glyph_font(destination: Path, text: str, family: str) -> None:
    """Build a minimal TrueType font with real outlines for each character in ``text``.

    Lets a test confirm actual glyph embedding (extractable text, not tofu)
    rather than only checking that font *metadata* was selected.
    """
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
    postscript_name = family.replace(" ", "") + "-Regular"
    builder.setupNameTable(
        {
            "familyName": family,
            "styleName": "Regular",
            "uniqueFontIdentifier": postscript_name,
            "fullName": f"{family} Regular",
            "psName": postscript_name,
        }
    )
    builder.setupOS2(sTypoAscender=900, sTypoDescender=-100, usWinAscent=900, usWinDescent=100)
    builder.setupPost()
    builder.setupMaxp()
    builder.save(str(destination))


def _minimal_japanese_docx(text: str, *, east_asia_font: str) -> bytes:
    """Build the smallest DOCX that requests ``east_asia_font`` for ``text``."""
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    package_rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rels_content_type = "application/vnd.openxmlformats-package.relationships+xml"
    document_content_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    )
    styles_content_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
    )
    office_document_rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    )
    content_types = f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="{rels_content_type}"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="{document_content_type}"/>
  <Override PartName="/word/styles.xml" ContentType="{styles_content_type}"/>
</Types>""".encode()
    root_rels = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{package_rels_ns}">
  <Relationship Id="rId1" Type="{office_document_rel_type}" Target="word/document.xml"/>
</Relationships>""".encode()
    styles = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{w_ns}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Helvetica" w:hAnsi="Helvetica" w:eastAsia="{east_asia_font}"/>
    <w:sz w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>""".encode()
    document = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{w_ns}">
  <w:body>
    <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>""".encode()
    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "word/document.xml": document,
        "word/styles.xml": styles,
    }
    output = BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    return output.getvalue()


def test_registry_resolves_reportlab_standard_font_without_a_path() -> None:
    record = FontRegistry(include_system_fonts=False).resolve("helvetica")

    assert record is not None
    assert record.family == "Helvetica"
    assert record.path is None
    assert record.source == "reportlab-standard"


def test_registry_resolves_explicit_alias_case_insensitively() -> None:
    record = FontRegistry(
        registered_fonts={"My Body Font": _vera_font()}, include_system_fonts=False
    ).resolve("my body font")

    assert record is not None
    assert record.path == _vera_font().resolve()
    assert record.postscript_name == "BitstreamVeraSans-Roman"
    assert record.source == "registered"


def test_registry_scans_configured_directory(tmp_path: Path) -> None:
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    copied = font_directory / "body.ttf"
    shutil.copyfile(_vera_font(), copied)

    record = FontRegistry(font_directories=(font_directory,), include_system_fonts=False).resolve(
        "Bitstream Vera Sans"
    )

    assert record is not None
    assert record.path == copied.resolve()
    assert record.source == "font-directory"


def test_configured_directory_has_priority_over_standard_font(tmp_path: Path) -> None:
    configured = tmp_path / "Helvetica.ttf"
    _copy_font_with_family(_vera_font(), configured, "Helvetica")

    record = FontRegistry(font_directories=(tmp_path,), include_system_fonts=False).resolve(
        "Helvetica"
    )

    assert record is not None
    assert record.path == configured.resolve()
    assert record.source == "font-directory"


def test_registry_reads_environment_directories(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    font_directory = tmp_path / "environment-fonts"
    font_directory.mkdir()
    copied = font_directory / "body.ttf"
    shutil.copyfile(_vera_font(), copied)
    monkeypatch.setenv("TEST_DOCX_FONTS", str(font_directory))

    record = FontRegistry(
        environment_variable="TEST_DOCX_FONTS", include_system_fonts=False
    ).resolve("Bitstream Vera Sans")

    assert record is not None
    assert record.path == copied.resolve()
    assert record.source == "environment"


def test_registry_ignores_invalid_files_found_while_scanning(tmp_path: Path) -> None:
    invalid = tmp_path / "invalid.ttf"
    invalid.write_bytes(b"not-a-font")

    registry = FontRegistry(font_directories=(tmp_path,), include_system_fonts=False)

    assert registry.resolve("Missing Font") is None


def test_registry_rejects_invalid_explicit_font() -> None:
    with pytest.raises(ValueError, match="registered font"):
        FontRegistry(registered_fonts={"Broken": Path("missing.ttf")}, include_system_fonts=False)


def test_resolver_uses_an_explicit_substitution_in_strict_mode(tmp_path: Path) -> None:
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    shutil.copyfile(_vera_font(), font_directory / "body.ttf")
    resolver = DefaultFontResolver(
        FontConfiguration(
            font_directories=(font_directory,),
            substitutions={"Unavailable Body": "Bitstream Vera Sans"},
            include_system_fonts=False,
        ),
        strict=True,
    )

    resolved = resolver.resolve("Unavailable Body", east_asia=True, paragraph_index=2, run_index=3)

    assert resolved.family == "Bitstream Vera Sans"
    assert resolved.path == (font_directory / "body.ttf").resolve()
    assert resolved.substituted is True
    assert len(resolver.substitutions) == 1
    assert resolver.substitutions[0].requested_font == "Unavailable Body"
    assert resolver.substitutions[0].east_asia is True


def test_resolver_raises_for_missing_font_in_strict_mode() -> None:
    resolver = DefaultFontResolver(
        FontConfiguration(include_system_fonts=False),
        strict=True,
    )

    with pytest.raises(FontNotFoundError, match="Unavailable Body"):
        resolver.resolve("Unavailable Body")


def test_resolver_uses_configured_default_in_lenient_mode() -> None:
    resolver = DefaultFontResolver(
        FontConfiguration(default_font="Helvetica", include_system_fonts=False),
        strict=False,
    )

    resolved = resolver.resolve("Unavailable Body", paragraph_index=4, run_index=5)

    assert resolved.family == "Helvetica"
    assert resolved.substituted is True
    assert resolver.substitutions[0].reason == "lenient-default"
    assert resolver.substitutions[0].paragraph_index == 4
    assert resolver.substitutions[0].run_index == 5


def test_resolver_rejects_missing_substitution_target() -> None:
    resolver = DefaultFontResolver(
        FontConfiguration(
            substitutions={"Unavailable Body": "Also Missing"}, include_system_fonts=False
        ),
        strict=False,
    )

    with pytest.raises(FontNotFoundError, match="Also Missing"):
        resolver.resolve("Unavailable Body")


def test_text_measurer_uses_standard_font_metrics() -> None:
    font = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=True).resolve(
        "Helvetica"
    )

    measurement = FontToolsTextMeasurer().measure("Hello", font, 12.0)

    assert measurement.width == pytest.approx(pdfmetrics.stringWidth("Hello", "Helvetica", 12.0))
    assert measurement.ascent > 0
    assert measurement.descent >= 0


def test_text_measurer_uses_fonttools_glyph_widths() -> None:
    font = DefaultFontResolver(
        FontConfiguration(registered_fonts={"Vera": _vera_font()}, include_system_fonts=False),
        strict=True,
    ).resolve("Vera")
    measurer = FontToolsTextMeasurer()

    wide = measurer.measure("WWW", font, 12.0)
    narrow = measurer.measure("iii", font, 12.0)

    assert wide.width > narrow.width * 2
    assert wide.ascent > 0
    assert wide.descent > 0


def test_text_measurer_scales_metrics_with_font_size() -> None:
    font = DefaultFontResolver(
        FontConfiguration(registered_fonts={"Vera": _vera_font()}, include_system_fonts=False),
        strict=True,
    ).resolve("Vera")
    measurer = FontToolsTextMeasurer()

    small = measurer.measure("Scale", font, 10.0)
    large = measurer.measure("Scale", font, 20.0)

    assert large.width == pytest.approx(small.width * 2)
    assert large.ascent == pytest.approx(small.ascent * 2)
    assert large.descent == pytest.approx(small.descent * 2)


def test_text_measurer_applies_spacing_between_grapheme_clusters() -> None:
    font = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=True).resolve(
        "Helvetica"
    )
    measurer = FontToolsTextMeasurer()

    unspaced = measurer.measure("A\N{COMBINING ACUTE ACCENT}B", font, 12.0)
    spaced = measurer.measure("A\N{COMBINING ACUTE ACCENT}B", font, 12.0, character_spacing=2.0)

    assert spaced.width == pytest.approx(unspaced.width + 2.0)


def test_text_measurer_clamps_extreme_negative_spacing_to_zero() -> None:
    font = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=True).resolve(
        "Helvetica"
    )

    measurement = FontToolsTextMeasurer().measure("AB", font, 12.0, character_spacing=-1_000.0)

    assert measurement.width == 0.0


def test_text_measurement_cache_is_instance_local() -> None:
    font = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=True).resolve(
        "Helvetica"
    )
    first_measurer = FontToolsTextMeasurer()
    second_measurer = FontToolsTextMeasurer()

    first = first_measurer.measure("cached", font, 12.0)
    repeated = first_measurer.measure("cached", font, 12.0)
    independent = second_measurer.measure("cached", font, 12.0)

    assert repeated is first
    assert independent is not first


# --- TrueType/OpenType Collection (.ttc/.otc) scanning -----------------------------------


def test_registry_scans_ttc_collection_and_indexes_each_face_by_number(tmp_path: Path) -> None:
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    ttc_path = font_directory / "bundle.ttc"
    _build_ttc(
        ttc_path,
        (
            (_vera_font(), "Test Collection Serif"),
            (_vera_bold_font(), "Test Collection Sans"),
        ),
    )
    registry = FontRegistry(font_directories=(font_directory,), include_system_fonts=False)

    serif = registry.resolve("Test Collection Serif")
    sans = registry.resolve("test collection sans")

    assert serif is not None
    assert serif.path == ttc_path.resolve()
    assert serif.font_number == 0
    assert serif.source == "font-directory"
    assert sans is not None
    assert sans.path == ttc_path.resolve()
    assert sans.font_number == 1


def test_registry_otc_suffix_is_scanned_like_ttc(tmp_path: Path) -> None:
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    otc_path = font_directory / "bundle.otc"
    _build_ttc(otc_path, ((_vera_font(), "Test OTC Face"),))
    registry = FontRegistry(font_directories=(font_directory,), include_system_fonts=False)

    record = registry.resolve("Test OTC Face")

    assert record is not None
    assert record.path == otc_path.resolve()
    assert record.font_number == 0


def test_registry_skips_cff_faces_within_a_collection(tmp_path: Path) -> None:
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    ttc_path = font_directory / "mixed.ttc"
    cff_face = tmp_path / "cff-face.otf"
    truetype_face = tmp_path / "truetype-face.ttf"
    _build_cff_face(cff_face, "Mixed CFF Face")
    _copy_font_with_family(_vera_font(), truetype_face, "Mixed TrueType Face")
    collection = TTCollection()
    opened = [TTFont(str(truetype_face)), TTFont(str(cff_face))]
    collection.fonts = opened
    collection.save(str(ttc_path))
    for font in opened:
        font.close()
    registry = FontRegistry(font_directories=(font_directory,), include_system_fonts=False)

    truetype_record = registry.resolve("Mixed TrueType Face")
    cff_record = registry.resolve("Mixed CFF Face")

    assert truetype_record is not None
    assert truetype_record.font_number == 0
    assert cff_record is None


def test_registry_collection_records_helper_returns_empty_for_all_cff_collection(
    tmp_path: Path,
) -> None:
    ttc_path = tmp_path / "all-cff.ttc"
    cff_a = tmp_path / "cff-a.otf"
    cff_b = tmp_path / "cff-b.otf"
    _build_cff_face(cff_a, "All CFF A")
    _build_cff_face(cff_b, "All CFF B")
    collection = TTCollection()
    opened = [TTFont(str(cff_a)), TTFont(str(cff_b))]
    collection.fonts = opened
    collection.save(str(ttc_path))
    for font in opened:
        font.close()

    records = FontRegistry._collection_records(ttc_path, source="test")

    assert records == ()


@pytest.mark.skipif(
    platform.system() != "Darwin"
    or not Path("/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc").is_file(),
    reason="depends on the real Hiragino Kaku Gothic collection shipped with macOS",
)
def test_registry_excludes_all_faces_of_the_real_hiragino_collection_on_macos() -> None:
    # Apple ships Hiragino as CFF-flavored (PostScript outline) faces inside a
    # TrueType Collection wrapper. ReportLab's TTFont backend cannot embed
    # PostScript outlines, so every face must be excluded deterministically
    # rather than indexing something that would blow up at PDF generation.
    path = Path("/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc")

    records = FontRegistry._collection_records(path, source="system")

    assert records == ()


def test_registry_resolve_face_recovers_collection_font_number(tmp_path: Path) -> None:
    ttc_path = tmp_path / "bundle.ttc"
    _build_ttc(
        ttc_path,
        (
            (_vera_font(), "Resolve Face A"),
            (_vera_bold_font(), "Resolve Face B"),
        ),
    )

    face_a = FontRegistry.resolve_face(ttc_path, "Resolve Face A")
    face_b = FontRegistry.resolve_face(ttc_path, "resolve face b")
    missing = FontRegistry.resolve_face(ttc_path, "Nonexistent Face")

    assert face_a is not None
    assert face_a.font_number == 0
    assert face_b is not None
    assert face_b.font_number == 1
    assert missing is None


def test_registry_resolve_face_handles_a_standalone_font_file() -> None:
    face = FontRegistry.resolve_face(_vera_font(), "Bitstream Vera Sans")

    assert face is not None
    assert face.font_number is None
    assert face.path == _vera_font().resolve()


def test_registered_font_pointing_at_a_collection_uses_first_embeddable_face(
    tmp_path: Path,
) -> None:
    ttc_path = tmp_path / "bundle.ttc"
    _build_ttc(
        ttc_path,
        (
            (_vera_font(), "Alias Face A"),
            (_vera_bold_font(), "Alias Face B"),
        ),
    )

    registry = FontRegistry(registered_fonts={"My Alias": ttc_path}, include_system_fonts=False)
    record = registry.resolve("my alias")

    assert record is not None
    assert record.font_number == 0
    assert record.family == "Alias Face A"
    # Registering an alias only indexes the one face picked for that alias
    # (deterministically the first embeddable face); it does not implicitly
    # register every other face in the collection.
    assert registry.resolve("Alias Face A") is not None
    assert registry.resolve("Alias Face B") is None


def test_registry_rejects_registered_collection_with_no_embeddable_faces(tmp_path: Path) -> None:
    ttc_path = tmp_path / "all-cff.ttc"
    cff_a = tmp_path / "cff-a.otf"
    _build_cff_face(cff_a, "Only CFF")
    collection = TTCollection()
    opened = [TTFont(str(cff_a))]
    collection.fonts = opened
    collection.save(str(ttc_path))
    for font in opened:
        font.close()

    with pytest.raises(ValueError, match="registered font"):
        FontRegistry(registered_fonts={"Broken": ttc_path}, include_system_fonts=False)


# --- Text measurement of a specific collection face --------------------------------------


def test_text_measurer_reads_the_correct_face_from_a_collection(tmp_path: Path) -> None:
    ttc_path = tmp_path / "bundle.ttc"
    _build_ttc(
        ttc_path,
        (
            (_vera_font(), "Metrics Regular Face"),
            (_vera_bold_font(), "Metrics Bold Face"),
        ),
    )
    measurer = FontToolsTextMeasurer()
    regular_standalone = measurer.measure(
        "Testing Width", ResolvedFont(family="Vera", path=_vera_font(), source="test"), 12.0
    )
    bold_standalone = measurer.measure(
        "Testing Width", ResolvedFont(family="VeraBd", path=_vera_bold_font(), source="test"), 12.0
    )
    assert regular_standalone.width != bold_standalone.width

    regular_face = FontRegistry.resolve_face(ttc_path, "Metrics Regular Face")
    bold_face = FontRegistry.resolve_face(ttc_path, "Metrics Bold Face")
    assert regular_face is not None
    assert bold_face is not None

    regular_from_collection = measurer.measure(
        "Testing Width",
        ResolvedFont(
            family="Metrics Regular Face",
            path=ttc_path,
            source="test",
            font_number=regular_face.font_number,
        ),
        12.0,
    )
    bold_from_collection = measurer.measure(
        "Testing Width",
        ResolvedFont(
            family="Metrics Bold Face",
            path=ttc_path,
            source="test",
            font_number=bold_face.font_number,
        ),
        12.0,
    )

    assert regular_from_collection.width == pytest.approx(regular_standalone.width)
    assert bold_from_collection.width == pytest.approx(bold_standalone.width)
    assert regular_from_collection.width != bold_from_collection.width


def test_text_measurer_recovers_face_when_font_number_is_missing(tmp_path: Path) -> None:
    """A ResolvedFont rebuilt further down the pipeline may lack font_number.

    The measurer must still read the face matching ``family``, not silently
    default to face 0, by re-deriving the index from ``(path, family)``.
    """
    ttc_path = tmp_path / "bundle.ttc"
    _build_ttc(
        ttc_path,
        (
            (_vera_font(), "Recover Regular Face"),
            (_vera_bold_font(), "Recover Bold Face"),
        ),
    )
    measurer = FontToolsTextMeasurer()
    bold_standalone = measurer.measure(
        "Testing Width", ResolvedFont(family="VeraBd", path=_vera_bold_font(), source="test"), 12.0
    )

    # font_number intentionally omitted, mirroring how the layout engine
    # reconstructs a ResolvedFont from just a family name and path.
    reconstructed = ResolvedFont(family="Recover Bold Face", path=ttc_path, source="test")
    measured = FontToolsTextMeasurer().measure("Testing Width", reconstructed, 12.0)

    assert measured.width == pytest.approx(bold_standalone.width)


# --- Built-in Japanese font name -> category map (fonts/cjk.py) --------------------------


@pytest.mark.parametrize(
    ("name", "expected_category"),
    [
        ("MS Mincho", "serif"),
        ("MS 明朝", "serif"),
        ("ＭＳ 明朝", "serif"),  # noqa: RUF001 - full-width spelling variant under test
        ("MS P明朝", "serif"),
        ("游明朝", "serif"),
        ("Yu Mincho", "serif"),
        ("MS Gothic", "sans"),
        ("MS ゴシック", "sans"),
        ("ＭＳ ゴシック", "sans"),  # noqa: RUF001 - full-width spelling variant under test
        ("MS Pゴシック", "sans"),
        ("游ゴシック", "sans"),
        ("Yu Gothic", "sans"),
        ("メイリオ", "sans"),
        ("Meiryo", "sans"),
        ("MS UI Gothic", "sans"),
    ],
)
def test_known_japanese_font_category_matches_documented_names(
    name: str, expected_category: str
) -> None:
    assert known_japanese_font_category(name) == expected_category


def test_known_japanese_font_category_normalizes_fullwidth_space_between_words() -> None:
    # Full-width "MS" plus U+3000 IDEOGRAPHIC SPACE before the Japanese word.
    fullwidth_variant = "ＭＳ　ゴシック"  # noqa: RUF001 - full-width spelling under test

    assert known_japanese_font_category(fullwidth_variant) == "sans"


def test_known_japanese_font_category_returns_none_for_unrelated_fonts() -> None:
    assert known_japanese_font_category("Arial") is None
    assert known_japanese_font_category("Times New Roman") is None
    assert known_japanese_font_category("Noto Sans CJK JP") is None


def test_contains_cjk_characters_detects_kana_and_kanji() -> None:
    assert contains_cjk_characters("日本語") is True
    assert contains_cjk_characters("ひらがな") is True
    assert contains_cjk_characters("カタカナ") is True
    assert contains_cjk_characters("Hello World") is False
    assert contains_cjk_characters("") is False


def test_system_cjk_candidates_are_deterministic_and_categorized() -> None:
    first = system_cjk_candidates()
    second = system_cjk_candidates()

    assert first == second
    assert set(first) <= {"serif", "sans"}
    assert all(isinstance(names, tuple) and names for names in first.values())


def test_detect_system_cjk_fonts_picks_first_resolvable_candidate_per_category(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(cjk.platform, "system", lambda: "Linux")
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    _copy_font_with_family(_vera_font(), font_directory / "serif.ttf", "Noto Serif CJK JP")
    _copy_font_with_family(_vera_font(), font_directory / "sans.ttf", "Noto Sans CJK JP")
    registry = FontRegistry(font_directories=(font_directory,), include_system_fonts=False)

    detected = detect_system_cjk_fonts(registry)

    assert detected["serif"].family == "Noto Serif CJK JP"
    assert detected["sans"].family == "Noto Sans CJK JP"


def test_detect_system_cjk_fonts_returns_empty_when_nothing_is_available() -> None:
    registry = FontRegistry(include_system_fonts=False)

    assert detect_system_cjk_fonts(registry) == {}


# --- DefaultFontResolver: built-in Japanese substitution + CJK-aware fallback ------------


def _resolver_with_detected_cjk_fonts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    serif: bool = True,
    sans: bool = True,
    strict: bool,
    default_font: str | None = None,
) -> DefaultFontResolver:
    monkeypatch.setattr(cjk.platform, "system", lambda: "Linux")
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    if serif:
        _copy_font_with_family(_vera_font(), font_directory / "serif.ttf", "Noto Serif CJK JP")
    if sans:
        _copy_font_with_family(_vera_font(), font_directory / "sans.ttf", "Noto Sans CJK JP")
    return DefaultFontResolver(
        FontConfiguration(
            font_directories=(font_directory,),
            include_system_fonts=False,
            default_font=default_font,
        ),
        strict=strict,
    )


def test_resolver_applies_builtin_cjk_substitution_even_in_strict_mode(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(tmp_path, monkeypatch, strict=True)

    resolved = resolver.resolve("MS Gothic", east_asia=True, paragraph_index=1, run_index=0)

    assert resolved.family == "Noto Sans CJK JP"
    assert resolved.substituted is True
    assert resolver.substitutions[0].reason == "builtin-cjk"
    assert resolver.substitutions[0].requested_font == "MS Gothic"


def test_resolver_builtin_cjk_substitution_maps_mincho_to_serif_and_gothic_to_sans(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(tmp_path, monkeypatch, strict=True)

    mincho = resolver.resolve("MS 明朝")
    gothic = resolver.resolve("MS ゴシック")

    assert mincho.family == "Noto Serif CJK JP"
    assert gothic.family == "Noto Sans CJK JP"


def test_resolver_builtin_cjk_substitution_normalizes_fullwidth_spelling(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(tmp_path, monkeypatch, strict=True)

    resolved = resolver.resolve("ＭＳ　ゴシック")  # noqa: RUF001 - full-width spelling under test

    assert resolved.family == "Noto Sans CJK JP"
    assert resolved.substituted is True


def test_resolver_still_raises_in_strict_mode_for_unknown_non_cjk_font(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(tmp_path, monkeypatch, strict=True)

    with pytest.raises(FontNotFoundError, match="Totally Unknown Font"):
        resolver.resolve("Totally Unknown Font")


def test_resolver_lenient_fallback_uses_system_cjk_font_for_east_asia_flag(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(tmp_path, monkeypatch, strict=False, sans=True)

    resolved = resolver.resolve("Some Custom Body Font", east_asia=True)

    assert resolved.family == "Noto Sans CJK JP"
    assert resolved.substituted is True
    assert resolver.substitutions[0].reason == "lenient-default"


def test_resolver_lenient_fallback_uses_system_cjk_font_for_cjk_characters_in_name(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(tmp_path, monkeypatch, strict=False)

    resolved = resolver.resolve("日本語フォント")

    assert resolved.family == "Noto Sans CJK JP"
    assert resolved.substituted is True


def test_resolver_explicit_default_font_has_priority_over_cjk_auto_fallback(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resolver = _resolver_with_detected_cjk_fonts(
        tmp_path, monkeypatch, strict=False, default_font="Helvetica"
    )

    resolved = resolver.resolve("Some Custom Body Font", east_asia=True)

    assert resolved.family == "Helvetica"
    assert resolver.substitutions[0].reason == "lenient-default"


def test_resolver_lenient_fallback_warns_and_uses_helvetica_when_no_cjk_font_detected(
    caplog: pytest.LogCaptureFixture,
) -> None:
    resolver = DefaultFontResolver(
        FontConfiguration(include_system_fonts=False),
        strict=False,
    )

    with caplog.at_level(logging.WARNING, logger="docxpdf_native.fonts.resolver"):
        resolved = resolver.resolve("MS Gothic", east_asia=True)

    assert resolved.family == "Helvetica"
    assert resolved.substituted is True
    assert any("Helvetica" in record.message for record in caplog.records)


# --- End-to-end: Japanese DOCX converts with zero font configuration ---------------------


def test_converter_converts_japanese_docx_with_default_options_and_records_substitution() -> None:
    """Reproduces the reported problem: no font_configuration, no fonts/ directory.

    The document requests "MS Gothic" (never installed on a conversion
    machine); with the default ConversionOptions() -- strict mode, no
    registered_fonts, no font_directories -- the conversion must still
    succeed by way of the built-in Japanese substitution map and a detected
    system CJK font, and that substitution must be recorded.
    """
    registry = FontRegistry()
    if not detect_system_cjk_fonts(registry):
        pytest.skip("no system CJK font is available in this environment")

    text = "日本語の文書"
    source = _minimal_japanese_docx(text, east_asia_font="MS Gothic")

    result = Converter(ConversionOptions()).convert(source)

    assert result.page_count >= 1
    assert result.font_substitutions
    assert result.font_substitutions[0].requested_font == "MS Gothic"
    assert result.font_substitutions[0].reason in {"builtin-cjk", "lenient-default"}


def test_converter_embeds_real_japanese_glyphs_via_builtin_cjk_substitution(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """The substituted system CJK font must produce extractable, non-tofu text."""
    monkeypatch.setattr(cjk.platform, "system", lambda: "Linux")
    text = "日本語の文書"
    font_directory = tmp_path / "fonts"
    font_directory.mkdir()
    _build_glyph_font(font_directory / "sans.ttf", text, "Noto Sans CJK JP")
    source = _minimal_japanese_docx(text, east_asia_font="MS Gothic")

    options = ConversionOptions(
        strict=True,
        font_configuration=FontConfiguration(
            font_directories=(font_directory,),
            include_system_fonts=False,
        ),
    )

    result = Converter(options).convert(source)

    extracted = PdfReader(BytesIO(result.pdf_bytes or b"")).pages[0].extract_text()
    assert text in extracted
    assert result.font_substitutions
    assert result.font_substitutions[0].reason == "builtin-cjk"
    assert result.font_substitutions[0].selected_font == "Noto Sans CJK JP"


# --- Bundled metric-compatible fonts (Carlito for Calibri, Caladea for Cambria) ----------


def test_bundled_font_paths_are_deterministic_and_complete() -> None:
    first = bundled_font_paths()
    second = bundled_font_paths()

    assert first == second
    names = [path.name for path in first]
    assert names == [
        "Carlito-Regular.ttf",
        "Carlito-Bold.ttf",
        "Carlito-Italic.ttf",
        "Carlito-BoldItalic.ttf",
        "Caladea-Regular.ttf",
        "Caladea-Bold.ttf",
        "Caladea-Italic.ttf",
        "Caladea-BoldItalic.ttf",
    ]
    assert all(path.is_file() for path in first)


def test_bundled_font_directories_ship_their_ofl_license() -> None:
    family_directories = {path.parent for path in bundled_font_paths()}

    assert family_directories
    for directory in family_directories:
        assert (directory / "OFL.txt").is_file(), f"missing OFL.txt next to fonts in {directory}"


def test_registry_resolves_bundled_families_to_their_regular_faces() -> None:
    registry = FontRegistry(include_system_fonts=False)

    carlito = registry.resolve("Carlito")
    caladea = registry.resolve("caladea")

    assert carlito is not None
    assert carlito.source == "bundled"
    assert carlito.postscript_name == "Carlito-Regular"
    assert caladea is not None
    assert caladea.source == "bundled"
    assert caladea.postscript_name == "Caladea-Regular"


def test_registry_resolves_bundled_styled_faces_by_full_name() -> None:
    registry = FontRegistry(include_system_fonts=False)

    bold = registry.resolve("Carlito Bold")
    bold_italic = registry.resolve("caladea bold italic")

    assert bold is not None
    assert bold.postscript_name == "Carlito-Bold"
    assert bold_italic is not None
    assert bold_italic.postscript_name == "Caladea-BoldItalic"


def test_configured_directory_has_priority_over_bundled_group(tmp_path: Path) -> None:
    override = tmp_path / "Carlito.ttf"
    _copy_font_with_family(_vera_font(), override, "Carlito")
    registry = FontRegistry(font_directories=(tmp_path,), include_system_fonts=False)

    record = registry.resolve("Carlito")

    assert record is not None
    assert record.source == "font-directory"
    assert record.path == override.resolve()


def test_bundled_group_still_resolves_standard_fonts_after_it() -> None:
    registry = FontRegistry(include_system_fonts=False)

    record = registry.resolve("Helvetica")

    assert record is not None
    assert record.source == "reportlab-standard"


@pytest.mark.parametrize(
    ("requested", "expected_target"),
    [
        ("Calibri", "Carlito"),
        ("calibri", "Carlito"),
        ("Calibri Light", "Carlito"),
        ("CALIBRI BOLD", "Carlito Bold"),
        ("Calibri Italic", "Carlito Italic"),
        ("Calibri Light Italic", "Carlito Italic"),
        ("Calibri Bold Italic", "Carlito Bold Italic"),
        ("Cambria", "Caladea"),
        ("Cambria Math", "Caladea"),
        ("Cambria Bold", "Caladea Bold"),
        ("Cambria Italic", "Caladea Italic"),
        ("Cambria Bold Italic", "Caladea Bold Italic"),
    ],
)
def test_metric_compatible_substitution_maps_word_defaults(
    requested: str, expected_target: str
) -> None:
    assert metric_compatible_substitution(requested) == expected_target


def test_metric_compatible_substitution_ignores_unrelated_fonts() -> None:
    assert metric_compatible_substitution("Arial") is None
    assert metric_compatible_substitution("Helvetica") is None
    assert metric_compatible_substitution("Carlito") is None
    assert metric_compatible_substitution("MS Gothic") is None


def test_resolver_applies_metric_compatible_substitution_even_in_strict_mode() -> None:
    resolver = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=True)

    resolved = resolver.resolve("Calibri", paragraph_index=3, run_index=1)

    assert resolved.family == "Carlito"
    assert resolved.postscript_name == "Carlito-Regular"
    assert resolved.substituted is True
    assert len(resolver.substitutions) == 1
    substitution = resolver.substitutions[0]
    assert substitution.requested_font == "Calibri"
    assert substitution.selected_font == "Carlito"
    assert substitution.reason == "builtin-metric-compatible"
    assert substitution.paragraph_index == 3
    assert substitution.run_index == 1


def test_resolver_metric_compatible_substitution_selects_styled_face() -> None:
    resolver = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=False)

    resolved = resolver.resolve("Cambria Bold")

    assert resolved.family == "Caladea"
    assert resolved.postscript_name == "Caladea-Bold"
    assert resolver.substitutions[0].reason == "builtin-metric-compatible"


def test_resolver_prefers_an_actually_installed_calibri_over_the_bundled_substitute(
    tmp_path: Path,
) -> None:
    installed = tmp_path / "Calibri.ttf"
    _copy_font_with_family(_vera_font(), installed, "Calibri")
    resolver = DefaultFontResolver(
        FontConfiguration(font_directories=(tmp_path,), include_system_fonts=False),
        strict=True,
    )

    resolved = resolver.resolve("Calibri")

    assert resolved.family == "Calibri"
    assert resolved.path == installed.resolve()
    assert resolved.substituted is False
    assert resolver.substitutions == ()


def test_resolver_user_substitution_beats_builtin_metric_compatible_map() -> None:
    resolver = DefaultFontResolver(
        FontConfiguration(substitutions={"Calibri": "Helvetica"}, include_system_fonts=False),
        strict=True,
    )

    resolved = resolver.resolve("Calibri")

    assert resolved.family == "Helvetica"
    assert resolver.substitutions[0].reason == "configured"


def test_bundled_carlito_is_narrower_than_helvetica_for_typical_text() -> None:
    """The reason Carlito is bundled: Helvetica substitution widened Calibri
    text enough to add ~31% more pages on a real-world document."""
    resolver = DefaultFontResolver(FontConfiguration(include_system_fonts=False), strict=True)
    carlito = resolver.resolve("Calibri")
    helvetica = resolver.resolve("Helvetica")
    measurer = FontToolsTextMeasurer()
    sample = "Software Safety Standard requirements shall be documented"

    carlito_width = measurer.measure(sample, carlito, 11.0).width
    helvetica_width = measurer.measure(sample, helvetica, 11.0).width

    assert 0 < carlito_width < helvetica_width


# --- Regular face wins over styled faces that share the family name ----------------------


def test_scanned_family_name_prefers_regular_face_over_styled_sorted_earlier(
    tmp_path: Path,
) -> None:
    """Every face of a family carries the family name; the bare family name
    must resolve to the Regular face even when a styled file sorts first.

    Mirrors macOS supplemental fonts, where "Times New Roman Bold Italic.ttf"
    sorts before "Times New Roman.ttf" and used to claim the family key,
    silently measuring and drawing all regular body text with bold-italic
    metrics.
    """
    bold_first = tmp_path / "aaa-bold.ttf"  # sorts before the regular file
    regular_last = tmp_path / "zzz-regular.ttf"
    _copy_font_with_family(_vera_bold_font(), bold_first, "Style Probe")
    _copy_font_with_family(_vera_font(), regular_last, "Style Probe")
    registry = FontRegistry(font_directories=(tmp_path,), include_system_fonts=False)

    record = registry.resolve("Style Probe")

    assert record is not None
    assert record.path == regular_last.resolve()
    assert record.bold is False
    assert record.italic is False


def test_scanned_styled_face_still_wins_when_no_regular_face_exists(tmp_path: Path) -> None:
    only_bold = tmp_path / "only-bold.ttf"
    _copy_font_with_family(_vera_bold_font(), only_bold, "Bold Only Family")
    registry = FontRegistry(font_directories=(tmp_path,), include_system_fonts=False)

    record = registry.resolve("Bold Only Family")

    assert record is not None
    assert record.path == only_bold.resolve()
    assert record.bold is True


def test_font_record_style_flags_reflect_mac_style_bits() -> None:
    regular = FontRegistry._font_record(_vera_font(), source="test")
    bold = FontRegistry._font_record(_vera_bold_font(), source="test")

    assert regular is not None
    assert regular.styled is False
    assert bold is not None
    assert bold.bold is True
    assert bold.styled is True


def test_resolve_face_prefers_regular_face_in_a_collection(tmp_path: Path) -> None:
    ttc_path = tmp_path / "family.ttc"
    with tempfile.TemporaryDirectory() as scratch:
        bold_face = Path(scratch) / "bold.ttf"
        regular_face = Path(scratch) / "regular.ttf"
        _copy_font_with_family(_vera_bold_font(), bold_face, "Shared Family")
        _copy_font_with_family(_vera_font(), regular_face, "Shared Family")
        collection = TTCollection()
        opened = [TTFont(str(bold_face)), TTFont(str(regular_face))]  # bold face first
        collection.fonts = opened
        collection.save(str(ttc_path))
        for font in opened:
            font.close()

    face = FontRegistry.resolve_face(ttc_path, "Shared Family")

    assert face is not None
    assert face.font_number == 1  # the regular face, not the bold one at index 0
    assert face.styled is False
