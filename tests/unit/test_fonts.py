from __future__ import annotations

import shutil
from pathlib import Path

import pytest
import reportlab
from fontTools.ttLib import TTFont
from reportlab.pdfbase import pdfmetrics

from docxpdf_native.exceptions import FontNotFoundError
from docxpdf_native.fonts.metrics import FontToolsTextMeasurer
from docxpdf_native.fonts.registry import FontRegistry
from docxpdf_native.fonts.resolver import DefaultFontResolver
from docxpdf_native.models.fonts import FontConfiguration


def _vera_font() -> Path:
    return Path(reportlab.__file__).parent / "fonts" / "Vera.ttf"


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
