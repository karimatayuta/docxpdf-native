from __future__ import annotations

import json
from pathlib import Path

import pytest

from docxpdf_native import __version__
from docxpdf_native.exceptions import (
    FontNotFoundError,
    InvalidDocxError,
    InvalidOoxmlError,
    LayoutError,
    MissingPartError,
    PdfGenerationError,
    RelationshipError,
    ResourceLimitError,
    UnsupportedFeatureError,
)
from docxpdf_native.models import (
    ConversionResult,
    ConversionWarning,
    DiagnosticInfo,
    LayoutDocument,
    UnsupportedFeature,
)


def _paths(tmp_path: Path) -> tuple[Path, Path]:
    source = tmp_path / "input.docx"
    source.write_bytes(b"fixture")
    return source, tmp_path / "output.pdf"


def _install_converter(
    monkeypatch: pytest.MonkeyPatch,
    *,
    result: ConversionResult | None = None,
    layout: LayoutDocument | None = None,
    error: Exception | None = None,
) -> dict[str, object]:
    from docxpdf_native import cli

    calls: dict[str, object] = {}

    class FakeConverter:
        def __init__(self, options: object) -> None:
            calls["options"] = options
            self.last_layout = layout

        def convert(self, source: Path, destination: Path) -> ConversionResult:
            calls["source"] = source
            calls["destination"] = destination
            if error is not None:
                raise error
            return result or ConversionResult(page_count=1, destination=destination)

    def create(options: object) -> FakeConverter:
        return FakeConverter(options)

    monkeypatch.setattr(cli, "_create_converter", create)
    return calls


def test_help_returns_success_and_prints_usage(capsys: pytest.CaptureFixture[str]) -> None:
    from docxpdf_native.cli import main

    exit_code = main(["--help"])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert "usage: docxpdf-native" in captured.out
    assert "--font-substitution" in captured.out
    assert captured.err == ""


def test_version_returns_success_and_prints_version(capsys: pytest.CaptureFixture[str]) -> None:
    from docxpdf_native.cli import main

    exit_code = main(["--version"])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert captured.out.strip() == f"docxpdf-native {__version__}"
    assert captured.err == ""


@pytest.mark.parametrize("arguments", [[], ["input.docx"]])
def test_required_arguments_return_input_error(
    arguments: list[str], capsys: pytest.CaptureFixture[str]
) -> None:
    from docxpdf_native.cli import main

    exit_code = main(arguments)

    assert exit_code == 2
    assert "error:" in capsys.readouterr().err


def test_cli_builds_options_and_invokes_converter(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native.cli import main
    from docxpdf_native.models import ConversionOptions

    source, destination = _paths(tmp_path)
    first_fonts = tmp_path / "fonts-a"
    second_fonts = tmp_path / "fonts-b"
    first_fonts.mkdir()
    second_fonts.mkdir()
    calls = _install_converter(
        monkeypatch,
        result=ConversionResult(page_count=3, destination=destination),
    )

    exit_code = main(
        [
            str(source),
            "--output",
            str(destination),
            "--font-dir",
            str(first_fonts),
            "--font-dir",
            str(second_fonts),
            "--font-substitution",
            "MS Mincho=Noto Serif CJK JP",
            "--font-substitution",
            "MS Gothic=Noto Sans CJK JP",
            "--lenient",
            "--deterministic",
            "--max-pages",
            "25",
        ]
    )

    options = calls["options"]
    assert isinstance(options, ConversionOptions)
    assert exit_code == 0
    assert calls["source"] == source
    assert calls["destination"] == destination
    assert options.strict is False
    assert options.deterministic is True
    assert options.font_configuration.font_directories == (first_fonts, second_fonts)
    assert options.font_configuration.substitutions == {
        "MS Mincho": "Noto Serif CJK JP",
        "MS Gothic": "Noto Sans CJK JP",
    }
    assert options.resource_limits.max_pages == 25
    assert "3 pages" in capsys.readouterr().out


def test_lenient_mode_is_the_default(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from docxpdf_native.cli import main
    from docxpdf_native.models import ConversionOptions

    source, destination = _paths(tmp_path)
    calls = _install_converter(monkeypatch)

    assert main([str(source), "-o", str(destination)]) == 0

    options = calls["options"]
    assert isinstance(options, ConversionOptions)
    assert options.strict is False


def test_strict_and_lenient_are_mutually_exclusive(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)
    calls = _install_converter(monkeypatch)

    exit_code = main([str(source), "-o", str(destination), "--strict", "--lenient"])

    assert exit_code == 2
    assert calls == {}
    assert "not allowed with argument" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("arguments", "message"),
    [
        (["--font-substitution", "missing-separator"], "NAME=TARGET"),
        (["--font-substitution", "=missing-name"], "NAME=TARGET"),
        (["--max-pages", "0"], "positive integer"),
    ],
)
def test_invalid_option_value_returns_input_error(
    arguments: list[str],
    message: str,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)

    exit_code = main([str(source), "-o", str(destination), *arguments])

    assert exit_code == 2
    assert message in capsys.readouterr().err


def test_missing_input_file_returns_input_error(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from docxpdf_native.cli import main

    exit_code = main([str(tmp_path / "missing.docx"), "-o", str(tmp_path / "output.pdf")])

    assert exit_code == 2
    assert "input file does not exist" in capsys.readouterr().err


def test_missing_font_directory_returns_input_error(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)
    exit_code = main(
        [
            str(source),
            "-o",
            str(destination),
            "--font-dir",
            str(tmp_path / "missing-fonts"),
        ]
    )

    assert exit_code == 2
    assert "font directory does not exist" in capsys.readouterr().err


def test_cli_writes_diagnostics_and_layout_json(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)
    diagnostics_path = tmp_path / "diagnostics.json"
    layout_path = tmp_path / "layout.json"
    diagnostics = DiagnosticInfo(counters={"pages": 2})
    _install_converter(
        monkeypatch,
        result=ConversionResult(page_count=2, diagnostics=diagnostics),
        layout=LayoutDocument(),
    )

    exit_code = main(
        [
            str(source),
            "-o",
            str(destination),
            "--diagnostics",
            str(diagnostics_path),
            "--layout-json",
            str(layout_path),
        ]
    )

    assert exit_code == 0
    assert json.loads(diagnostics_path.read_text(encoding="utf-8"))["counters"] == {"pages": 2}
    assert json.loads(layout_path.read_text(encoding="utf-8"))["pages"] == []


def test_requested_layout_json_without_layout_returns_unexpected_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)
    _install_converter(monkeypatch, layout=None)

    exit_code = main(
        [
            str(source),
            "-o",
            str(destination),
            "--layout-json",
            str(tmp_path / "layout.json"),
        ]
    )

    assert exit_code == 10
    assert "layout data is unavailable" in capsys.readouterr().err


def test_lenient_warnings_are_printed_to_stderr(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)
    warning = ConversionWarning(code="unsupported", message="text box was omitted")
    _install_converter(
        monkeypatch,
        result=ConversionResult(page_count=1, warnings=(warning,)),
    )

    assert main([str(source), "-o", str(destination), "--lenient"]) == 0

    assert "warning[unsupported]: text box was omitted" in capsys.readouterr().err


def test_converter_initialization_failure_returns_unexpected_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native import cli

    source, destination = _paths(tmp_path)

    def fail_to_create(options: object) -> object:
        del options
        raise RuntimeError("initialization failed")

    monkeypatch.setattr(cli, "_create_converter", fail_to_create)

    exit_code = cli.main([str(source), "-o", str(destination)])

    assert exit_code == 10
    assert "initialization failed" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("error", "expected_code"),
    [
        (
            UnsupportedFeatureError(UnsupportedFeature(name="text box", part="word/document.xml")),
            3,
        ),
        (FontNotFoundError.for_font("Missing Font"), 4),
        (InvalidDocxError("invalid ZIP"), 5),
        (InvalidOoxmlError("invalid XML"), 5),
        (MissingPartError("missing document"), 5),
        (RelationshipError("broken relationship"), 5),
        (ResourceLimitError("expanded size exceeded"), 5),
        (LayoutError("layout failed"), 6),
        (PdfGenerationError("PDF failed"), 6),
        (RuntimeError("unexpected"), 10),
    ],
)
def test_conversion_errors_map_to_documented_exit_codes(
    error: Exception,
    expected_code: int,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docxpdf_native.cli import main

    source, destination = _paths(tmp_path)
    _install_converter(monkeypatch, error=error)

    exit_code = main([str(source), "-o", str(destination)])

    assert exit_code == expected_code
    assert str(error) in capsys.readouterr().err
