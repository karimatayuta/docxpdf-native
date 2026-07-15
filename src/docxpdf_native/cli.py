from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Protocol

from docxpdf_native import __version__
from docxpdf_native.diagnostics import DiagnosticsExporter
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
    ConversionOptions,
    ConversionResult,
    FontConfiguration,
    LayoutDocument,
    ResourceLimits,
)
from docxpdf_native.models.base import FrozenModel
from docxpdf_native.security.paths import OutputPathGuard

EXIT_SUCCESS = 0
EXIT_INPUT_ERROR = 2
EXIT_UNSUPPORTED_FEATURE = 3
EXIT_FONT_NOT_FOUND = 4
EXIT_INVALID_DOCX = 5
EXIT_PDF_GENERATION = 6
EXIT_UNEXPECTED = 10


class _ConverterProtocol(Protocol):
    @property
    def last_layout(self) -> LayoutDocument | None: ...

    def convert(self, source: Path, destination: Path) -> ConversionResult: ...


class _CliArguments(FrozenModel):
    source: Path
    output: Path
    font_dirs: tuple[Path, ...] = ()
    font_substitutions: tuple[tuple[str, str], ...] = ()
    strict: bool = False
    deterministic: bool = True
    diagnostics: Path | None = None
    layout_json: Path | None = None
    max_pages: int


def _create_converter(options: ConversionOptions) -> _ConverterProtocol:
    from docxpdf_native.converter import Converter

    return Converter(options=options)


def _positive_integer(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("expected a positive integer") from error
    if parsed <= 0:
        raise argparse.ArgumentTypeError("expected a positive integer")
    return parsed


def _font_substitution(value: str) -> tuple[str, str]:
    if "=" not in value:
        raise argparse.ArgumentTypeError("font substitution must use NAME=TARGET")
    requested, selected = (item.strip() for item in value.split("=", maxsplit=1))
    if not requested or not selected:
        raise argparse.ArgumentTypeError("font substitution must use NAME=TARGET")
    return requested, selected


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docxpdf-native",
        description="Convert a supported DOCX subset directly to PDF.",
    )
    parser.add_argument("source", type=Path, help="input DOCX file")
    parser.add_argument("-o", "--output", type=Path, required=True, help="output PDF file")
    parser.add_argument(
        "--font-dir",
        dest="font_dirs",
        action="append",
        default=[],
        type=Path,
        metavar="DIRECTORY",
        help="font search directory; may be repeated",
    )
    parser.add_argument(
        "--font-substitution",
        dest="font_substitutions",
        action="append",
        default=[],
        type=_font_substitution,
        metavar="NAME=TARGET",
        help="explicit font substitution; may be repeated",
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--strict",
        dest="strict",
        action="store_true",
        help="fail the moment any content would not render with full fidelity",
    )
    mode.add_argument(
        "--lenient",
        dest="strict",
        action="store_false",
        help=(
            "convert everything possible, replacing unrenderable content with "
            "same-size placeholders and reporting warnings (default)"
        ),
    )
    parser.set_defaults(strict=False)
    parser.add_argument(
        "--deterministic",
        action="store_true",
        default=ConversionOptions().deterministic,
        help="request deterministic PDF metadata and drawing order",
    )
    parser.add_argument("--diagnostics", type=Path, help="write diagnostics JSON")
    parser.add_argument("--layout-json", type=Path, help="write layout model JSON")
    parser.add_argument(
        "--max-pages",
        type=_positive_integer,
        default=ResourceLimits().max_pages,
        metavar="COUNT",
        help="maximum number of generated pages",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def _parse_arguments(argv: Sequence[str] | None) -> _CliArguments | int:
    try:
        namespace = _build_parser().parse_args(argv)
    except SystemExit as error:
        return error.code if isinstance(error.code, int) else EXIT_INPUT_ERROR
    return _CliArguments.model_validate(vars(namespace))


def _validate_paths(arguments: _CliArguments) -> str | None:
    if not arguments.source.is_file():
        return f"input file does not exist: {arguments.source}"
    if not arguments.output.parent.is_dir():
        return f"output parent directory does not exist: {arguments.output.parent}"
    for directory in arguments.font_dirs:
        if not directory.is_dir():
            return f"font directory does not exist: {directory}"
    for destination in (arguments.diagnostics, arguments.layout_json):
        if destination is not None and not destination.parent.is_dir():
            return f"JSON output parent directory does not exist: {destination.parent}"
    destinations = tuple(
        destination
        for destination in (
            arguments.output,
            arguments.diagnostics,
            arguments.layout_json,
        )
        if destination is not None
    )
    keys = tuple(OutputPathGuard.comparison_key(destination) for destination in destinations)
    if len(set(keys)) != len(keys):
        return "PDF, diagnostics, and layout JSON must use different paths"
    return None


def _conversion_options(arguments: _CliArguments) -> ConversionOptions:
    return ConversionOptions(
        strict=arguments.strict,
        deterministic=arguments.deterministic,
        font_configuration=FontConfiguration(
            font_directories=arguments.font_dirs,
            substitutions=dict(arguments.font_substitutions),
        ),
        resource_limits=ResourceLimits(max_pages=arguments.max_pages),
    )


def _print_error(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)


def _convert(arguments: _CliArguments) -> int:
    try:
        converter = _create_converter(_conversion_options(arguments))
        result = converter.convert(arguments.source, arguments.output)
    except UnsupportedFeatureError as error:
        _print_error(str(error))
        return EXIT_UNSUPPORTED_FEATURE
    except FontNotFoundError as error:
        _print_error(str(error))
        return EXIT_FONT_NOT_FOUND
    except (
        InvalidDocxError,
        InvalidOoxmlError,
        MissingPartError,
        RelationshipError,
        ResourceLimitError,
    ) as error:
        _print_error(str(error))
        return EXIT_INVALID_DOCX
    except (LayoutError, PdfGenerationError) as error:
        _print_error(str(error))
        return EXIT_PDF_GENERATION
    except Exception as error:
        _print_error(str(error))
        return EXIT_UNEXPECTED

    try:
        if arguments.diagnostics is not None:
            DiagnosticsExporter.write(result.diagnostics, arguments.diagnostics)
        if arguments.layout_json is not None:
            if converter.last_layout is None:
                _print_error("layout data is unavailable")
                return EXIT_UNEXPECTED
            DiagnosticsExporter.write(converter.last_layout, arguments.layout_json)
    except (OSError, ValueError) as error:
        _print_error(str(error))
        return EXIT_INPUT_ERROR

    for warning in result.warnings:
        print(f"warning[{warning.code}]: {warning.message}", file=sys.stderr)
    page_label = "page" if result.page_count == 1 else "pages"
    print(f"Converted {arguments.source} -> {arguments.output} ({result.page_count} {page_label})")
    return EXIT_SUCCESS


def main(argv: Sequence[str] | None = None) -> int:
    parsed = _parse_arguments(argv)
    if isinstance(parsed, int):
        return parsed
    path_error = _validate_paths(parsed)
    if path_error is not None:
        _print_error(path_error)
        return EXIT_INPUT_ERROR
    return _convert(parsed)
