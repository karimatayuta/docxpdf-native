from __future__ import annotations

import argparse
from pathlib import Path

from docxpdf_native import ConversionOptions, Converter
from docxpdf_native.diagnostics import DiagnosticsExporter


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export conversion diagnostics and layout JSON.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("diagnostics", type=Path)
    parser.add_argument("layout", type=Path)
    return parser.parse_args()


def main() -> None:
    arguments = parse_arguments()
    converter = Converter(options=ConversionOptions(strict=False, deterministic=True))
    result = converter.convert(arguments.source, arguments.output)
    DiagnosticsExporter.write(result.diagnostics, arguments.diagnostics)

    layout = converter.last_layout
    if layout is None:
        raise RuntimeError("conversion completed without a layout model")
    DiagnosticsExporter.write(layout, arguments.layout)

    print(f"pages={result.page_count}")
    print(f"warnings={len(result.warnings)}")


if __name__ == "__main__":
    main()
