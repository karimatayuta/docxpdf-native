from __future__ import annotations

import argparse
from pathlib import Path

from docxpdf_native import ConversionOptions, Converter


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert one DOCX file to PDF.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    return parser.parse_args()


def main() -> None:
    arguments = parse_arguments()
    converter = Converter(options=ConversionOptions(strict=True, deterministic=True))
    result = converter.convert(arguments.source, arguments.output)
    print(f"pages={result.page_count}")
    print(f"pdf_sha256={result.pdf_sha256}")


if __name__ == "__main__":
    main()
