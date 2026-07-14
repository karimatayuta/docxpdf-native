from __future__ import annotations

import argparse
from pathlib import Path

from docxpdf_native import ConversionOptions, Converter, FontConfiguration


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert DOCX with an explicit Japanese font.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument(
        "font",
        type=Path,
        help="Japanese TrueType font (.ttf, or .otf with glyf outlines)",
    )
    return parser.parse_args()


def main() -> None:
    arguments = parse_arguments()
    fonts = FontConfiguration(
        registered_fonts={"Japanese Body": arguments.font},
        substitutions={
            "MS Gothic": "Japanese Body",
            "MS Mincho": "Japanese Body",
        },
        default_font="Japanese Body",
        include_system_fonts=False,
    )
    options = ConversionOptions(
        strict=True,
        deterministic=True,
        font_configuration=fonts,
    )
    result = Converter(options=options).convert(arguments.source, arguments.output)
    print(f"pages={result.page_count}")
    for substitution in result.font_substitutions:
        print(
            f"font: {substitution.requested_font} -> "
            f"{substitution.selected_font} ({substitution.reason})"
        )


if __name__ == "__main__":
    main()
