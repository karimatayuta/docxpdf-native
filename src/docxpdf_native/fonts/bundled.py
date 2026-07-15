"""Metric-compatible fallback fonts bundled with the package.

Word's modern default Latin fonts -- Calibri (sans, default since Word 2007)
and Cambria (serif) -- are proprietary and rarely installed on conversion
machines. Substituting them with an arbitrary fallback such as Helvetica
changes advance widths, which changes line wrapping and therefore page
counts (observed: +31% pages on a real-world document).

To keep layout close to the original without a network connection, the
package bundles two SIL OFL 1.1 licensed families that are *metrically
compatible* (same advance widths) with those Word defaults:

- **Carlito** (compatible with Calibri) -- Google,
  https://github.com/googlefonts/carlito
- **Caladea** (compatible with Cambria) -- Huerta Tipografica,
  https://github.com/huertatipografica/Caladea

The font files and their OFL.txt license texts live in
``fonts/data/carlito/`` and ``fonts/data/caladea/`` next to this module.

Face ordering matters: every face of a family (Regular, Bold, Italic, Bold
Italic) carries the same family name in its ``name`` table, and the registry
indexes with first-wins semantics. :data:`_BUNDLED_FACE_FILES` therefore
lists Regular faces first so the bare family name always resolves to the
Regular face, while the styled faces remain reachable through their full
names ("Carlito Bold" and so on).
"""

from __future__ import annotations

from pathlib import Path

from docxpdf_native.fonts.cjk import normalize_font_name

_DATA_DIRECTORY = Path(__file__).parent / "data"

_BUNDLED_FACE_FILES = (
    "carlito/Carlito-Regular.ttf",
    "carlito/Carlito-Bold.ttf",
    "carlito/Carlito-Italic.ttf",
    "carlito/Carlito-BoldItalic.ttf",
    "caladea/Caladea-Regular.ttf",
    "caladea/Caladea-Bold.ttf",
    "caladea/Caladea-Italic.ttf",
    "caladea/Caladea-BoldItalic.ttf",
)


def bundled_font_paths() -> tuple[Path, ...]:
    """Return the bundled font files in deterministic indexing order.

    Missing files (e.g. a stripped-down installation) are silently skipped;
    resolution then simply falls through to the next priority group.
    """
    return tuple(
        path for relative in _BUNDLED_FACE_FILES if (path := _DATA_DIRECTORY / relative).is_file()
    )


# Requested-name -> bundled-face-name map. Keys are normalized with
# normalize_font_name(); values are names resolvable through the registry
# (family names and full face names of the bundled fonts). "Light" weights
# have no bundled counterpart and map to the nearest available face.
_METRIC_COMPATIBLE_SUBSTITUTIONS: dict[str, str] = {
    "calibri": "Carlito",
    "calibri light": "Carlito",
    "calibri bold": "Carlito Bold",
    "calibri italic": "Carlito Italic",
    "calibri light italic": "Carlito Italic",
    "calibri bold italic": "Carlito Bold Italic",
    "cambria": "Caladea",
    "cambria math": "Caladea",
    "cambria bold": "Caladea Bold",
    "cambria italic": "Caladea Italic",
    "cambria bold italic": "Caladea Bold Italic",
}


def metric_compatible_substitution(name: str) -> str | None:
    """Map a well-known proprietary Latin font name to a bundled equivalent.

    Returns the name of a metrically compatible bundled face for Calibri and
    Cambria (including common styled variants), or ``None`` for any other
    font. The caller is expected to resolve the returned name through the
    registry so that an actually installed font of that name still wins over
    the bundled copy.
    """
    return _METRIC_COMPATIBLE_SUBSTITUTIONS.get(normalize_font_name(name))
