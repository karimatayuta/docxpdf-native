"""Deterministic Japanese/CJK font knowledge shared by the font resolver.

Two independent problems are solved here:

1. Documents commonly request well-known Japanese font names ("MS 明朝",
   "MS Gothic", "游ゴシック", "Meiryo", ...) that are Windows/Office system
   fonts and are almost never present on the machine performing the
   conversion. :data:`_KNOWN_JAPANESE_FONTS` maps these (including common
   full-width/half-width spelling variants) to a generic ``serif``/``sans``
   category.
2. Once a request is known to want a serif or sans CJK face, something
   concrete has to back it. :func:`detect_system_cjk_fonts` deterministically
   probes a fixed, ordered list of real font family names that are likely to
   already be installed for the current platform (no network access, no
   bundled fonts).
"""

from __future__ import annotations

import platform
import unicodedata
from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from docxpdf_native.fonts.registry import FontRecord, FontRegistry

FontCategory = Literal["serif", "sans"]

# Ordered, deterministic candidates per platform and category. The first
# candidate that the registry can actually resolve wins.
_SYSTEM_CJK_CANDIDATES: dict[str, dict[FontCategory, tuple[str, ...]]] = {
    "Darwin": {
        "serif": (
            "Hiragino Mincho ProN",
            "Hiragino Mincho Pro",
            "ヒラギノ明朝 ProN",
            "Arial Unicode MS",
        ),
        "sans": (
            "Hiragino Sans",
            "Hiragino Kaku Gothic ProN",
            "Hiragino Kaku Gothic Pro",
            "ヒラギノ角ゴシック",
            "Arial Unicode MS",
        ),
    },
    "Windows": {
        "serif": ("Yu Mincho", "MS Mincho", "MS PMincho", "游明朝"),
        "sans": (
            "Yu Gothic",
            "Meiryo",
            "MS Gothic",
            "MS PGothic",
            "MS UI Gothic",
            "游ゴシック",
        ),
    },
    "Linux": {
        "serif": (
            "Noto Serif CJK JP",
            "IPAex明朝",
            "IPAExMincho",
            "IPA明朝",
            "IPAMincho",
        ),
        "sans": (
            "Noto Sans CJK JP",
            "IPAexゴシック",
            "IPAExGothic",
            "IPAゴシック",
            "IPAGothic",
        ),
    },
}
_DEFAULT_CANDIDATES = _SYSTEM_CJK_CANDIDATES["Linux"]

_KNOWN_JAPANESE_FONTS: dict[str, FontCategory] = {}


def normalize_font_name(name: str) -> str:
    """Collapse whitespace/casing and fold full-width forms to half-width.

    NFKC turns full-width Latin letters and the full-width ideographic space
    (U+3000) into their ordinary ASCII equivalents, so a font name spelled
    with full-width "MS" and/or a full-width space between "MS" and the
    Japanese suffix normalizes identically to the plain half-width spelling.
    Shared with the built-in Latin substitution map in
    :mod:`docxpdf_native.fonts.bundled`.
    """
    return " ".join(unicodedata.normalize("NFKC", name).split()).casefold()


def _register(names: tuple[str, ...], category: FontCategory) -> None:
    for name in names:
        _KNOWN_JAPANESE_FONTS[normalize_font_name(name)] = category


_register(
    (
        "MS 明朝",
        "MS Mincho",
        "ＭＳ 明朝",  # noqa: RUF001 - full-width "MS" spelling seen in real DOCX files
        "MS P明朝",
        "MS PMincho",
        "ＭＳ Ｐ明朝",  # noqa: RUF001 - full-width "MS P" spelling seen in real DOCX files
        "游明朝",
        "Yu Mincho",
        "YuMincho",
    ),
    "serif",
)
_register(
    (
        "MS ゴシック",
        "MS Gothic",
        "ＭＳ ゴシック",  # noqa: RUF001 - full-width "MS" spelling seen in real DOCX files
        "MS Pゴシック",
        "MS PGothic",
        "ＭＳ Ｐゴシック",  # noqa: RUF001 - full-width "MS P" spelling seen in real DOCX files
        "游ゴシック",
        "Yu Gothic",
        "YuGothic",
        "メイリオ",
        "Meiryo",
        "MS UI Gothic",
        "MS UIGothic",
    ),
    "sans",
)


def known_japanese_font_category(name: str) -> FontCategory | None:
    """Return the generic category for a well-known Japanese font name.

    Recognizes common full-width/half-width spelling variants of "MS
    Mincho"/"MS Gothic", "Yu Mincho"/"Yu Gothic", and "Meiryo". Returns
    ``None`` for anything else, including fonts that merely happen to be
    Japanese but are not one of these ubiquitous, almost-never-installed
    Office/Windows defaults.
    """
    return _KNOWN_JAPANESE_FONTS.get(normalize_font_name(name))


_CJK_CODEPOINT_RANGES = (
    (0x3040, 0x30FF),  # Hiragana + Katakana
    (0x31F0, 0x31FF),  # Katakana phonetic extensions
    (0x3400, 0x4DBF),  # CJK unified ideographs extension A
    (0x4E00, 0x9FFF),  # CJK unified ideographs
    (0xF900, 0xFAFF),  # CJK compatibility ideographs
    (0xFF00, 0xFFEF),  # Halfwidth and fullwidth forms
)


def contains_cjk_characters(text: str) -> bool:
    """Return whether ``text`` contains a Japanese/CJK script character."""
    for character in text:
        codepoint = ord(character)
        if any(low <= codepoint <= high for low, high in _CJK_CODEPOINT_RANGES):
            return True
    return False


def system_cjk_candidates() -> dict[FontCategory, tuple[str, ...]]:
    """Return the ordered serif/sans candidate names for the current OS."""
    return _SYSTEM_CJK_CANDIDATES.get(platform.system(), _DEFAULT_CANDIDATES)


def detect_system_cjk_fonts(registry: FontRegistry) -> dict[FontCategory, FontRecord]:
    """Deterministically resolve the best available serif/sans CJK font.

    Every candidate name for the current platform is looked up through the
    already-configured registry (so explicit ``registered_fonts``,
    ``font_directories``, and the environment variable are honored ahead of
    scanned system directories); the first candidate that resolves wins.
    """
    detected: dict[FontCategory, FontRecord] = {}
    for category, names in system_cjk_candidates().items():
        for name in names:
            record = registry.resolve(name)
            if record is not None:
                detected[category] = record
                break
    return detected
