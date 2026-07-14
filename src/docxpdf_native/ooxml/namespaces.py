"""Central OOXML namespace and lexical value helpers."""

from __future__ import annotations

import re
from types import MappingProxyType
from typing import ClassVar


class OoxmlNamespaces:
    """OOXML namespace registry.

    Namespace URIs belong here so parsers never duplicate string literals.  All
    helpers are stateless and deliberately exposed as static methods.
    """

    WORDPROCESSINGML: ClassVar[str] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    WORDPROCESSINGML_STRICT: ClassVar[str] = "http://purl.oclc.org/ooxml/wordprocessingml/main"
    DRAWINGML: ClassVar[str] = "http://schemas.openxmlformats.org/drawingml/2006/main"
    DRAWINGML_PICTURE: ClassVar[str] = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    WORDPROCESSING_DRAWING: ClassVar[str] = (
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    )
    RELATIONSHIPS: ClassVar[str] = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    PACKAGE_RELATIONSHIPS: ClassVar[str] = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    CONTENT_TYPES: ClassVar[str] = "http://schemas.openxmlformats.org/package/2006/content-types"
    MARKUP_COMPATIBILITY: ClassVar[str] = (
        "http://schemas.openxmlformats.org/markup-compatibility/2006"
    )
    MATH: ClassVar[str] = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    VML: ClassVar[str] = "urn:schemas-microsoft-com:vml"
    OFFICE: ClassVar[str] = "urn:schemas-microsoft-com:office:office"
    CHART: ClassVar[str] = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    DIAGRAM: ClassVar[str] = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
    WORDPROCESSING_SHAPE: ClassVar[str] = (
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    )
    WORDPROCESSING_GROUP: ClassVar[str] = (
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    )
    XML: ClassVar[str] = "http://www.w3.org/XML/1998/namespace"
    CORE_PROPERTIES: ClassVar[str] = (
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    )
    EXTENDED_PROPERTIES: ClassVar[str] = (
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    )

    PREFIXES: ClassVar[MappingProxyType[str, str]] = MappingProxyType(
        {
            "w": WORDPROCESSINGML,
            "w-strict": WORDPROCESSINGML_STRICT,
            "a": DRAWINGML,
            "pic": DRAWINGML_PICTURE,
            "wp": WORDPROCESSING_DRAWING,
            "r": RELATIONSHIPS,
            "pr": PACKAGE_RELATIONSHIPS,
            "ct": CONTENT_TYPES,
            "mc": MARKUP_COMPATIBILITY,
            "m": MATH,
            "v": VML,
            "o": OFFICE,
            "c": CHART,
            "dgm": DIAGRAM,
            "wps": WORDPROCESSING_SHAPE,
            "wpg": WORDPROCESSING_GROUP,
            "xml": XML,
            "cp": CORE_PROPERTIES,
            "ep": EXTENDED_PROPERTIES,
        }
    )

    _LOCAL_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]*$")
    _RGB = re.compile(r"^[0-9A-Fa-f]{6}$")
    _HIGHLIGHT_COLORS: ClassVar[MappingProxyType[str, str]] = MappingProxyType(
        {
            "black": "000000",
            "blue": "0000FF",
            "cyan": "00FFFF",
            "green": "00FF00",
            "magenta": "FF00FF",
            "red": "FF0000",
            "yellow": "FFFF00",
            "white": "FFFFFF",
            "darkblue": "000080",
            "darkcyan": "008080",
            "darkgreen": "008000",
            "darkmagenta": "800080",
            "darkred": "800000",
            "darkyellow": "808000",
            "darkgray": "808080",
            "lightgray": "C0C0C0",
        }
    )

    @staticmethod
    def qn(prefix: str, local_name: str) -> str:
        """Return an ElementTree Clark name for a registered prefix."""

        namespace = OoxmlNamespaces.PREFIXES.get(prefix)
        if namespace is None:
            raise ValueError(f"Unknown OOXML namespace prefix: {prefix!r}")
        if OoxmlNamespaces._LOCAL_NAME.fullmatch(local_name) is None:
            raise ValueError(f"Invalid XML local name: {local_name!r}")
        return f"{{{namespace}}}{local_name}"

    @staticmethod
    def parse_on_off(value: str | None, *, default: bool = True) -> bool:
        """Parse the lexical values used by the OOXML ``ST_OnOff`` type."""

        if value is None:
            return default
        normalized = value.strip().lower()
        if normalized in {"1", "true", "on"}:
            return True
        if normalized in {"0", "false", "off"}:
            return False
        raise ValueError(f"Invalid OOXML boolean value: {value!r}")

    @staticmethod
    def normalize_color(value: str | None) -> str | None:
        """Normalize a six-digit sRGB value, treating ``auto`` as unspecified."""

        if value is None or value.strip().lower() == "auto":
            return None
        normalized = value.strip().removeprefix("#")
        if OoxmlNamespaces._RGB.fullmatch(normalized) is None:
            raise ValueError(f"Invalid OOXML color value: {value!r}")
        return normalized.upper()

    @staticmethod
    def normalize_highlight(value: str | None) -> str | None:
        """Convert an OOXML named highlight color to six-digit sRGB."""

        if value is None or value.strip().casefold() == "none":
            return None
        normalized = value.strip().casefold()
        color = OoxmlNamespaces._HIGHLIGHT_COLORS.get(normalized)
        if color is None:
            raise ValueError(f"Invalid OOXML highlight color: {value!r}")
        return color
