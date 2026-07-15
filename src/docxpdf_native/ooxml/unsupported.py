"""Central unsupported-feature detection and strict/lenient policy handlers."""

from __future__ import annotations

from collections.abc import Iterator
from typing import ClassVar
from xml.etree.ElementTree import Element

from docxpdf_native.abstractions import UnsupportedFeatureHandler
from docxpdf_native.exceptions import UnsupportedFeatureError
from docxpdf_native.models.results import ConversionWarning, UnsupportedFeature
from docxpdf_native.ooxml.namespaces import OoxmlNamespaces
from docxpdf_native.ooxml.package import OoxmlPackage


class StrictUnsupportedFeatureHandler(UnsupportedFeatureHandler):
    """Stop conversion at the first unsupported feature."""

    def handle(self, feature: UnsupportedFeature) -> ConversionWarning | None:
        """Raise an error containing the complete feature context."""

        raise UnsupportedFeatureError(feature)


class LenientUnsupportedFeatureHandler(UnsupportedFeatureHandler):
    """Collect deterministic warnings while allowing supported content through."""

    def __init__(self) -> None:
        self._features: list[UnsupportedFeature] = []
        self._warnings: list[ConversionWarning] = []

    @property
    def features(self) -> tuple[UnsupportedFeature, ...]:
        """Return unsupported features in document order."""

        return tuple(self._features)

    @property
    def warnings(self) -> tuple[ConversionWarning, ...]:
        """Return warnings in document order."""

        return tuple(self._warnings)

    def handle(self, feature: UnsupportedFeature) -> ConversionWarning:
        """Record a feature and return its public warning representation."""

        warning = ConversionWarning(
            code="unsupported_feature",
            message=f"Unsupported feature {feature.name!r} was not rendered",
            part=feature.part,
            location=feature.location,
            feature=feature,
        )
        self._features.append(feature)
        self._warnings.append(warning)
        return warning


class UnsupportedFeatureDetector:
    """Detect known unsupported OOXML constructs before model generation.

    Constructs the parser now renders natively (content controls, tracked
    changes, complex fields, VML/OLE images, anchored drawings) are
    intentionally absent from ``_RULES``: they are no longer "unsupported",
    they degrade gracefully (rendered content or a same-size placeholder)
    and are reported through the parser's own placeholder diagnostics
    instead of this pre-scan, so they never block strict mode.
    """

    _RULES: ClassVar[dict[str, tuple[str, str | None]]] = {
        OoxmlNamespaces.qn("w", "txbxContent"): (
            "text_box",
            "Move text out of the text box into ordinary paragraphs.",
        ),
        OoxmlNamespaces.qn("v", "textbox"): (
            "text_box",
            "Move text out of the text box into ordinary paragraphs.",
        ),
        OoxmlNamespaces.qn("wps", "txbx"): (
            "text_box",
            "Move text out of the text box into ordinary paragraphs.",
        ),
        OoxmlNamespaces.qn("dgm", "relIds"): (
            "smart_art",
            "Replace SmartArt with an inline PNG or JPEG image.",
        ),
        OoxmlNamespaces.qn("c", "chart"): (
            "chart",
            "Replace the chart with an inline PNG or JPEG image.",
        ),
        OoxmlNamespaces.qn("m", "oMath"): (
            "math",
            "Replace the equation with text or an inline image.",
        ),
        OoxmlNamespaces.qn("m", "oMathPara"): (
            "math",
            "Replace the equation with text or an inline image.",
        ),
        OoxmlNamespaces.qn("v", "textpath"): (
            "word_art",
            "Replace WordArt with ordinary text or an inline image.",
        ),
        OoxmlNamespaces.qn("w", "ruby"): (
            "ruby",
            "Use plain parenthetical readings instead of ruby markup.",
        ),
        OoxmlNamespaces.qn("w", "bidi"): (
            "bidirectional_layout",
            "Convert the paragraph to left-to-right horizontal layout.",
        ),
        OoxmlNamespaces.qn("w", "bidiVisual"): (
            "bidirectional_layout",
            "Convert the table to left-to-right layout.",
        ),
        OoxmlNamespaces.qn("w", "rtl"): (
            "bidirectional_layout",
            "Convert the run to left-to-right text.",
        ),
        OoxmlNamespaces.qn("w", "cs"): (
            "complex_arabic_shaping",
            "Convert the complex-script run to pre-shaped text or an image.",
        ),
        OoxmlNamespaces.qn("w", "commentReference"): (
            "comments",
            "Remove comments or render a clean document copy.",
        ),
        OoxmlNamespaces.qn("w", "commentRangeStart"): (
            "comments",
            "Remove comments or render a clean document copy.",
        ),
        OoxmlNamespaces.qn("w", "commentRangeEnd"): (
            "comments",
            "Remove comments or render a clean document copy.",
        ),
        OoxmlNamespaces.qn("w", "footnoteReference"): (
            "footnotes",
            "Move footnote text into the document body.",
        ),
        OoxmlNamespaces.qn("w", "endnoteReference"): (
            "endnotes",
            "Move endnote text into the document body.",
        ),
    }

    def __init__(self, handler: UnsupportedFeatureHandler) -> None:
        self._handler = handler

    def scan_element(self, root: Element, *, part_name: str) -> None:
        """Walk an XML tree and invoke the configured policy for each match."""

        for element, location in self._walk(root):
            spec = self._feature_for(element)
            if spec is None:
                continue
            name, workaround = spec
            self._handler.handle(
                UnsupportedFeature(
                    name=name,
                    part=part_name,
                    element=self._display_name(element.tag),
                    location=location,
                    workaround=workaround,
                )
            )

    def scan_package(self, package: OoxmlPackage) -> None:
        """Detect unsupported binary/auxiliary parts.

        Embedded-object binaries and non-PNG/JPEG image relationships are no
        longer flagged here: the parser now resolves them into rendered
        images or same-size placeholders, with degradation reported through
        its own placeholder diagnostics regardless of strict/lenient mode.
        """

        exact_parts = {
            "word/vbaProject.bin": (
                "macro",
                "Save a macro-free DOCX copy before conversion.",
            ),
        }
        prefix_parts = {
            "word/charts/": (
                "chart",
                "Replace the chart with an inline PNG or JPEG image.",
            ),
            "word/diagrams/": (
                "smart_art",
                "Replace SmartArt with an inline PNG or JPEG image.",
            ),
        }
        for part_name in package.part_names:
            spec = exact_parts.get(part_name)
            if spec is None:
                spec = next(
                    (
                        candidate
                        for prefix, candidate in prefix_parts.items()
                        if part_name.startswith(prefix)
                    ),
                    None,
                )
            if spec is None:
                continue
            name, workaround = spec
            self._handler.handle(
                UnsupportedFeature(
                    name=name,
                    part=part_name,
                    location="/",
                    workaround=workaround,
                )
            )

    @classmethod
    def _feature_for(cls, element: Element) -> tuple[str, str | None] | None:
        if element.tag == OoxmlNamespaces.qn("w", "textDirection"):
            direction = element.get(OoxmlNamespaces.qn("w", "val"), "lrTb")
            if direction in {"lrTb", "lr", "tb"}:
                return None
            return (
                "vertical_writing",
                "Convert the content to horizontal left-to-right text.",
            )
        if element.tag == OoxmlNamespaces.qn("w", "cols"):
            raw_count = element.get(OoxmlNamespaces.qn("w", "num"), "1")
            try:
                count = int(raw_count)
            except ValueError:
                count = 2
            if count <= 1:
                return None
            return (
                "multiple_columns",
                "Change the section to a single text column.",
            )
        return cls._RULES.get(element.tag)

    @classmethod
    def _walk(cls, root: Element) -> Iterator[tuple[Element, str]]:
        root_name = cls._local_name(root.tag)
        yield root, f"/{root_name}[1]"
        yield from cls._walk_children(root, parent_path=f"/{root_name}[1]")

    @classmethod
    def _walk_children(
        cls,
        parent: Element,
        *,
        parent_path: str,
    ) -> Iterator[tuple[Element, str]]:
        counts: dict[str, int] = {}
        for child in parent:
            name = cls._local_name(child.tag)
            counts[name] = counts.get(name, 0) + 1
            path = f"{parent_path}/{name}[{counts[name]}]"
            yield child, path
            yield from cls._walk_children(child, parent_path=path)

    @staticmethod
    def _local_name(tag: str) -> str:
        return tag.rsplit("}", maxsplit=1)[-1]

    @staticmethod
    def _display_name(tag: str) -> str:
        namespace, separator, local_name = tag.removeprefix("{").partition("}")
        if not separator:
            return tag
        for prefix, uri in OoxmlNamespaces.PREFIXES.items():
            if uri == namespace:
                return f"{prefix}:{local_name}"
        return tag
