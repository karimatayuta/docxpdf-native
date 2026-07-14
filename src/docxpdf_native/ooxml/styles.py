from __future__ import annotations

from collections.abc import Iterable
from typing import Literal, cast
from xml.etree.ElementTree import Element

from pydantic import Field

from docxpdf_native.layout.units import LengthConverter, OoxmlValueParser
from docxpdf_native.models.base import FrozenModel
from docxpdf_native.models.styles import (
    DocumentDefaults,
    ParagraphAlignment,
    ParagraphProperties,
    RunProperties,
    StyleModel,
    StyleType,
    TabStop,
    ThemeFonts,
)
from docxpdf_native.ooxml.namespaces import OoxmlNamespaces


class ParsedStyleSheet(FrozenModel):
    styles: tuple[StyleModel, ...] = ()
    defaults: DocumentDefaults = Field(default_factory=DocumentDefaults)


class OoxmlStylesParser:
    """Parse the supported subset of ``word/styles.xml`` from a safe XML root."""

    @staticmethod
    def parse(root: Element) -> ParsedStyleSheet:
        defaults_element = root.find(OoxmlNamespaces.qn("w", "docDefaults"))
        defaults = DocumentDefaults()
        if defaults_element is not None:
            run_default = defaults_element.find(
                f"{OoxmlNamespaces.qn('w', 'rPrDefault')}/{OoxmlNamespaces.qn('w', 'rPr')}"
            )
            paragraph_default = defaults_element.find(
                f"{OoxmlNamespaces.qn('w', 'pPrDefault')}/{OoxmlNamespaces.qn('w', 'pPr')}"
            )
            defaults = DocumentDefaults(
                run=OoxmlStylesParser._run_properties(run_default),
                paragraph=OoxmlStylesParser._paragraph_properties(paragraph_default),
            )

        styles = tuple(
            OoxmlStylesParser._style(element)
            for element in root.findall(OoxmlNamespaces.qn("w", "style"))
        )
        return ParsedStyleSheet(styles=styles, defaults=defaults)

    @staticmethod
    def _style(element: Element) -> StyleModel:
        style_id = OoxmlStylesParser._required_attribute(element, "styleId")
        style_type_value = OoxmlStylesParser._required_attribute(element, "type")
        allowed_types = {"paragraph", "character", "table", "numbering"}
        if style_type_value not in allowed_types:
            raise ValueError(f"unsupported style type: {style_type_value}")
        return StyleModel(
            style_id=style_id,
            style_type=cast(StyleType, style_type_value),
            name=OoxmlStylesParser._child_value(element, "name"),
            based_on=OoxmlStylesParser._child_value(element, "basedOn"),
            next_style=OoxmlStylesParser._child_value(element, "next"),
            link=OoxmlStylesParser._child_value(element, "link"),
            is_default=OoxmlValueParser.boolean(
                element.get(OoxmlNamespaces.qn("w", "default")), default=False
            ),
            paragraph=OoxmlStylesParser._paragraph_properties(
                element.find(OoxmlNamespaces.qn("w", "pPr"))
            ),
            run=OoxmlStylesParser._run_properties(element.find(OoxmlNamespaces.qn("w", "rPr"))),
        )

    @staticmethod
    def _run_properties(element: Element | None) -> RunProperties:
        if element is None:
            return RunProperties()
        fonts = element.find(OoxmlNamespaces.qn("w", "rFonts"))
        ascii_font = OoxmlStylesParser._font_attribute(fonts, "ascii", "asciiTheme")
        high_ansi_font = OoxmlStylesParser._font_attribute(fonts, "hAnsi", "hAnsiTheme")
        east_asia_font = OoxmlStylesParser._font_attribute(fonts, "eastAsia", "eastAsiaTheme")
        complex_script_font = OoxmlStylesParser._font_attribute(fonts, "cs", "csTheme")
        size = OoxmlStylesParser._child_value(element, "sz")
        spacing = OoxmlStylesParser._child_value(element, "spacing")
        color = OoxmlStylesParser._child_value(element, "color")
        underline_element = element.find(OoxmlNamespaces.qn("w", "u"))
        underline = OoxmlStylesParser._attribute(underline_element, "val")
        if underline_element is not None and underline is None:
            underline = "single"
        values: dict[str, object] = {}
        OoxmlStylesParser._put(values, "font_family", ascii_font or high_ansi_font)
        OoxmlStylesParser._put(values, "ascii_font", ascii_font)
        OoxmlStylesParser._put(values, "high_ansi_font", high_ansi_font)
        OoxmlStylesParser._put(values, "east_asia_font", east_asia_font)
        OoxmlStylesParser._put(values, "complex_script_font", complex_script_font)
        OoxmlStylesParser._put(
            values,
            "font_size",
            LengthConverter.half_point_to_point(int(size)) if size else None,
        )
        OoxmlStylesParser._put(values, "bold", OoxmlStylesParser._on_off(element, "b"))
        OoxmlStylesParser._put(values, "italic", OoxmlStylesParser._on_off(element, "i"))
        OoxmlStylesParser._put(
            values,
            "underline",
            False if underline == "none" else underline,
        )
        OoxmlStylesParser._put(values, "strike", OoxmlStylesParser._on_off(element, "strike"))
        if element.find(OoxmlNamespaces.qn("w", "color")) is not None:
            values["color"] = OoxmlValueParser.color(color)
        OoxmlStylesParser._put(
            values,
            "highlight",
            OoxmlNamespaces.normalize_highlight(
                OoxmlStylesParser._child_value(element, "highlight")
            ),
        )
        OoxmlStylesParser._put(values, "vertical_align", OoxmlStylesParser._vertical_align(element))
        OoxmlStylesParser._put(
            values,
            "character_spacing",
            LengthConverter.twip_to_point(int(spacing)) if spacing else None,
        )
        OoxmlStylesParser._put(values, "hidden", OoxmlStylesParser._on_off(element, "vanish"))
        return RunProperties.model_validate(values)

    @staticmethod
    def _paragraph_properties(element: Element | None) -> ParagraphProperties:
        if element is None:
            return ParagraphProperties()
        indentation = element.find(OoxmlNamespaces.qn("w", "ind"))
        spacing = element.find(OoxmlNamespaces.qn("w", "spacing"))
        line_value = OoxmlStylesParser._attribute(spacing, "line")
        line_rule = OoxmlStylesParser._attribute(spacing, "lineRule")
        line_spacing: float | None = None
        normalized_rule: Literal["auto", "at_least", "exact"] | None = None
        if line_value is not None:
            if line_rule in {None, "auto"}:
                line_spacing = int(line_value) / 240.0
                normalized_rule = "auto"
            else:
                line_spacing = LengthConverter.twip_to_point(int(line_value))
                normalized_rule = "at_least" if line_rule == "atLeast" else "exact"
        values: dict[str, object] = {}
        OoxmlStylesParser._put(values, "alignment", OoxmlStylesParser._alignment(element))
        OoxmlStylesParser._put(
            values,
            "left_indent",
            OoxmlStylesParser._twip_attribute(indentation, "left", "start"),
        )
        OoxmlStylesParser._put(
            values,
            "right_indent",
            OoxmlStylesParser._twip_attribute(indentation, "right", "end"),
        )
        OoxmlStylesParser._put(
            values,
            "first_line_indent",
            OoxmlStylesParser._twip_attribute(indentation, "firstLine"),
        )
        OoxmlStylesParser._put(
            values,
            "hanging_indent",
            OoxmlStylesParser._twip_attribute(indentation, "hanging"),
        )
        OoxmlStylesParser._put(
            values, "space_before", OoxmlStylesParser._twip_attribute(spacing, "before")
        )
        OoxmlStylesParser._put(
            values, "space_after", OoxmlStylesParser._twip_attribute(spacing, "after")
        )
        OoxmlStylesParser._put(values, "line_spacing", line_spacing)
        OoxmlStylesParser._put(values, "line_spacing_rule", normalized_rule)
        for field_name, element_name in (
            ("keep_next", "keepNext"),
            ("keep_lines", "keepLines"),
            ("page_break_before", "pageBreakBefore"),
            ("widow_control", "widowControl"),
        ):
            OoxmlStylesParser._put(
                values,
                field_name,
                OoxmlStylesParser._on_off(element, element_name),
            )
        if element.find(OoxmlNamespaces.qn("w", "tabs")) is not None:
            values["tabs"] = OoxmlStylesParser._tabs(element)
        OoxmlStylesParser._put(
            values,
            "numbering_id",
            OoxmlStylesParser._numbering_value(element, "numId"),
        )
        OoxmlStylesParser._put(
            values,
            "numbering_level",
            OoxmlStylesParser._numbering_value(element, "ilvl"),
        )
        return ParagraphProperties.model_validate(values)

    @staticmethod
    def _tabs(element: Element) -> tuple[TabStop, ...]:
        tabs = element.find(OoxmlNamespaces.qn("w", "tabs"))
        if tabs is None:
            return ()
        result: list[TabStop] = []
        for tab in tabs.findall(OoxmlNamespaces.qn("w", "tab")):
            position = OoxmlStylesParser._attribute(tab, "pos")
            if position is None:
                continue
            alignment = OoxmlStylesParser._attribute(tab, "val") or "left"
            leader = (OoxmlStylesParser._attribute(tab, "leader") or "none").replace(
                "middleDot", "middle_dot"
            )
            result.append(
                TabStop.model_validate(
                    {
                        "position": LengthConverter.twip_to_point(int(position)),
                        "alignment": alignment,
                        "leader": leader,
                    }
                )
            )
        return tuple(result)

    @staticmethod
    def _numbering_value(element: Element, name: str) -> int | None:
        numbering = element.find(OoxmlNamespaces.qn("w", "numPr"))
        value = OoxmlStylesParser._child_value(numbering, name)
        return int(value) if value is not None else None

    @staticmethod
    def _alignment(element: Element) -> ParagraphAlignment | None:
        value = OoxmlStylesParser._child_value(element, "jc")
        if value == "both":
            value = "justify"
        if value not in {"left", "center", "right", "justify", "distribute"}:
            return None
        return cast(ParagraphAlignment, value)

    @staticmethod
    def _vertical_align(
        element: Element,
    ) -> Literal["baseline", "superscript", "subscript"] | None:
        value = OoxmlStylesParser._child_value(element, "vertAlign")
        if value not in {"baseline", "superscript", "subscript"}:
            return None
        return cast(Literal["baseline", "superscript", "subscript"], value)

    @staticmethod
    def _on_off(parent: Element, name: str) -> bool | None:
        element = parent.find(OoxmlNamespaces.qn("w", name))
        if element is None:
            return None
        return OoxmlValueParser.boolean(element.get(OoxmlNamespaces.qn("w", "val")), default=True)

    @staticmethod
    def _font_attribute(element: Element | None, direct: str, theme: str) -> str | None:
        direct_value = OoxmlStylesParser._attribute(element, direct)
        if direct_value is not None:
            return direct_value
        theme_value = OoxmlStylesParser._attribute(element, theme)
        return f"+{theme_value}" if theme_value is not None else None

    @staticmethod
    def _twip_attribute(element: Element | None, *names: str) -> float | None:
        for name in names:
            value = OoxmlStylesParser._attribute(element, name)
            if value is not None:
                return LengthConverter.twip_to_point(int(value))
        return None

    @staticmethod
    def _child_value(parent: Element | None, name: str) -> str | None:
        if parent is None:
            return None
        return OoxmlStylesParser._attribute(parent.find(OoxmlNamespaces.qn("w", name)), "val")

    @staticmethod
    def _attribute(element: Element | None, name: str) -> str | None:
        if element is None:
            return None
        return element.get(OoxmlNamespaces.qn("w", name))

    @staticmethod
    def _required_attribute(element: Element, name: str) -> str:
        value = OoxmlStylesParser._attribute(element, name)
        if value is None:
            raise ValueError(f"style is missing required w:{name} attribute")
        return value

    @staticmethod
    def _put(values: dict[str, object], name: str, value: object | None) -> None:
        if value is not None:
            values[name] = value


class ThemeFontParser:
    """Parse theme font faces from ``word/theme/theme1.xml``."""

    @staticmethod
    def parse(root: Element) -> ThemeFonts:
        scheme = root.find(f".//{OoxmlNamespaces.qn('a', 'fontScheme')}")
        if scheme is None:
            return ThemeFonts()
        major = scheme.find(OoxmlNamespaces.qn("a", "majorFont"))
        minor = scheme.find(OoxmlNamespaces.qn("a", "minorFont"))
        return ThemeFonts(
            major_latin=ThemeFontParser._typeface(major, "latin"),
            major_east_asia=ThemeFontParser._typeface(major, "ea"),
            major_complex_script=ThemeFontParser._typeface(major, "cs"),
            minor_latin=ThemeFontParser._typeface(minor, "latin"),
            minor_east_asia=ThemeFontParser._typeface(minor, "ea"),
            minor_complex_script=ThemeFontParser._typeface(minor, "cs"),
        )

    @staticmethod
    def _typeface(parent: Element | None, name: str) -> str | None:
        if parent is None:
            return None
        element = parent.find(OoxmlNamespaces.qn("a", name))
        if element is None:
            return None
        return element.get("typeface") or None


class ResolvedStyleProperties(FrozenModel):
    paragraph: ParagraphProperties = Field(default_factory=ParagraphProperties)
    run: RunProperties = Field(default_factory=RunProperties)


class ThemeFontResolver:
    """Resolve OOXML ``+major*`` and ``+minor*`` theme placeholders."""

    @staticmethod
    def resolve(value: str | None, theme: ThemeFonts, *, east_asia: bool) -> str | None:
        if value is None or not value.startswith("+"):
            return value
        key = value.casefold()
        major = key.startswith("+major")
        minor = key.startswith("+minor")
        if not major and not minor:
            return value

        if "eastasia" in key or east_asia:
            selected = theme.major_east_asia if major else theme.minor_east_asia
        elif "bidi" in key or key.endswith("cs"):
            selected = theme.major_complex_script if major else theme.minor_complex_script
        else:
            selected = theme.major_latin if major else theme.minor_latin
        return selected or value


class OoxmlStyleResolver:
    """Resolve OOXML style inheritance and direct-formatting precedence."""

    def __init__(
        self,
        *,
        styles: Iterable[StyleModel],
        document_run_defaults: RunProperties | None = None,
        document_paragraph_defaults: ParagraphProperties | None = None,
        theme_fonts: ThemeFonts | None = None,
    ) -> None:
        self._styles: dict[str, StyleModel] = {}
        for style in styles:
            if style.style_id in self._styles:
                raise ValueError(f"duplicate style id: {style.style_id}")
            self._styles[style.style_id] = style
        self._document_run_defaults = document_run_defaults or RunProperties()
        self._document_paragraph_defaults = document_paragraph_defaults or ParagraphProperties()
        self._theme_fonts = theme_fonts or ThemeFonts()
        self._chain_cache: dict[str, tuple[StyleModel, ...]] = {}

    def resolve(
        self,
        *,
        paragraph_style_id: str | None = None,
        character_style_id: str | None = None,
        paragraph_direct: ParagraphProperties | None = None,
        run_direct: RunProperties | None = None,
    ) -> ResolvedStyleProperties:
        paragraph_layers: list[ParagraphProperties] = [self._document_paragraph_defaults]
        run_layers: list[RunProperties] = [self._document_run_defaults]

        if paragraph_style_id is not None:
            for style in self._style_chain(paragraph_style_id):
                paragraph_layers.append(style.paragraph)
                run_layers.append(style.run)
        if character_style_id is not None:
            for style in self._style_chain(character_style_id):
                run_layers.append(style.run)
        if paragraph_direct is not None:
            paragraph_layers.append(paragraph_direct)
        if run_direct is not None:
            run_layers.append(run_direct)

        paragraph = self._merge_paragraphs(paragraph_layers)
        run = self._resolve_theme_fonts(self._merge_runs(run_layers))
        return ResolvedStyleProperties(paragraph=paragraph, run=run)

    def _style_chain(self, style_id: str) -> tuple[StyleModel, ...]:
        cached = self._chain_cache.get(style_id)
        if cached is not None:
            return cached
        chain = self._build_chain(style_id, path=())
        self._chain_cache[style_id] = chain
        return chain

    def _build_chain(self, style_id: str, *, path: tuple[str, ...]) -> tuple[StyleModel, ...]:
        if style_id in path:
            cycle_start = path.index(style_id)
            cycle = (*path[cycle_start:], style_id)
            raise ValueError(f"style inheritance cycle: {' -> '.join(cycle)}")
        style = self._styles.get(style_id)
        if style is None:
            raise ValueError(f"unknown style id: {style_id}")
        if style.based_on is None:
            return (style,)
        return (*self._build_chain(style.based_on, path=(*path, style_id)), style)

    @staticmethod
    def _merge_runs(layers: Iterable[RunProperties]) -> RunProperties:
        values: dict[str, object] = {}
        for layer in layers:
            values.update(layer.model_dump(exclude_unset=True))
        return RunProperties.model_validate(values)

    @staticmethod
    def _merge_paragraphs(layers: Iterable[ParagraphProperties]) -> ParagraphProperties:
        values: dict[str, object] = {}
        for layer in layers:
            values.update(layer.model_dump(exclude_unset=True))
        return ParagraphProperties.model_validate(values)

    def _resolve_theme_fonts(self, run: RunProperties) -> RunProperties:
        return run.model_copy(
            update={
                "font_family": ThemeFontResolver.resolve(
                    run.font_family, self._theme_fonts, east_asia=False
                ),
                "ascii_font": ThemeFontResolver.resolve(
                    run.ascii_font, self._theme_fonts, east_asia=False
                ),
                "high_ansi_font": ThemeFontResolver.resolve(
                    run.high_ansi_font, self._theme_fonts, east_asia=False
                ),
                "east_asia_font": ThemeFontResolver.resolve(
                    run.east_asia_font, self._theme_fonts, east_asia=True
                ),
                "complex_script_font": ThemeFontResolver.resolve(
                    run.complex_script_font, self._theme_fonts, east_asia=False
                ),
            }
        )
