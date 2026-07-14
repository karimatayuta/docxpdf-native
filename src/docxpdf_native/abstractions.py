from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import BinaryIO

from docxpdf_native.models.document import DocumentModel, ResolvedDocumentModel
from docxpdf_native.models.fonts import ResolvedFont, TextMeasurement
from docxpdf_native.models.layout import LayoutDocument
from docxpdf_native.models.options import ConversionOptions
from docxpdf_native.models.results import ConversionWarning, UnsupportedFeature


class DocumentParser(ABC):
    """Turn a validated DOCX byte sequence into a renderer-independent model."""

    @abstractmethod
    def parse(self, source: bytes, *, options: ConversionOptions) -> DocumentModel:
        raise NotImplementedError


class StyleResolver(ABC):
    """Resolve style inheritance and direct formatting into effective properties."""

    @abstractmethod
    def resolve(self, document: DocumentModel) -> ResolvedDocumentModel:
        raise NotImplementedError


class FontResolver(ABC):
    """Resolve a requested family without silently changing font identity."""

    @abstractmethod
    def resolve(
        self,
        font_name: str,
        *,
        east_asia: bool = False,
        paragraph_index: int | None = None,
        run_index: int | None = None,
    ) -> ResolvedFont:
        raise NotImplementedError


class TextMeasurer(ABC):
    """Measure text using the metrics of a concrete resolved font."""

    @abstractmethod
    def measure(
        self,
        text: str,
        font: ResolvedFont,
        font_size: float,
        *,
        character_spacing: float = 0,
    ) -> TextMeasurement:
        raise NotImplementedError


class LayoutEngine(ABC):
    """Lay out resolved content in top-left-origin point coordinates."""

    @abstractmethod
    def layout(
        self,
        document: ResolvedDocumentModel,
        *,
        options: ConversionOptions,
    ) -> LayoutDocument:
        raise NotImplementedError


class PdfBackend(ABC):
    """Render a completed layout without making pagination decisions."""

    @abstractmethod
    def render(
        self,
        document: LayoutDocument,
        destination: Path | BinaryIO | None = None,
    ) -> bytes:
        raise NotImplementedError


class UnsupportedFeatureHandler(ABC):
    """Either reject or record each detected unsupported OOXML feature."""

    @abstractmethod
    def handle(self, feature: UnsupportedFeature) -> ConversionWarning | None:
        raise NotImplementedError
