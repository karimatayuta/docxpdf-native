from __future__ import annotations

from io import BytesIO

import pytest

from docxpdf_native.abstractions import (
    DocumentParser,
    FontResolver,
    LayoutEngine,
    PdfBackend,
    StyleResolver,
    TextMeasurer,
    UnsupportedFeatureHandler,
)
from docxpdf_native.models import (
    ConversionOptions,
    ConversionWarning,
    DocumentModel,
    LayoutDocument,
    ResolvedDocumentModel,
    ResolvedFont,
    TextMeasurement,
    UnsupportedFeature,
)


@pytest.mark.parametrize(
    "contract",
    [
        DocumentParser,
        StyleResolver,
        FontResolver,
        TextMeasurer,
        LayoutEngine,
        PdfBackend,
        UnsupportedFeatureHandler,
    ],
)
def test_contracts_cannot_be_instantiated(contract: type[object]) -> None:
    with pytest.raises(TypeError):
        contract()


class ParserImplementation(DocumentParser):
    def parse(self, source: bytes, *, options: ConversionOptions) -> DocumentModel:
        return DocumentModel()


class StyleImplementation(StyleResolver):
    def resolve(self, document: DocumentModel) -> ResolvedDocumentModel:
        return ResolvedDocumentModel()


class FontImplementation(FontResolver):
    def resolve(
        self,
        font_name: str,
        *,
        east_asia: bool = False,
        paragraph_index: int | None = None,
        run_index: int | None = None,
    ) -> ResolvedFont:
        return ResolvedFont(family=font_name, source="registered")


class MeasurementImplementation(TextMeasurer):
    def measure(
        self,
        text: str,
        font: ResolvedFont,
        font_size: float,
        *,
        character_spacing: float = 0,
    ) -> TextMeasurement:
        return TextMeasurement(width=len(text) * font_size, ascent=font_size, descent=0)


class LayoutImplementation(LayoutEngine):
    def layout(
        self,
        document: ResolvedDocumentModel,
        *,
        options: ConversionOptions,
    ) -> LayoutDocument:
        return LayoutDocument()


class PdfImplementation(PdfBackend):
    def render(
        self,
        document: LayoutDocument,
        destination: BytesIO | None = None,
    ) -> bytes:
        payload = b"%PDF"
        if destination is not None:
            destination.write(payload)
        return payload


class UnsupportedImplementation(UnsupportedFeatureHandler):
    def handle(self, feature: UnsupportedFeature) -> ConversionWarning | None:
        return ConversionWarning(code="unsupported", message=feature.name, feature=feature)


def test_contracts_allow_typed_implementations() -> None:
    options = ConversionOptions()
    document = ParserImplementation().parse(b"docx", options=options)
    resolved = StyleImplementation().resolve(document)
    font = FontImplementation().resolve("Example")
    measurement = MeasurementImplementation().measure("a", font, 10)
    layout = LayoutImplementation().layout(resolved, options=options)
    pdf = PdfImplementation().render(layout)
    warning = UnsupportedImplementation().handle(UnsupportedFeature(name="field"))

    assert measurement.width == 10
    assert pdf == b"%PDF"
    assert warning is not None
