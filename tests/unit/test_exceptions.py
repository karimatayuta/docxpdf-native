from __future__ import annotations

import pytest

from docxpdf_native.exceptions import (
    DocxPdfError,
    FontNotFoundError,
    InvalidDocxError,
    InvalidOoxmlError,
    LayoutError,
    MissingPartError,
    PageLimitExceededError,
    PdfGenerationError,
    RelationshipError,
    ResourceLimitError,
    UnsupportedFeatureError,
)
from docxpdf_native.models import UnsupportedFeature


@pytest.mark.parametrize(
    "error_type",
    [
        InvalidDocxError,
        InvalidOoxmlError,
        MissingPartError,
        RelationshipError,
        FontNotFoundError,
        LayoutError,
        PageLimitExceededError,
        PdfGenerationError,
        ResourceLimitError,
    ],
)
def test_specific_errors_inherit_from_base(error_type: type[DocxPdfError]) -> None:
    error = error_type("failure", part="word/document.xml", location="/w:document")

    assert isinstance(error, DocxPdfError)


def test_error_message_contains_part_and_location() -> None:
    error = InvalidOoxmlError(
        "Malformed XML",
        part="word/document.xml",
        location="/w:document/w:body",
    )

    assert str(error) == ("Malformed XML [part=word/document.xml, location=/w:document/w:body]")


def test_unsupported_feature_error_preserves_structured_context() -> None:
    feature = UnsupportedFeature(
        name="text_box",
        part="word/document.xml",
        element="w:txbxContent",
        location="/w:document/w:body/w:p[1]",
        status="unsupported",
        workaround="Use an ordinary paragraph.",
    )

    error = UnsupportedFeatureError(feature)

    assert error.feature == feature
    assert "text_box" in str(error)
    assert "word/document.xml" in str(error)


def test_font_not_found_error_names_requested_font() -> None:
    error = FontNotFoundError.for_font("Missing Font", paragraph_index=3, run_index=4)

    assert error.requested_font == "Missing Font"
    assert "paragraph[3]/run[4]" in str(error)


def test_page_limit_error_reports_actual_and_allowed_counts() -> None:
    error = PageLimitExceededError.for_limit(actual=11, maximum=10)

    assert error.actual == 11
    assert error.maximum == 10
    assert "11" in str(error)
