from __future__ import annotations

from typing import Self

from docxpdf_native.models.results import UnsupportedFeature


class DocxPdfError(Exception):
    """Base class for expected conversion failures with OOXML context."""

    def __init__(
        self,
        message: str,
        *,
        part: str | None = None,
        location: str | None = None,
        cause: Exception | None = None,
    ) -> None:
        self.message = message
        self.part = part
        self.location = location
        self.cause = cause
        super().__init__(self._formatted_message())

    def _formatted_message(self) -> str:
        context: list[str] = []
        if self.part is not None:
            context.append(f"part={self.part}")
        if self.location is not None:
            context.append(f"location={self.location}")
        if not context:
            return self.message
        return f"{self.message} [{', '.join(context)}]"


class InvalidDocxError(DocxPdfError):
    pass


class InvalidOoxmlError(DocxPdfError):
    pass


class MissingPartError(DocxPdfError):
    pass


class RelationshipError(DocxPdfError):
    pass


class UnsupportedFeatureError(DocxPdfError):
    def __init__(self, feature: UnsupportedFeature) -> None:
        self.feature = feature
        message = f"unsupported feature: {feature.name} (status={feature.status})"
        if feature.element is not None:
            message = f"{message}; element={feature.element}"
        if feature.workaround is not None:
            message = f"{message}; workaround={feature.workaround}"
        super().__init__(message, part=feature.part, location=feature.location)


class FontNotFoundError(DocxPdfError):
    def __init__(
        self,
        message: str,
        *,
        requested_font: str | None = None,
        part: str | None = None,
        location: str | None = None,
        cause: Exception | None = None,
    ) -> None:
        self.requested_font = requested_font
        super().__init__(message, part=part, location=location, cause=cause)

    @classmethod
    def for_font(
        cls,
        requested_font: str,
        *,
        paragraph_index: int | None = None,
        run_index: int | None = None,
    ) -> Self:
        path_parts: list[str] = []
        if paragraph_index is not None:
            path_parts.append(f"paragraph[{paragraph_index}]")
        if run_index is not None:
            path_parts.append(f"run[{run_index}]")
        return cls(
            f"font not found: {requested_font}",
            requested_font=requested_font,
            location="/".join(path_parts) or None,
        )


class LayoutError(DocxPdfError):
    pass


class PageLimitExceededError(LayoutError):
    def __init__(
        self,
        message: str,
        *,
        actual: int | None = None,
        maximum: int | None = None,
        part: str | None = None,
        location: str | None = None,
        cause: Exception | None = None,
    ) -> None:
        self.actual = actual
        self.maximum = maximum
        super().__init__(message, part=part, location=location, cause=cause)

    @classmethod
    def for_limit(cls, *, actual: int, maximum: int) -> Self:
        return cls(
            f"page limit exceeded: produced {actual} pages, maximum is {maximum}",
            actual=actual,
            maximum=maximum,
        )


class PdfGenerationError(DocxPdfError):
    pass


class ResourceLimitError(DocxPdfError):
    pass
