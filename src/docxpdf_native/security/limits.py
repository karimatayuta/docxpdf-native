"""Consistent resource-limit checks used while reading and parsing DOCX files."""

from __future__ import annotations

from docxpdf_native.exceptions import ResourceLimitError
from docxpdf_native.models.options import ResourceLimits

__all__ = ["ResourceLimitGuard", "ResourceLimits"]


class ResourceLimitGuard:
    """Stateless helpers that raise a domain error for exceeded limits."""

    @staticmethod
    def ensure_at_most(*, actual: int, maximum: int, resource: str) -> None:
        """Raise when ``actual`` is greater than the configured maximum."""

        if actual > maximum:
            raise ResourceLimitError(
                f"{resource} is {actual} bytes/items; configured limit is {maximum}"
            )

    @staticmethod
    def checked_add(*, current: int, increment: int, maximum: int, resource: str) -> int:
        """Add a non-negative amount and enforce the result's limit."""

        if current < 0 or increment < 0:
            raise ValueError("Resource counters cannot be negative")
        result = current + increment
        ResourceLimitGuard.ensure_at_most(
            actual=result,
            maximum=maximum,
            resource=resource,
        )
        return result
