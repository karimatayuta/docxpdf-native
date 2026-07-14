"""Resource-limit helpers for untrusted DOCX input."""

from __future__ import annotations

from docxpdf_native.security.limits import ResourceLimitGuard, ResourceLimits

__all__ = ["ResourceLimitGuard", "ResourceLimits"]
