from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Literal

from pydantic import Field

from docxpdf_native.models.base import FrozenModel
from docxpdf_native.models.fonts import FontSubstitution


class UnsupportedFeature(FrozenModel):
    name: str
    part: str | None = None
    element: str | None = None
    location: str | None = None
    status: str = "unsupported"
    workaround: str | None = None


class ConversionWarning(FrozenModel):
    code: str
    message: str
    severity: Literal["warning", "error"] = "warning"
    part: str | None = None
    location: str | None = None
    feature: UnsupportedFeature | None = None


class DiagnosticInfo(FrozenModel):
    warnings: tuple[ConversionWarning, ...] = ()
    unsupported_features: tuple[UnsupportedFeature, ...] = ()
    font_substitutions: tuple[FontSubstitution, ...] = ()
    counters: Mapping[str, int] = Field(default_factory=dict)
    details: Mapping[str, str] = Field(default_factory=dict)


class ConversionResult(FrozenModel):
    page_count: int = Field(ge=0)
    warnings: tuple[ConversionWarning, ...] = ()
    unsupported_features: tuple[UnsupportedFeature, ...] = ()
    font_substitutions: tuple[FontSubstitution, ...] = ()
    destination: Path | None = None
    pdf_bytes: bytes | None = Field(default=None, exclude=True, repr=False)
    pdf_sha256: str | None = None
    layout_sha256: str | None = None
    layout_json: str | None = Field(default=None, exclude=True, repr=False)
    diagnostics: DiagnosticInfo = Field(default_factory=DiagnosticInfo)
