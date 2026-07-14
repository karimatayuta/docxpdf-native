from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from pydantic import Field

from docxpdf_native.models.base import FrozenModel


class FontConfiguration(FrozenModel):
    font_directories: tuple[Path, ...] = ()
    registered_fonts: Mapping[str, Path] = Field(default_factory=dict)
    substitutions: Mapping[str, str] = Field(default_factory=dict)
    default_font: str | None = None
    environment_variable: str = "DOCXPDF_NATIVE_FONT_DIRS"
    include_system_fonts: bool = True


class ResolvedFont(FrozenModel):
    family: str
    path: Path | None = None
    postscript_name: str | None = None
    source: str
    substituted: bool = False


class FontMetrics(FrozenModel):
    units_per_em: int = Field(gt=0)
    ascent: float
    descent: float
    line_gap: float = 0


class TextMeasurement(FrozenModel):
    width: float = Field(ge=0)
    ascent: float = Field(ge=0)
    descent: float = Field(ge=0)

    @property
    def height(self) -> float:
        return self.ascent + self.descent


class FontSubstitution(FrozenModel):
    requested_font: str
    selected_font: str
    reason: str
    paragraph_index: int | None = Field(default=None, ge=0)
    run_index: int | None = Field(default=None, ge=0)
    east_asia: bool = False
