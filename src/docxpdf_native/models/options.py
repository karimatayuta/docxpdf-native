from __future__ import annotations

from pydantic import Field

from docxpdf_native.models.base import FrozenModel
from docxpdf_native.models.fonts import FontConfiguration


class ResourceLimits(FrozenModel):
    max_docx_file_size: int = Field(default=100 * 1024 * 1024, gt=0)
    max_expanded_size: int = Field(default=512 * 1024 * 1024, gt=0)
    max_xml_part_size: int = Field(default=32 * 1024 * 1024, gt=0)
    max_image_size: int = Field(default=64 * 1024 * 1024, gt=0)
    max_xml_depth: int = Field(default=128, gt=0)
    max_paragraphs: int = Field(default=100_000, gt=0)
    max_runs: int = Field(default=1_000_000, gt=0)
    max_table_cells: int = Field(default=100_000, gt=0)
    max_pages: int = Field(default=10_000, gt=0)


class ConversionOptions(FrozenModel):
    strict: bool = True
    deterministic: bool = True
    font_configuration: FontConfiguration = Field(default_factory=FontConfiguration)
    resource_limits: ResourceLimits = Field(default_factory=ResourceLimits)
    allow_external_relationships: bool = False
