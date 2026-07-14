"""JSON exporters for Pydantic diagnostic and layout models."""

from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel

from docxpdf_native.security.paths import OutputPathGuard


class DiagnosticsExporter:
    """Serialize diagnostic models without locale- or time-dependent values."""

    @staticmethod
    def to_json(model: BaseModel) -> str:
        """Return compact, stable UTF-8 JSON terminated by one newline."""
        value = model.model_dump(mode="json", exclude_none=True)
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n"

    @staticmethod
    def write(model: BaseModel, destination: Path) -> None:
        """Write a diagnostic model as UTF-8 JSON."""
        if destination.is_symlink():
            raise ValueError(f"Refusing to overwrite symbolic link: {destination}")
        symlinked_parent = OutputPathGuard.symlinked_parent(destination)
        if symlinked_parent is not None:
            raise ValueError(f"Refusing output through symbolic link directory: {symlinked_parent}")
        destination.write_text(DiagnosticsExporter.to_json(model), encoding="utf-8")
