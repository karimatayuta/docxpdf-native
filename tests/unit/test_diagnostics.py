from __future__ import annotations

from pathlib import Path

import pytest
from pydantic import BaseModel, ConfigDict

from docxpdf_native.diagnostics import DiagnosticsExporter


class _SampleDiagnostics(BaseModel):
    model_config = ConfigDict(frozen=True)

    zeta: int
    alpha: str


def test_diagnostics_json_is_sorted_and_utf8() -> None:
    model = _SampleDiagnostics(zeta=2, alpha="日本語")

    value = DiagnosticsExporter.to_json(model)

    assert value == '{"alpha":"日本語","zeta":2}\n'


def test_diagnostics_write_uses_utf8(tmp_path: Path) -> None:
    destination = tmp_path / "diagnostics.json"

    DiagnosticsExporter.write(_SampleDiagnostics(zeta=2, alpha="日本語"), destination)

    assert destination.read_text(encoding="utf-8") == '{"alpha":"日本語","zeta":2}\n'


def test_diagnostics_write_rejects_symbolic_link(tmp_path: Path) -> None:
    actual = tmp_path / "actual.json"
    actual.write_text("keep", encoding="utf-8")
    destination = tmp_path / "diagnostics.json"
    destination.symlink_to(actual)

    with pytest.raises(ValueError, match="symbolic link"):
        DiagnosticsExporter.write(_SampleDiagnostics(zeta=2, alpha="日本語"), destination)
