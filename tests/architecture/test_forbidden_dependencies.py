from __future__ import annotations

from pathlib import Path

import pytest

SOURCE_ROOT = Path(__file__).parents[2] / "src" / "docxpdf_native"


@pytest.mark.parametrize(
    "forbidden",
    (
        "dataclasses",
        "@dataclass",
        "subprocess",
        "soffice",
        "libreoffice",
        "win32com",
        "python-docx",
        "docx2pdf",
        "pandoc",
        "playwright",
        "selenium",
    ),
)
def test_source_does_not_use_forbidden_implementation(forbidden: str) -> None:
    matches = tuple(
        str(path.relative_to(SOURCE_ROOT))
        for path in SOURCE_ROOT.rglob("*.py")
        if forbidden in path.read_text(encoding="utf-8").casefold()
    )

    assert matches == ()
