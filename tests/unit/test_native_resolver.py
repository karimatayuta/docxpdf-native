from __future__ import annotations

from docxpdf_native.abstractions import FontResolver
from docxpdf_native.models import (
    DocumentDefaults,
    DocumentModel,
    ParagraphModel,
    ResolvedFont,
    RunModel,
    RunProperties,
    SectionModel,
)
from docxpdf_native.resolver import NativeStyleResolver


class _RecordingFontResolver(FontResolver):
    def __init__(self) -> None:
        self.calls: list[tuple[str, bool]] = []

    def resolve(
        self,
        font_name: str,
        *,
        east_asia: bool = False,
        paragraph_index: int | None = None,
        run_index: int | None = None,
    ) -> ResolvedFont:
        self.calls.append((font_name, east_asia))
        return ResolvedFont(
            family=font_name,
            postscript_name=font_name,
            source="test",
        )


def test_native_resolver_splits_mixed_script_run_between_declared_fonts() -> None:
    document = DocumentModel(
        defaults=DocumentDefaults(
            run=RunProperties(
                ascii_font="Latin Face",
                high_ansi_font="Latin Face",
                east_asia_font="Japanese Face",
            ),
        ),
        sections=(
            SectionModel(
                blocks=(ParagraphModel(runs=(RunModel(text="ABC日本語123"),)),),
            ),
        ),
    )
    font_resolver = _RecordingFontResolver()

    resolved = NativeStyleResolver(font_resolver).resolve(document)

    runs = resolved.paragraphs[0].runs
    assert [(run.text, run.font_name) for run in runs] == [
        ("ABC", "Latin Face"),
        ("日本語", "Japanese Face"),
        ("123", "Latin Face"),
    ]
    assert font_resolver.calls == [
        ("Latin Face", False),
        ("Japanese Face", True),
        ("Latin Face", False),
    ]
