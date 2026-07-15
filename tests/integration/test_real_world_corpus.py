"""Bulk conversion tests against real-world DOCX files collected from the web.

The corpus lives in ``tests/fixtures/real_world`` together with a
``manifest.json`` recording the source URL, SHA-256, and the page count that
``docProps/app.xml`` reported when the document was last saved in Word. These
files are intentionally difficult: legacy Word generations, embedded OLE
objects, field codes, VML images, tracked changes, and footnotes.

Two guarantees are checked for every file:

1. Conversion with default options never raises.
2. The generated page count stays within tolerance of the reference count
   (the reference itself depends on the fonts of the machine that saved the
   document, so exact equality cannot be guaranteed).
"""

from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

import pytest

from docxpdf_native import Converter

FIXTURE_DIRECTORY = Path(__file__).resolve().parent.parent / "fixtures" / "real_world"
MANIFEST_PATH = FIXTURE_DIRECTORY / "manifest.json"

PAGE_COUNT_TOLERANCE = 0.15


def _manifest_entries() -> list[dict[str, Any]]:
    if not MANIFEST_PATH.is_file():
        return []
    with MANIFEST_PATH.open(encoding="utf-8") as handle:
        manifest = json.load(handle)
    entries: list[dict[str, Any]] = manifest["files"]
    return [entry for entry in entries if (FIXTURE_DIRECTORY / entry["filename"]).is_file()]


_ENTRIES = _manifest_entries()

pytestmark = pytest.mark.skipif(
    not _ENTRIES,
    reason="real-world corpus fixtures are not present",
)


@pytest.fixture(scope="module")
def conversion_results() -> dict[str, Any]:
    """Convert every corpus file once and share the results across tests."""

    results: dict[str, Any] = {}
    for entry in _ENTRIES:
        source = FIXTURE_DIRECTORY / entry["filename"]
        try:
            results[entry["filename"]] = Converter().convert(source.read_bytes())
        except Exception as error:  # recorded and asserted per file
            results[entry["filename"]] = error
    return results


@pytest.mark.parametrize(
    "entry",
    _ENTRIES,
    ids=[entry["filename"] for entry in _ENTRIES],
)
def test_converts_without_error(entry: dict[str, Any], conversion_results: dict[str, Any]) -> None:
    result = conversion_results[entry["filename"]]
    assert not isinstance(result, Exception), (
        f"{entry['filename']} failed to convert: {type(result).__name__}: {result}"
    )
    assert result.pdf_bytes, f"{entry['filename']} produced no PDF bytes"
    assert result.page_count >= 1


def _has_reliable_reference(entry: dict[str, Any]) -> bool:
    return bool(entry.get("app_pages")) and entry.get("app_pages_reliable", True)


@pytest.mark.parametrize(
    "entry",
    [entry for entry in _ENTRIES if _has_reliable_reference(entry)],
    ids=[entry["filename"] for entry in _ENTRIES if _has_reliable_reference(entry)],
)
def test_page_count_close_to_reference(
    entry: dict[str, Any],
    conversion_results: dict[str, Any],
) -> None:
    result = conversion_results[entry["filename"]]
    if isinstance(result, Exception):
        pytest.fail(f"{entry['filename']} failed to convert: {result}")
    reference = int(entry["app_pages"])
    allowed = max(1, math.ceil(reference * PAGE_COUNT_TOLERANCE))
    deviation = abs(result.page_count - reference)
    assert deviation <= allowed, (
        f"{entry['filename']}: generated {result.page_count} pages, reference {reference} "
        f"(allowed deviation {allowed}, warnings {len(result.warnings)}, "
        f"substitutions {len(result.font_substitutions)})"
    )


def test_corpus_summary_report(conversion_results: dict[str, Any]) -> None:
    """Emit a per-file summary so ``pytest -s`` shows the corpus health."""

    lines = [
        f"{'file':<55} {'ref':>4} {'got':>4} {'warn':>5} {'subst':>5}",
    ]
    failures = []
    for entry in _ENTRIES:
        name = entry["filename"]
        result = conversion_results[name]
        if isinstance(result, Exception):
            failures.append(name)
            lines.append(f"{name:<55} {'-':>4} {'ERR':>4} {type(result).__name__}")
            continue
        reference = entry.get("app_pages") or "-"
        lines.append(
            f"{name:<55} {reference!s:>4} {result.page_count:>4} "
            f"{len(result.warnings):>5} {len(result.font_substitutions):>5}"
        )
    print("\n".join(lines))
    assert not failures, f"corpus files failed to convert: {failures}"
