from __future__ import annotations

import hashlib
import sys
from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen.canvas import Canvas

sys.path.insert(0, str(Path(__file__).parents[2]))

from benchmark.pdf_compare import (
    compare_pdfs,
    normalize_comparison_text,
    normalize_pdf_text,
    snapshot_pdf,
)


def _write_pdf(
    destination: Path,
    pages: tuple[str | None, ...],
    *,
    page_size: tuple[float, float] = A4,
) -> None:
    pdf = Canvas(str(destination), pagesize=page_size, invariant=1)
    for text in pages:
        if text is not None:
            pdf.drawString(72.0, page_size[1] - 72.0, text)
        pdf.showPage()
    pdf.save()


def _set_page_boxes_and_rotation(source: Path) -> None:
    reader = PdfReader(source)
    page = reader.pages[0]
    page.cropbox.lower_left = (10, 20)
    page.cropbox.upper_right = (500, 800)
    page.trimbox.lower_left = (30, 40)
    page.trimbox.upper_right = (480, 780)
    page.rotate(90)
    writer = PdfWriter()
    writer.add_page(page)
    with source.open("wb") as stream:
        writer.write(stream)


def test_normalize_pdf_text_uses_nfc_and_canonical_newlines() -> None:
    assert normalize_pdf_text("Cafe\u0301\r\n次\r頁") == "Caf\u00e9\n次\n頁"


def test_normalize_comparison_text_removes_unicode_whitespace() -> None:
    assert normalize_comparison_text(" Cafe\u0301\r\n次\t頁 ") == "Caf\u00e9次頁"


def test_snapshot_records_page_geometry_text_and_boundaries(tmp_path: Path) -> None:
    source = tmp_path / "snapshot.pdf"
    _write_pdf(source, ("first", None))
    _set_page_boxes_and_rotation(source)

    snapshot = snapshot_pdf(source)

    assert snapshot.sha256 == hashlib.sha256(source.read_bytes()).hexdigest()
    assert snapshot.page_count == 1
    page = snapshot.pages[0]
    assert page.media_box.width == pytest.approx(A4[0])
    assert page.media_box.height == pytest.approx(A4[1])
    assert page.crop_box.coordinates == pytest.approx((10.0, 20.0, 500.0, 800.0))
    assert page.trim_box.coordinates == pytest.approx((30.0, 40.0, 480.0, 780.0))
    assert page.rotation == 90
    assert page.text == "first\n"
    assert page.comparison_text == "first"
    assert page.first_200 == "first"
    assert page.last_200 == "first"
    assert snapshot.full_text == "first\n"
    assert snapshot.comparison_text == "first"
    assert snapshot.empty_pages == ()
    assert snapshot.page_boundaries == (5,)


def test_snapshot_accepts_pdf_bytes_and_marks_empty_pages(tmp_path: Path) -> None:
    source = tmp_path / "empty-page.pdf"
    _write_pdf(source, ("body", None))

    snapshot = snapshot_pdf(source.read_bytes())

    assert tuple(page.text for page in snapshot.pages) == ("body\n", "")
    assert snapshot.empty_pages == (2,)
    assert snapshot.page_boundaries == (4, 4)


def test_compare_returns_exact_match_without_requiring_binary_sha_match(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("same",))
    candidate.write_bytes(oracle.read_bytes() + b"\n% deterministic but distinct container bytes\n")

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.oracle is not None
    assert comparison.candidate is not None
    assert comparison.oracle.sha256 != comparison.candidate.sha256
    assert comparison.verdict == "exact_match"
    assert comparison.reasons == ()


def test_compare_returns_mismatch_when_page_text_differs(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("AAAA",))
    _write_pdf(candidate, ("BBBBBB",))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "mismatch"
    assert comparison.reasons == (
        "full text differs",
        "page 1 text differs",
        "page boundaries differ",
    )
    assert comparison.metrics.missing_text_chars == 4
    assert comparison.metrics.duplicated_text_chars == 6


def test_compare_returns_layout_match_for_whitespace_only_extraction_difference(
    tmp_path: Path,
) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("same text",))
    _write_pdf(candidate, ("same  text",))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "layout_match"
    assert comparison.metrics.text_sequence_similarity == 1.0


def test_compare_reports_page_boundary_drift_for_same_full_text(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("abc", "def"))
    _write_pdf(candidate, ("abcdef", None))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.metrics.page_boundary_exact_matches == 1
    assert comparison.metrics.page_boundary_match_rate == 0.5
    assert comparison.metrics.mean_boundary_drift_chars == 1.5
    assert comparison.metrics.max_boundary_drift_chars == 3
    assert comparison.metrics.first_different_page == 1
    assert comparison.metrics.missing_text_chars == 0
    assert comparison.metrics.duplicated_text_chars == 0
    assert comparison.metrics.text_sequence_similarity == 1.0


def test_compare_returns_mismatch_when_page_geometry_differs(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("same",), page_size=A4)
    _write_pdf(candidate, ("same",), page_size=(A4[0] + 10.0, A4[1]))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "mismatch"
    assert comparison.reasons == (
        "page 1 media box differs",
        "page 1 crop box differs",
        "page 1 trim box differs",
    )


def test_compare_accepts_page_geometry_within_point_one_point(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("same",), page_size=A4)
    _write_pdf(candidate, ("same",), page_size=(A4[0] + 0.05, A4[1] - 0.05))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "exact_match"


def test_compare_returns_mismatch_when_page_count_differs(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("first",))
    _write_pdf(candidate, ("first", "second"))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "mismatch"
    assert comparison.reasons[0] == "page count differs: oracle=1, candidate=2"


def test_compare_returns_invalid_oracle_for_unreadable_reference(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    oracle.write_bytes(b"not a PDF")
    _write_pdf(candidate, ("candidate",))

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "invalid_oracle"
    assert comparison.oracle is None
    assert comparison.candidate is None
    assert comparison.reasons == ("oracle PDF is invalid",)


def test_compare_returns_mismatch_for_unreadable_candidate(tmp_path: Path) -> None:
    oracle = tmp_path / "oracle.pdf"
    candidate = tmp_path / "candidate.pdf"
    _write_pdf(oracle, ("oracle",))
    candidate.write_bytes(b"not a PDF")

    comparison = compare_pdfs(oracle, candidate)

    assert comparison.verdict == "mismatch"
    assert comparison.oracle is not None
    assert comparison.candidate is None
    assert comparison.reasons == ("candidate PDF is invalid",)
