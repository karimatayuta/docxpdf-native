"""Deterministic structural comparison for an oracle PDF and a candidate PDF."""

from __future__ import annotations

import hashlib
import unicodedata
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field
from pypdf import PdfReader
from pypdf.errors import PdfReadError
from pypdf.generic import RectangleObject

PdfSource = Path | str | bytes
PdfComparisonVerdict = Literal[
    "exact_match",
    "layout_match",
    "mismatch",
    "invalid_oracle",
]


class _FrozenModel(BaseModel):
    model_config = ConfigDict(frozen=True)


class PdfBoxSnapshot(_FrozenModel):
    """One PDF page box in native PDF coordinates."""

    left: float
    bottom: float
    right: float
    top: float

    @property
    def coordinates(self) -> tuple[float, float, float, float]:
        return (self.left, self.bottom, self.right, self.top)

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.top - self.bottom


class PdfPageSnapshot(_FrozenModel):
    """Comparable state extracted from one PDF page."""

    number: int = Field(ge=1)
    media_box: PdfBoxSnapshot
    crop_box: PdfBoxSnapshot
    trim_box: PdfBoxSnapshot
    rotation: int
    text: str
    comparison_text: str
    first_200: str
    last_200: str
    empty: bool


class PdfSnapshot(_FrozenModel):
    """Comparable document state independent of PDF container bytes."""

    sha256: str
    page_count: int = Field(ge=0)
    pages: tuple[PdfPageSnapshot, ...] = ()
    full_text: str = ""
    comparison_text: str = ""
    empty_pages: tuple[int, ...] = ()
    page_boundaries: tuple[int, ...] = ()


class PdfComparisonMetrics(_FrozenModel):
    """Text-integrity and pagination metrics requested by the corpus protocol."""

    page_boundary_exact_matches: int = Field(default=0, ge=0)
    page_boundary_match_rate: float = Field(default=0.0, ge=0.0, le=1.0)
    mean_boundary_drift_chars: float = Field(default=0.0, ge=0.0)
    max_boundary_drift_chars: int = Field(default=0, ge=0)
    first_different_page: int | None = Field(default=None, ge=1)
    missing_text_chars: int = Field(default=0, ge=0)
    duplicated_text_chars: int = Field(default=0, ge=0)
    text_sequence_similarity: float = Field(default=0.0, ge=0.0, le=1.0)


class PdfComparison(_FrozenModel):
    """A classified comparison with stable, human-readable reasons."""

    verdict: PdfComparisonVerdict
    reasons: tuple[str, ...] = ()
    oracle: PdfSnapshot | None = None
    candidate: PdfSnapshot | None = None
    metrics: PdfComparisonMetrics = Field(default_factory=PdfComparisonMetrics)


class PdfSnapshotError(ValueError):
    """Raised when a PDF cannot be read into a reliable snapshot."""


def normalize_pdf_text(text: str) -> str:
    """Normalize PDF-extracted text to NFC with canonical LF newlines."""

    canonical_newlines = text.replace("\r\n", "\n").replace("\r", "\n")
    return unicodedata.normalize("NFC", canonical_newlines)


def normalize_comparison_text(text: str) -> str:
    """Return NFC text without extraction-dependent Unicode whitespace."""

    return "".join(normalize_pdf_text(text).split())


def snapshot_pdf(source: PdfSource) -> PdfSnapshot:
    """Read a PDF path or byte string into a deterministic comparison snapshot."""

    try:
        payload = _read_payload(source)
        reader = PdfReader(BytesIO(payload), strict=False)
        pages: list[PdfPageSnapshot] = []
        boundary = 0
        boundaries: list[int] = []
        empty_pages: list[int] = []
        text_parts: list[str] = []
        comparison_parts: list[str] = []
        for page_number, page in enumerate(reader.pages, start=1):
            text = normalize_pdf_text(page.extract_text() or "")
            comparison_text = normalize_comparison_text(text)
            empty = not comparison_text
            if empty:
                empty_pages.append(page_number)
            text_parts.append(text)
            comparison_parts.append(comparison_text)
            boundary += len(comparison_text)
            boundaries.append(boundary)
            pages.append(
                PdfPageSnapshot(
                    number=page_number,
                    media_box=_box_snapshot(page.mediabox),
                    crop_box=_box_snapshot(page.cropbox),
                    trim_box=_box_snapshot(page.trimbox),
                    rotation=int(page.rotation or 0) % 360,
                    text=text,
                    comparison_text=comparison_text,
                    first_200=comparison_text[:200],
                    last_200=comparison_text[-200:],
                    empty=empty,
                )
            )
        return PdfSnapshot(
            sha256=hashlib.sha256(payload).hexdigest(),
            page_count=len(pages),
            pages=tuple(pages),
            full_text="".join(text_parts),
            comparison_text="".join(comparison_parts),
            empty_pages=tuple(empty_pages),
            page_boundaries=tuple(boundaries),
        )
    except (OSError, PdfReadError, TypeError, ValueError) as error:
        raise PdfSnapshotError("PDF could not be read") from error


def compare_pdfs(oracle_source: PdfSource, candidate_source: PdfSource) -> PdfComparison:
    """Compare oracle and candidate without treating their container SHA as a match gate."""

    try:
        oracle = snapshot_pdf(oracle_source)
    except PdfSnapshotError:
        return PdfComparison(
            verdict="invalid_oracle",
            reasons=("oracle PDF is invalid",),
        )
    try:
        candidate = snapshot_pdf(candidate_source)
    except PdfSnapshotError:
        return PdfComparison(
            verdict="mismatch",
            reasons=("candidate PDF is invalid",),
            oracle=oracle,
        )

    layout_reasons = _layout_reasons(oracle, candidate)
    text_reasons = _text_reasons(oracle, candidate)
    metrics = _comparison_metrics(oracle, candidate)
    if layout_reasons or text_reasons:
        return PdfComparison(
            verdict="mismatch",
            reasons=(*layout_reasons, *text_reasons),
            oracle=oracle,
            candidate=candidate,
            metrics=metrics,
        )
    if _raw_extraction_differs(oracle, candidate):
        return PdfComparison(
            verdict="layout_match",
            reasons=("only extraction whitespace differs",),
            oracle=oracle,
            candidate=candidate,
            metrics=metrics,
        )
    return PdfComparison(
        verdict="exact_match",
        oracle=oracle,
        candidate=candidate,
        metrics=metrics,
    )


def _read_payload(source: PdfSource) -> bytes:
    if isinstance(source, bytes):
        return source
    return Path(source).read_bytes()


def _box_snapshot(box: RectangleObject) -> PdfBoxSnapshot:
    return PdfBoxSnapshot(
        left=float(box.left),
        bottom=float(box.bottom),
        right=float(box.right),
        top=float(box.top),
    )


def _layout_reasons(oracle: PdfSnapshot, candidate: PdfSnapshot) -> tuple[str, ...]:
    reasons: list[str] = []
    if oracle.page_count != candidate.page_count:
        reasons.append(
            f"page count differs: oracle={oracle.page_count}, candidate={candidate.page_count}"
        )
    for page_number, (oracle_page, candidate_page) in enumerate(
        zip(oracle.pages, candidate.pages, strict=False),
        start=1,
    ):
        for label, oracle_box, candidate_box in (
            ("media", oracle_page.media_box, candidate_page.media_box),
            ("crop", oracle_page.crop_box, candidate_page.crop_box),
            ("trim", oracle_page.trim_box, candidate_page.trim_box),
        ):
            if not _boxes_equal(oracle_box, candidate_box):
                reasons.append(f"page {page_number} {label} box differs")
        if oracle_page.rotation != candidate_page.rotation:
            reasons.append(f"page {page_number} rotation differs")
    if oracle.empty_pages != candidate.empty_pages:
        reasons.append("empty page positions differ")
    return tuple(reasons)


def _text_reasons(oracle: PdfSnapshot, candidate: PdfSnapshot) -> tuple[str, ...]:
    reasons: list[str] = []
    if oracle.comparison_text != candidate.comparison_text:
        reasons.append("full text differs")
    for page_number, (oracle_page, candidate_page) in enumerate(
        zip(oracle.pages, candidate.pages, strict=False),
        start=1,
    ):
        if oracle_page.comparison_text != candidate_page.comparison_text:
            reasons.append(f"page {page_number} text differs")
    if oracle.page_boundaries != candidate.page_boundaries:
        reasons.append("page boundaries differ")
    return tuple(reasons)


def _raw_extraction_differs(oracle: PdfSnapshot, candidate: PdfSnapshot) -> bool:
    return oracle.full_text != candidate.full_text or any(
        oracle_page.text != candidate_page.text
        for oracle_page, candidate_page in zip(
            oracle.pages,
            candidate.pages,
            strict=False,
        )
    )


def _comparison_metrics(oracle: PdfSnapshot, candidate: PdfSnapshot) -> PdfComparisonMetrics:
    matcher = SequenceMatcher(
        None,
        oracle.comparison_text,
        candidate.comparison_text,
        autojunk=False,
    )
    missing = 0
    duplicated = 0
    for (
        operation,
        oracle_start,
        oracle_end,
        candidate_start,
        candidate_end,
    ) in matcher.get_opcodes():
        if operation in {"delete", "replace"}:
            missing += oracle_end - oracle_start
        if operation in {"insert", "replace"}:
            duplicated += candidate_end - candidate_start

    drifts = tuple(
        _map_candidate_boundary(matcher, boundary) - oracle.page_boundaries[index]
        for index, boundary in enumerate(candidate.page_boundaries[: len(oracle.page_boundaries)])
    )
    total_boundaries = max(len(oracle.page_boundaries), len(candidate.page_boundaries))
    exact_matches = sum(drift == 0 for drift in drifts)
    absolute_drifts = tuple(abs(drift) for drift in drifts)
    first_different = next(
        (index for index, drift in enumerate(drifts, start=1) if drift != 0),
        None,
    )
    if first_different is None and len(oracle.page_boundaries) != len(candidate.page_boundaries):
        first_different = len(drifts) + 1
    return PdfComparisonMetrics(
        page_boundary_exact_matches=exact_matches,
        page_boundary_match_rate=(exact_matches / total_boundaries if total_boundaries else 1.0),
        mean_boundary_drift_chars=(
            sum(absolute_drifts) / len(absolute_drifts) if absolute_drifts else 0.0
        ),
        max_boundary_drift_chars=max(absolute_drifts, default=0),
        first_different_page=first_different,
        missing_text_chars=missing,
        duplicated_text_chars=duplicated,
        text_sequence_similarity=matcher.ratio(),
    )


def _map_candidate_boundary(matcher: SequenceMatcher[str], boundary: int) -> int:
    for (
        operation,
        oracle_start,
        oracle_end,
        candidate_start,
        candidate_end,
    ) in matcher.get_opcodes():
        if not candidate_start <= boundary <= candidate_end:
            continue
        if operation == "equal":
            return oracle_start + min(boundary - candidate_start, oracle_end - oracle_start)
        if operation == "insert" or candidate_end == candidate_start:
            return oracle_start
        span = (boundary - candidate_start) / (candidate_end - candidate_start)
        return oracle_start + round(span * (oracle_end - oracle_start))
    return len(matcher.a)


def _boxes_equal(first: PdfBoxSnapshot, second: PdfBoxSnapshot) -> bool:
    tolerance = 0.1
    return all(
        abs(first_value - second_value) <= tolerance
        for first_value, second_value in zip(
            first.coordinates,
            second.coordinates,
            strict=True,
        )
    )
