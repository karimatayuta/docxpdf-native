"""Run reproducible docxpdf-native conversions against reference PDFs."""

from __future__ import annotations

import csv
import hashlib
import html
import time
from collections.abc import Mapping
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from benchmark.pdf_compare import PdfComparison, compare_pdfs, snapshot_pdf
from docxpdf_native import ConversionOptions, Converter, FontConfiguration
from docxpdf_native.exceptions import DocxPdfError, UnsupportedFeatureError


class _FrozenModel(BaseModel):
    model_config = ConfigDict(frozen=True)


class ManifestSource(_FrozenModel):
    path: Path
    byte_size: int = Field(ge=0)
    sha256: str
    docx_metadata_pages: int | None = Field(default=None, ge=0)


class ManifestReference(_FrozenModel):
    path: Path
    byte_size: int = Field(ge=0)
    sha256: str
    pdf_pages: int = Field(ge=0)


class ManifestDocument(_FrozenModel):
    id: str
    title: str
    revision: str
    oracle_valid: bool
    oracle_validation_notes: str
    source: ManifestSource
    reference: ManifestReference


class ReferenceManifest(_FrozenModel):
    schema_version: int = Field(ge=1)
    generated_at: str
    documents: tuple[ManifestDocument, ...]


class BenchmarkConfig(_FrozenModel):
    """Configuration for one deterministic corpus run."""

    manifest_path: Path
    output_root: Path
    repetitions: int = Field(default=3, ge=1, le=10)
    font_path: Path | None = None
    font_alias: str = "Benchmark Japanese"


class ConversionExecution(_FrozenModel):
    number: int = Field(ge=1)
    pdf_path: Path
    pdf_sha256: str
    layout_sha256: str | None = None
    page_count: int = Field(ge=0)
    duration_ms: float = Field(ge=0)


BenchmarkStatus = Literal[
    "exact_match",
    "layout_match",
    "mismatch",
    "invalid_oracle",
    "conversion_failed",
]
StrictProbeStatus = Literal["passed", "unsupported", "failed", "not_run"]


class DocumentBenchmarkResult(_FrozenModel):
    document_id: str
    title: str
    status: BenchmarkStatus
    reference_pdf_pages: int = Field(ge=0)
    candidate_pdf_pages: int | None = Field(default=None, ge=0)
    reference_rendered_pages: int | None = Field(default=None, ge=0)
    candidate_rendered_pages: int | None = Field(default=None, ge=0)
    docx_metadata_pages: int | None = Field(default=None, ge=0)
    strict_probe_status: StrictProbeStatus = "not_run"
    strict_probe_error: str | None = None
    deterministic: bool = False
    executions: tuple[ConversionExecution, ...] = ()
    warning_count: int = Field(default=0, ge=0)
    font_substitution_count: int = Field(default=0, ge=0)
    comparison: PdfComparison | None = None
    failure: str | None = None


class BenchmarkReport(_FrozenModel):
    schema_version: int = 1
    corpus_generated_at: str
    repetitions: int = Field(ge=1)
    documents: tuple[DocumentBenchmarkResult, ...]


_FONT_SUBSTITUTION_NAMES = (
    "\uff2d\uff33 ゴシック",
    "MS Gothic",
    "\uff2d\uff33 \uff30ゴシック",
    "MS PGothic",
    "\uff2d\uff33 明朝",
    "MS Mincho",
    "\uff2d\uff33",
    "游ゴシック",
    "游ゴシック Light",
    "Yu Gothic",
    "游明朝",
    "メイリオ",
    "Meiryo",
    "Meiryo UI",
    "HGP創英角ｺﾞｼｯｸUB",
    "Century",
    "Arial",
    "Calibri",
    "Courier New",
    "Times New Roman",
)


def run_benchmark(config: BenchmarkConfig) -> BenchmarkReport:
    """Convert every valid manifest document and write machine-readable results."""

    manifest = ReferenceManifest.model_validate_json(
        config.manifest_path.read_text(encoding="utf-8")
    )
    config.output_root.mkdir(parents=True, exist_ok=True)
    artifacts_root = config.output_root / "artifacts"
    artifacts_root.mkdir(exist_ok=True)
    results = tuple(
        _run_document(document, config=config, artifacts_root=artifacts_root)
        for document in manifest.documents
    )
    report = BenchmarkReport(
        corpus_generated_at=manifest.generated_at,
        repetitions=config.repetitions,
        documents=results,
    )
    _write_outputs(report, config.output_root)
    return report


def _run_document(
    document: ManifestDocument,
    *,
    config: BenchmarkConfig,
    artifacts_root: Path,
) -> DocumentBenchmarkResult:
    source = _resolve_manifest_path(config.manifest_path, document.source.path)
    reference = _resolve_manifest_path(config.manifest_path, document.reference.path)
    integrity_failure = _integrity_failure(
        source,
        document.source.sha256,
        document.source.byte_size,
    )
    integrity_failure = integrity_failure or _integrity_failure(
        reference,
        document.reference.sha256,
        document.reference.byte_size,
    )
    if not document.oracle_valid or integrity_failure is not None:
        return DocumentBenchmarkResult(
            document_id=document.id,
            title=document.title,
            status="invalid_oracle",
            reference_pdf_pages=document.reference.pdf_pages,
            docx_metadata_pages=document.source.docx_metadata_pages,
            failure=integrity_failure or document.oracle_validation_notes,
        )

    try:
        reference_snapshot = snapshot_pdf(reference)
    except ValueError as error:
        return DocumentBenchmarkResult(
            document_id=document.id,
            title=document.title,
            status="invalid_oracle",
            reference_pdf_pages=document.reference.pdf_pages,
            docx_metadata_pages=document.source.docx_metadata_pages,
            failure=str(error),
        )
    if reference_snapshot.page_count != document.reference.pdf_pages:
        return DocumentBenchmarkResult(
            document_id=document.id,
            title=document.title,
            status="invalid_oracle",
            reference_pdf_pages=reference_snapshot.page_count,
            docx_metadata_pages=document.source.docx_metadata_pages,
            failure="reference page count does not match the manifest",
        )

    strict_status, strict_error = _strict_probe(source, config=config)
    document_artifacts = artifacts_root / document.id
    document_artifacts.mkdir(parents=True, exist_ok=True)
    executions: list[ConversionExecution] = []
    first_result = None
    try:
        for run_number in range(1, config.repetitions + 1):
            destination = document_artifacts / f"candidate-{run_number}.pdf"
            started = time.perf_counter()
            conversion = Converter(_conversion_options(config, strict=False)).convert(
                source,
                destination,
            )
            duration_ms = round((time.perf_counter() - started) * 1000.0, 3)
            first_result = first_result or conversion
            if conversion.layout_json is not None:
                (document_artifacts / f"layout-{run_number}.json").write_text(
                    conversion.layout_json,
                    encoding="utf-8",
                )
            (document_artifacts / f"diagnostics-{run_number}.json").write_text(
                conversion.diagnostics.model_dump_json(indent=2),
                encoding="utf-8",
            )
            executions.append(
                ConversionExecution(
                    number=run_number,
                    pdf_path=destination,
                    pdf_sha256=conversion.pdf_sha256 or _sha256(destination),
                    layout_sha256=conversion.layout_sha256,
                    page_count=conversion.page_count,
                    duration_ms=duration_ms,
                )
            )
    except DocxPdfError as error:
        return DocumentBenchmarkResult(
            document_id=document.id,
            title=document.title,
            status="conversion_failed",
            reference_pdf_pages=reference_snapshot.page_count,
            docx_metadata_pages=document.source.docx_metadata_pages,
            strict_probe_status=strict_status,
            strict_probe_error=strict_error,
            executions=tuple(executions),
            failure=f"{type(error).__name__}: {error}",
        )

    first_candidate = executions[0].pdf_path
    comparison = compare_pdfs(reference, first_candidate)
    deterministic = _executions_are_deterministic(executions)
    reference_rendered_pages = _rendered_page_count(
        config.output_root,
        document.id,
        kind="reference",
    )
    candidate_rendered_pages = _rendered_page_count(
        config.output_root,
        document.id,
        kind="candidate",
    )
    return DocumentBenchmarkResult(
        document_id=document.id,
        title=document.title,
        status=comparison.verdict,
        reference_pdf_pages=reference_snapshot.page_count,
        candidate_pdf_pages=executions[0].page_count,
        reference_rendered_pages=reference_rendered_pages,
        candidate_rendered_pages=candidate_rendered_pages,
        docx_metadata_pages=document.source.docx_metadata_pages,
        strict_probe_status=strict_status,
        strict_probe_error=strict_error,
        deterministic=deterministic,
        executions=tuple(executions),
        warning_count=len(first_result.warnings) if first_result is not None else 0,
        font_substitution_count=(
            len(first_result.font_substitutions) if first_result is not None else 0
        ),
        comparison=comparison,
    )


def _strict_probe(source: Path, *, config: BenchmarkConfig) -> tuple[StrictProbeStatus, str | None]:
    try:
        Converter(_conversion_options(config, strict=True)).convert(source)
    except UnsupportedFeatureError as error:
        return "unsupported", f"{type(error).__name__}: {error}"
    except DocxPdfError as error:
        return "failed", f"{type(error).__name__}: {error}"
    return "passed", None


def _conversion_options(config: BenchmarkConfig, *, strict: bool) -> ConversionOptions:
    if config.font_path is None:
        fonts = FontConfiguration()
    else:
        substitutions: Mapping[str, str] = dict.fromkeys(
            _FONT_SUBSTITUTION_NAMES,
            config.font_alias,
        )
        fonts = FontConfiguration(
            registered_fonts={config.font_alias: config.font_path},
            substitutions=substitutions,
            default_font=config.font_alias,
            include_system_fonts=False,
        )
    return ConversionOptions(
        strict=strict,
        deterministic=True,
        font_configuration=fonts,
        allow_external_relationships=True,
    )


def _resolve_manifest_path(manifest_path: Path, relative_path: Path) -> Path:
    candidates = (
        manifest_path.parent / relative_path,
        manifest_path.parent.parent / relative_path,
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _integrity_failure(path: Path, expected_sha256: str, expected_size: int) -> str | None:
    if not path.is_file():
        return f"manifest file is missing: {path}"
    if path.stat().st_size != expected_size:
        return f"manifest file size differs: {path}"
    if _sha256(path) != expected_sha256:
        return f"manifest SHA-256 differs: {path}"
    return None


def _executions_are_deterministic(executions: list[ConversionExecution]) -> bool:
    return (
        len({execution.pdf_sha256 for execution in executions}) == 1
        and len({execution.layout_sha256 for execution in executions}) == 1
        and len({execution.page_count for execution in executions}) == 1
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _rendered_page_count(
    output_root: Path,
    document_id: str,
    *,
    kind: Literal["reference", "candidate"],
) -> int | None:
    directory = output_root / "rendered" / kind / document_id
    if not directory.is_dir():
        return None
    return len(tuple(directory.glob("page-*.png")))


def _write_outputs(report: BenchmarkReport, output_root: Path) -> None:
    (output_root / "results.json").write_text(
        report.model_dump_json(indent=2),
        encoding="utf-8",
    )
    _write_page_count_csv(report, output_root / "page_count_validation.csv")
    markdown = _markdown_report(report)
    (output_root / "report.md").write_text(markdown, encoding="utf-8")
    (output_root / "report.html").write_text(
        "<!doctype html><html lang=\"ja\"><meta charset=\"utf-8\">"
        "<title>docxpdf-native reference benchmark</title>"
        f"<body><pre>{html.escape(markdown)}</pre></body></html>\n",
        encoding="utf-8",
    )


def _write_page_count_csv(report: BenchmarkReport, destination: Path) -> None:
    with destination.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=(
                "document",
                "candidate",
                "mode",
                "docx_metadata_pages",
                "reference_pdf_pages",
                "reference_rendered_pages",
                "candidate_pdf_pages",
                "candidate_rendered_pages",
                "pdfinfo_pages",
                "manual_visible_pages",
                "status",
                "notes",
            ),
        )
        writer.writeheader()
        for document in report.documents:
            writer.writerow(
                {
                    "document": document.document_id,
                    "candidate": "docxpdf-native",
                    "mode": "font-controlled-lenient",
                    "docx_metadata_pages": document.docx_metadata_pages,
                    "reference_pdf_pages": document.reference_pdf_pages,
                    "reference_rendered_pages": document.reference_rendered_pages,
                    "candidate_pdf_pages": document.candidate_pdf_pages,
                    "candidate_rendered_pages": document.candidate_rendered_pages,
                    "pdfinfo_pages": "",
                    "manual_visible_pages": "",
                    "status": document.status,
                    "notes": document.failure or "; ".join(
                        document.comparison.reasons if document.comparison is not None else ()
                    ),
                }
            )


def _markdown_report(report: BenchmarkReport) -> str:
    matched = tuple(
        document
        for document in report.documents
        if document.status in {"exact_match", "layout_match"}
    )
    lines = [
        "# docxpdf-native 公式PDF比較",
        "",
        "## 結論",
        "",
        (
            "全テスト文書が構造上合致しました。"
            if len(matched) == len(report.documents)
            else f"構造上合致した文書は{len(matched)}/{len(report.documents)}件で、"
            "今回の公式PDF正解データには合致しませんでした。"
        ),
        "",
        "同一版と確認したDOCX/PDFペアを使い、各DOCXを3回変換して比較した結果です。",
        "PDF全体のSHA-256は候補PDF相互の再現性にだけ使用し、公式PDFとの一致条件にはしません。",
        "",
        "## ページ数",
        "",
        "| 文書 | DOCX記録値 | 正解PDF | 正解画像 | 候補PDF | 候補画像 | 判定 | 3回再現 |",
        "| --- | ---: | ---: | ---: | ---: | ---: | --- | --- |",
    ]
    for document in report.documents:
        metadata_pages = (
            str(document.docx_metadata_pages)
            if document.docx_metadata_pages is not None
            else "—"
        )
        candidate_pages = (
            str(document.candidate_pdf_pages)
            if document.candidate_pdf_pages is not None
            else "未生成"
        )
        reference_rendered = (
            str(document.reference_rendered_pages)
            if document.reference_rendered_pages is not None
            else "未計測"
        )
        candidate_rendered = (
            str(document.candidate_rendered_pages)
            if document.candidate_rendered_pages is not None
            else "未計測"
        )
        lines.append(
            f"| {document.title} | {metadata_pages} | {document.reference_pdf_pages} | "
            f"{reference_rendered} | {candidate_pages} | {candidate_rendered} | "
            f"{document.status} | {'一致' if document.deterministic else '不一致/未実行'} |"
        )
    lines.extend(
        [
            "",
            "## 改ページ・テキスト",
            "",
            "| 文書 | 改ページ一致数 | 改ページ一致率 | 平均ずれ | 最大ずれ | "
            "欠落 | 余分 | テキスト類似度 |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
    )
    for document in report.documents:
        metrics = document.comparison.metrics if document.comparison is not None else None
        if metrics is None:
            lines.append(f"| {document.title} | — | — | — | — | — | — | — |")
            continue
        lines.append(
            f"| {document.title} | {metrics.page_boundary_exact_matches} | "
            f"{metrics.page_boundary_match_rate:.3f} | "
            f"{metrics.mean_boundary_drift_chars:.1f}文字 | "
            f"{metrics.max_boundary_drift_chars}文字 | {metrics.missing_text_chars}文字 | "
            f"{metrics.duplicated_text_chars}文字 | {metrics.text_sequence_similarity:.6f} |"
        )
    lines.extend(
        [
            "",
            "## strict事前検査",
            "",
        ]
    )
    for document in report.documents:
        detail = document.strict_probe_error or "対応範囲内で変換成功"
        lines.append(f"- {document.title}: `{document.strict_probe_status}` — {detail}")
    lines.extend(
        [
            "",
            "## 判定上の注意",
            "",
            "- `mismatch` はページ数、ページ枠、回転、空ページ位置、全文、"
            "ページ別テキスト、改ページ境界のいずれかが違うことを示します。",
            "- 画像比較は同じPDFレンダラーで別途行い、この構造判定と分けて記録します。",
            "- strict事前検査の失敗とlenient候補PDFの結果は混同しません。",
            "",
        ]
    )
    return "\n".join(lines)


def main() -> int:
    """Run the checked-in corpus with the local Arial Unicode font when present."""

    root = Path(__file__).resolve().parent
    font = Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf")
    config = BenchmarkConfig(
        manifest_path=root / "data" / "reference-manifest.json",
        output_root=root / "results",
        font_path=font if font.is_file() else None,
    )
    report = run_benchmark(config)
    return 0 if all(document.status != "conversion_failed" for document in report.documents) else 1


if __name__ == "__main__":
    raise SystemExit(main())
