from __future__ import annotations

import hashlib
import json
import sys
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from reportlab.pdfgen.canvas import Canvas

sys.path.insert(0, str(Path(__file__).parents[2]))

from benchmark.validate_reference import BenchmarkConfig, run_benchmark


def _write_minimal_docx(path: Path, text: str) -> None:
    parts = {
        "[Content_Types].xml": b"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>""",
        "_rels/.rels": b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>""",
        "word/document.xml": f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>{text}</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>""".encode(),
        "word/styles.xml": b"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Helvetica" w:hAnsi="Helvetica" w:eastAsia="Helvetica"/>
    <w:sz w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>""",
    }
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])


def _write_reference_pdf(path: Path, text: str) -> None:
    output = BytesIO()
    canvas = Canvas(output, pagesize=(595.3, 841.9), invariant=1)
    canvas.setFont("Helvetica", 11)
    canvas.drawString(72, 760, text)
    canvas.showPage()
    canvas.save()
    path.write_bytes(output.getvalue())


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_benchmark_converts_three_times_and_compares_reference(tmp_path: Path) -> None:
    source = tmp_path / "fixture.docx"
    reference = tmp_path / "fixture.pdf"
    manifest = tmp_path / "manifest.json"
    _write_minimal_docx(source, "reference benchmark")
    _write_reference_pdf(reference, "reference benchmark")
    manifest.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "generated_at": "2026-07-15",
                "documents": [
                    {
                        "id": "fixture",
                        "title": "Fixture",
                        "revision": "test",
                        "publication_date": "2026-07-15",
                        "listing_url": "https://example.invalid/fixture",
                        "oracle_valid": True,
                        "oracle_validation_notes": "synthetic fixture",
                        "same_revision_evidence": ["same deterministic text"],
                        "source": {
                            "path": source.name,
                            "url": "https://example.invalid/fixture.docx",
                            "byte_size": source.stat().st_size,
                            "sha256": _sha256(source),
                            "http_last_modified": None,
                            "docx_metadata_pages": None,
                            "opening_text": "reference benchmark",
                            "closing_text": "reference benchmark",
                        },
                        "reference": {
                            "path": reference.name,
                            "url": "https://example.invalid/fixture.pdf",
                            "byte_size": reference.stat().st_size,
                            "sha256": _sha256(reference),
                            "http_last_modified": None,
                            "pdf_pages": 1,
                            "opening_text": "reference benchmark",
                            "closing_text": "reference benchmark",
                        },
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    reference_render = tmp_path / "results" / "rendered" / "reference" / "fixture"
    candidate_render = tmp_path / "results" / "rendered" / "candidate" / "fixture"
    reference_render.mkdir(parents=True)
    candidate_render.mkdir(parents=True)
    (reference_render / "page-1.png").write_bytes(b"fixture")
    (candidate_render / "page-1.png").write_bytes(b"fixture")

    report = run_benchmark(
        BenchmarkConfig(
            manifest_path=manifest,
            output_root=tmp_path / "results",
            repetitions=3,
        )
    )

    assert report.documents[0].candidate_pdf_pages == 1
    assert report.documents[0].reference_pdf_pages == 1
    assert report.documents[0].reference_rendered_pages == 1
    assert report.documents[0].candidate_rendered_pages == 1
    assert report.documents[0].deterministic is True
    assert len(tuple((tmp_path / "results" / "artifacts" / "fixture").glob("candidate-*.pdf"))) == 3
    markdown_report = (tmp_path / "results" / "report.md").read_text(encoding="utf-8")
    assert "改ページ一致率" in markdown_report
    assert "テキスト類似度" in markdown_report
