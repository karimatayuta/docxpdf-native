"""Public conversion pipeline."""

from __future__ import annotations

import hashlib
import logging
import os
import tempfile
from pathlib import Path
from typing import BinaryIO

from docxpdf_native.diagnostics import DiagnosticsExporter
from docxpdf_native.exceptions import InvalidDocxError, PdfGenerationError
from docxpdf_native.fonts.metrics import FontToolsTextMeasurer
from docxpdf_native.fonts.resolver import DefaultFontResolver
from docxpdf_native.layout.engine import NativeLayoutEngine
from docxpdf_native.models import (
    ConversionOptions,
    ConversionResult,
    DiagnosticInfo,
    LayoutDocument,
)
from docxpdf_native.ooxml.parser import DocumentSource, OoxmlDocumentParser
from docxpdf_native.pdf import ReportLabPdfBackend
from docxpdf_native.resolver import NativeStyleResolver
from docxpdf_native.security.paths import OutputPathGuard

logger = logging.getLogger(__name__)

Destination = Path | str | BinaryIO


class Converter:
    """Convert DOCX sources through parse, resolve, layout, and PDF stages."""

    def __init__(self, options: ConversionOptions | None = None) -> None:
        self.options = options or ConversionOptions()
        self._last_layout: LayoutDocument | None = None

    @property
    def last_layout(self) -> LayoutDocument | None:
        """Return the immutable layout from the latest successful conversion."""
        return self._last_layout

    @property
    def layout_document(self) -> LayoutDocument | None:
        """Compatibility alias used by the CLI diagnostics layer."""
        return self._last_layout

    def convert(
        self,
        source: DocumentSource,
        destination: Destination | None = None,
    ) -> ConversionResult:
        """Convert a path, byte string, or binary stream to PDF."""
        output_path = self._validate_destination(source, destination)
        parser = OoxmlDocumentParser(limits=self.options.resource_limits)
        document = parser.parse(source, options=self.options)

        font_resolver = DefaultFontResolver(
            self.options.font_configuration,
            strict=self.options.strict,
        )
        style_resolver = NativeStyleResolver(font_resolver)
        resolved_document = style_resolver.resolve(document).model_copy(
            update={"font_substitutions": font_resolver.substitutions}
        )

        text_measurer = FontToolsTextMeasurer()
        layout = NativeLayoutEngine(text_measurer).layout(
            resolved_document,
            options=self.options,
        )
        backend = ReportLabPdfBackend(
            deterministic=self.options.deterministic,
            font_paths=style_resolver.font_paths,
        )
        pdf_bytes = backend.render(layout)
        self._write_destination(pdf_bytes, destination, output_path=output_path)

        layout_json = DiagnosticsExporter.to_json(layout)
        warnings = (*parser.warnings, *layout.warnings)
        diagnostics = DiagnosticInfo(
            warnings=warnings,
            unsupported_features=parser.unsupported_features,
            font_substitutions=font_resolver.substitutions,
            counters={
                "pages": layout.page_count,
                "warnings": len(warnings),
                "font_substitutions": len(font_resolver.substitutions),
            },
        )
        result = ConversionResult(
            page_count=layout.page_count,
            warnings=warnings,
            unsupported_features=parser.unsupported_features,
            font_substitutions=font_resolver.substitutions,
            destination=output_path,
            pdf_bytes=pdf_bytes,
            pdf_sha256=hashlib.sha256(pdf_bytes).hexdigest(),
            layout_sha256=hashlib.sha256(layout_json.encode("utf-8")).hexdigest(),
            layout_json=layout_json,
            diagnostics=diagnostics,
        )
        self._last_layout = layout
        logger.info("Converted DOCX to %d PDF page(s)", result.page_count)
        return result

    def convert_bytes(self, source: bytes) -> bytes:
        """Convert DOCX bytes and return PDF bytes."""
        result = self.convert(source)
        if result.pdf_bytes is None:
            raise PdfGenerationError("PDF backend returned no bytes")
        return result.pdf_bytes

    def convert_stream(self, source: BinaryIO, destination: BinaryIO) -> ConversionResult:
        """Convert from one binary stream into another."""
        return self.convert(source, destination)

    @staticmethod
    def _validate_destination(
        source: DocumentSource,
        destination: Destination | None,
    ) -> Path | None:
        if not isinstance(destination, (Path, str)):
            return None
        path = Path(destination).expanduser()
        parent = path.parent
        if not parent.exists() or not parent.is_dir():
            raise PdfGenerationError(f"output parent directory does not exist: {parent}")
        if path.is_symlink():
            raise PdfGenerationError(f"refusing to overwrite symbolic link: {path}")
        symlinked_parent = OutputPathGuard.symlinked_parent(path)
        if symlinked_parent is not None:
            raise PdfGenerationError(
                f"refusing output through symbolic link directory: {symlinked_parent}"
            )
        if isinstance(source, (Path, str)):
            source_path = Path(source).expanduser()
            same_existing_file = (
                source_path.exists() and path.exists() and source_path.samefile(path)
            )
            if same_existing_file or source_path.resolve() == path.resolve():
                raise InvalidDocxError("input and output must be different files")
        return path

    @staticmethod
    def _write_destination(
        payload: bytes,
        destination: Destination | None,
        *,
        output_path: Path | None,
    ) -> None:
        if destination is None:
            return
        if output_path is None:
            try:
                destination.write(payload)  # type: ignore[union-attr]
            except (OSError, TypeError, ValueError) as error:
                raise PdfGenerationError("could not write PDF stream", cause=error) from error
            return
        temporary_name: str | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                prefix=f".{output_path.name}.",
                suffix=".tmp",
                dir=output_path.parent,
                delete=False,
            ) as temporary:
                temporary.write(payload)
                temporary.flush()
                os.fsync(temporary.fileno())
                temporary_name = temporary.name
            os.replace(temporary_name, output_path)
        except OSError as error:
            if temporary_name is not None:
                try:
                    Path(temporary_name).unlink(missing_ok=True)
                except OSError:
                    logger.warning("Could not remove temporary PDF: %s", temporary_name)
            raise PdfGenerationError(
                f"could not write PDF output: {output_path}",
                cause=error,
            ) from error
