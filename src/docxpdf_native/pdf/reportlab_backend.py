from __future__ import annotations

from collections.abc import Mapping
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas

from docxpdf_native.abstractions import PdfBackend
from docxpdf_native.exceptions import PdfGenerationError
from docxpdf_native.models.document import BorderModel
from docxpdf_native.models.layout import (
    CellBox,
    ImageBox,
    LayoutDocument,
    PageModel,
    ParagraphBox,
    TableBox,
)


class ReportLabPdfBackend(PdfBackend):
    """Draw a completed top-left-origin layout with ReportLab primitives."""

    def __init__(
        self,
        *,
        deterministic: bool = False,
        font_paths: Mapping[str, Path] | None = None,
    ) -> None:
        self._deterministic = deterministic
        self._font_paths = tuple(sorted((font_paths or {}).items()))

    def render(
        self,
        document: LayoutDocument,
        destination: Path | BinaryIO | None = None,
    ) -> bytes:
        try:
            self._register_fonts()
            output = BytesIO()
            pdf = Canvas(output, invariant=int(self._deterministic))
            for page in document.pages:
                self._draw_page(pdf, page)
                pdf.showPage()
            pdf.save()
            payload = output.getvalue()
            if isinstance(destination, Path):
                destination.write_bytes(payload)
            elif destination is not None:
                destination.write(payload)
            return payload
        except PdfGenerationError:
            raise
        except Exception as exc:
            raise PdfGenerationError("PDF generation failed", cause=exc) from exc

    def _register_fonts(self) -> None:
        for font_name, font_path in self._font_paths:
            pdfmetrics.registerFont(TTFont(font_name, str(font_path)))

    @staticmethod
    def _draw_page(pdf: Canvas, page: PageModel) -> None:
        pdf.setPageSize((page.width, page.height))
        for block in (*page.header, *page.body, *page.footer):
            if isinstance(block, ImageBox):
                ReportLabPdfBackend._draw_image(pdf, page, block)
                continue
            if isinstance(block, TableBox):
                for cell in block.cells:
                    if cell.background_color is not None:
                        pdf.setFillColor(HexColor(f"#{cell.background_color.lstrip('#')}"))
                        pdf.rect(
                            cell.x,
                            page.height - cell.y - cell.height,
                            cell.width,
                            cell.height,
                            stroke=0,
                            fill=1,
                        )
                    for cell_block in cell.blocks:
                        if isinstance(cell_block, ParagraphBox):
                            ReportLabPdfBackend._draw_paragraph(pdf, page, cell_block)
                        elif isinstance(cell_block, ImageBox):
                            ReportLabPdfBackend._draw_image(pdf, page, cell_block)
                    ReportLabPdfBackend._draw_cell_borders(pdf, page, cell)
                ReportLabPdfBackend._draw_table_borders(pdf, page, block)
                continue
            if not isinstance(block, ParagraphBox):
                continue
            ReportLabPdfBackend._draw_paragraph(pdf, page, block)

    @staticmethod
    def _draw_paragraph(pdf: Canvas, page: PageModel, paragraph: ParagraphBox) -> None:
        for line in paragraph.lines:
            baseline_y = page.height - line.y - line.baseline
            for fragment in line.fragments:
                if isinstance(fragment, ImageBox):
                    ReportLabPdfBackend._draw_image(pdf, page, fragment)
                    continue
                text_baseline = baseline_y + fragment.baseline_shift
                text_color = HexColor(f"#{fragment.color.lstrip('#')}")
                if fragment.highlight is not None:
                    pdf.setFillColor(HexColor(f"#{fragment.highlight.lstrip('#')}"))
                    pdf.rect(
                        fragment.x,
                        page.height - fragment.y - fragment.height,
                        fragment.width,
                        fragment.height,
                        stroke=0,
                        fill=1,
                    )
                pdf.setFillColor(text_color)
                font_name = ReportLabPdfBackend._standard_font_face(
                    fragment.font_name,
                    bold=fragment.bold,
                    italic=fragment.italic,
                )
                if fragment.character_spacing:
                    text_object = pdf.beginText(fragment.x, text_baseline)
                    text_object.setFont(font_name, fragment.font_size)
                    text_object.setCharSpace(fragment.character_spacing)
                    text_object.textOut(fragment.text)
                    pdf.drawText(text_object)
                else:
                    pdf.setFont(font_name, fragment.font_size)
                    pdf.drawString(
                        fragment.x,
                        text_baseline,
                        fragment.text,
                    )
                if fragment.underline:
                    underline_y = text_baseline - max(0.5, fragment.font_size * 0.1)
                    pdf.setStrokeColor(text_color)
                    pdf.setLineWidth(max(0.5, fragment.font_size / 16))
                    pdf.line(fragment.x, underline_y, fragment.x + fragment.width, underline_y)
                if fragment.strike:
                    strike_y = text_baseline + fragment.font_size * 0.3
                    pdf.setStrokeColor(text_color)
                    pdf.setLineWidth(max(0.5, fragment.font_size / 16))
                    pdf.line(fragment.x, strike_y, fragment.x + fragment.width, strike_y)

    @staticmethod
    def _draw_image(pdf: Canvas, page: PageModel, image: ImageBox) -> None:
        reader = ImageReader(BytesIO(image.image_data))
        pdf.drawImage(
            reader,
            image.x,
            page.height - image.y - image.height,
            width=image.width,
            height=image.height,
            mask="auto",
        )

    @staticmethod
    def _draw_cell_borders(pdf: Canvas, page: PageModel, cell: CellBox) -> None:
        top = page.height - cell.y
        bottom = top - cell.height
        left = cell.x
        right = left + cell.width
        ReportLabPdfBackend._draw_border(pdf, cell.borders.top, left, top, right, top)
        ReportLabPdfBackend._draw_border(pdf, cell.borders.right, right, top, right, bottom)
        ReportLabPdfBackend._draw_border(pdf, cell.borders.bottom, left, bottom, right, bottom)
        ReportLabPdfBackend._draw_border(pdf, cell.borders.left, left, top, left, bottom)

    @staticmethod
    def _draw_table_borders(pdf: Canvas, page: PageModel, table: TableBox) -> None:
        top = page.height - table.y
        bottom = top - table.height
        left = table.x
        right = left + table.width
        ReportLabPdfBackend._draw_border(pdf, table.borders.top, left, top, right, top)
        ReportLabPdfBackend._draw_border(pdf, table.borders.right, right, top, right, bottom)
        ReportLabPdfBackend._draw_border(pdf, table.borders.bottom, left, bottom, right, bottom)
        ReportLabPdfBackend._draw_border(pdf, table.borders.left, left, top, left, bottom)

        for cell in sorted(table.cells, key=lambda item: (item.row_index, item.column_index)):
            cell_top = page.height - cell.y
            cell_bottom = cell_top - cell.height
            cell_left = cell.x
            cell_right = cell_left + cell.width
            if cell_right < right - 1e-6:
                ReportLabPdfBackend._draw_border(
                    pdf,
                    table.borders.inside_vertical,
                    cell_right,
                    cell_top,
                    cell_right,
                    cell_bottom,
                )
            if cell_bottom > bottom + 1e-6:
                ReportLabPdfBackend._draw_border(
                    pdf,
                    table.borders.inside_horizontal,
                    cell_left,
                    cell_bottom,
                    cell_right,
                    cell_bottom,
                )

    @staticmethod
    def _draw_border(
        pdf: Canvas,
        border: BorderModel | None,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
    ) -> None:
        if border is None or border.width <= 0:
            return
        style = border.style.casefold()
        if style in {"nil", "none"}:
            return
        pdf.setStrokeColor(HexColor(f"#{border.color.lstrip('#')}"))
        if style == "double":
            pdf.setLineWidth(max(border.width / 3, 0.25))
            pdf.setDash()
            offset = border.width / 3
            if abs(y1 - y2) < 1e-6:
                pdf.line(x1, y1 - offset, x2, y2 - offset)
                pdf.line(x1, y1 + offset, x2, y2 + offset)
            else:
                pdf.line(x1 - offset, y1, x2 - offset, y2)
                pdf.line(x1 + offset, y1, x2 + offset, y2)
            return
        pdf.setLineWidth(border.width)
        if style in {"dashed", "dashsmallgap", "dashdot", "dashdotstroked"}:
            pdf.setDash(border.width * 3, border.width * 2)
        elif style in {"dotted", "dot"}:
            pdf.setDash(border.width, border.width * 2)
        else:
            pdf.setDash()
        pdf.line(x1, y1, x2, y2)

    @staticmethod
    def _standard_font_face(font_name: str, *, bold: bool, italic: bool) -> str:
        if not bold and not italic:
            return font_name
        if font_name == "Helvetica":
            if bold and italic:
                return "Helvetica-BoldOblique"
            return "Helvetica-Bold" if bold else "Helvetica-Oblique"
        if font_name == "Times-Roman":
            if bold and italic:
                return "Times-BoldItalic"
            return "Times-Bold" if bold else "Times-Italic"
        if font_name == "Courier":
            if bold and italic:
                return "Courier-BoldOblique"
            return "Courier-Bold" if bold else "Courier-Oblique"
        return font_name
