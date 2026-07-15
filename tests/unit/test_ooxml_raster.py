from __future__ import annotations

import sys
from io import BytesIO

import pytest

from docxpdf_native.ooxml.raster import RasterImageConverter

PIL = pytest.importorskip("PIL")
from PIL import Image  # noqa: E402


def _encode(image: Image.Image, *, image_format: str) -> bytes:
    buffer = BytesIO()
    image.save(buffer, format=image_format)
    return buffer.getvalue()


def test_is_available_reflects_whether_pillow_can_be_imported() -> None:
    assert RasterImageConverter.is_available() is True


def test_is_available_returns_false_when_pillow_is_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setitem(sys.modules, "PIL", None)

    assert RasterImageConverter.is_available() is False


def test_convert_to_png_returns_none_when_pillow_is_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setitem(sys.modules, "PIL", None)

    assert RasterImageConverter.convert_to_png(b"anything") is None


def test_convert_to_png_decodes_rgb_gif() -> None:
    data = _encode(Image.new("RGB", (3, 2), color=(10, 20, 30)), image_format="GIF")

    converted = RasterImageConverter.convert_to_png(data)

    assert converted is not None
    assert converted[:8] == b"\x89PNG\r\n\x1a\n"
    with Image.open(BytesIO(converted)) as decoded:
        assert decoded.format == "PNG"
        assert decoded.size == (3, 2)


def test_convert_to_png_normalizes_palette_mode_bmp() -> None:
    palette_image = Image.new("P", (4, 4))
    palette_image.putpalette([0, 0, 0, 255, 255, 255] + [0] * (256 * 3 - 6))
    data = _encode(palette_image, image_format="BMP")

    converted = RasterImageConverter.convert_to_png(data)

    assert converted is not None
    with Image.open(BytesIO(converted)) as decoded:
        assert decoded.mode in {"RGBA", "RGB"}


def test_convert_to_png_returns_none_for_undecodable_bytes() -> None:
    assert RasterImageConverter.convert_to_png(b"not an image") is None
