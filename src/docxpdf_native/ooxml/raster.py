"""Optional Pillow-based conversion of legacy raster image formats to PNG.

Pillow is an optional dependency (the ``images`` extra).  When it is not
installed, callers fall back to a same-size placeholder instead of raising,
so that documents built without the extra still convert end to end.
"""

from __future__ import annotations

from io import BytesIO


class RasterImageConverter:
    """Convert GIF/BMP/TIFF (and other Pillow-readable) bytes to PNG."""

    @staticmethod
    def is_available() -> bool:
        """Return whether Pillow is importable in the current environment."""

        try:
            import PIL  # noqa: F401
        except ImportError:
            return False
        return True

    @staticmethod
    def convert_to_png(data: bytes) -> bytes | None:
        """Return PNG bytes decoded from ``data``, or ``None`` on any failure.

        Every failure mode collapses to ``None`` on purpose: a corrupt or
        unrecognized embedded image must never abort the whole conversion,
        it should fall back to a dimension-preserving placeholder instead.
        """

        try:
            from PIL import Image
        except ImportError:
            return None
        try:
            with Image.open(BytesIO(data)) as opened:
                opened.load()
                normalized = (
                    opened.convert("RGBA") if opened.mode not in {"RGB", "RGBA", "L"} else opened
                )
                buffer = BytesIO()
                normalized.save(buffer, format="PNG")
                return buffer.getvalue()
        except Exception:
            return None
