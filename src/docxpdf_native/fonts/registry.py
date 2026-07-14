from __future__ import annotations

import logging
import os
import platform
from collections.abc import Mapping, Sequence
from pathlib import Path

from fontTools.ttLib import TTFont, TTLibError
from pydantic import BaseModel, ConfigDict

logger = logging.getLogger(__name__)


class FontRecord(BaseModel):
    """A concrete font face known to the resolver."""

    model_config = ConfigDict(frozen=True)

    family: str
    path: Path | None = None
    postscript_name: str | None = None
    source: str
    aliases: tuple[str, ...] = ()


class FontRegistry:
    """Deterministic, lazily scanned font index with explicit priority groups."""

    _FONT_SUFFIXES = frozenset({".ttf", ".otf"})
    _STANDARD_FONTS = (
        "Courier",
        "Courier-Bold",
        "Courier-Oblique",
        "Courier-BoldOblique",
        "Helvetica",
        "Helvetica-Bold",
        "Helvetica-Oblique",
        "Helvetica-BoldOblique",
        "Times-Roman",
        "Times-Bold",
        "Times-Italic",
        "Times-BoldItalic",
        "Symbol",
        "ZapfDingbats",
    )

    def __init__(
        self,
        *,
        registered_fonts: Mapping[str, Path] | None = None,
        font_directories: Sequence[Path] = (),
        environment_variable: str = "DOCXPDF_NATIVE_FONT_DIRS",
        include_system_fonts: bool = True,
    ) -> None:
        self._explicit: dict[str, FontRecord] = {}
        self._directory_groups: list[tuple[str, tuple[Path, ...]]] = []
        self._scanned_groups: list[dict[str, FontRecord]] = []

        for alias, path in (registered_fonts or {}).items():
            record = self._font_record(path, source="registered")
            if record is None:
                raise ValueError(f"registered font is missing or invalid: {alias!r} at {path}")
            self._explicit[self._normalize(alias)] = record
            self._index_record(self._explicit, record, overwrite=False)

        configured = self._existing_directories(font_directories)
        if configured:
            self._directory_groups.append(("font-directory", configured))

        environment = os.environ.get(environment_variable, "")
        environment_directories = self._existing_directories(
            tuple(Path(item).expanduser() for item in environment.split(os.pathsep) if item)
        )
        if environment_directories:
            self._directory_groups.append(("environment", environment_directories))

        if include_system_fonts:
            system = self._existing_directories(self._system_font_directories())
            if system:
                self._directory_groups.append(("system", system))

    def resolve(self, family: str) -> FontRecord | None:
        normalized = self._normalize(family)
        if not normalized:
            return None
        explicit = self._explicit.get(normalized)
        if explicit is not None:
            return explicit

        for index, (source, directories) in enumerate(self._directory_groups):
            if index == len(self._scanned_groups):
                self._scanned_groups.append(self._scan(directories, source=source))
            record = self._scanned_groups[index].get(normalized)
            if record is not None:
                return record
        return self._standard_font(normalized)

    @staticmethod
    def _normalize(name: str) -> str:
        return " ".join(name.split()).casefold()

    @classmethod
    def _standard_font(cls, normalized: str) -> FontRecord | None:
        for family in cls._STANDARD_FONTS:
            if cls._normalize(family) == normalized:
                return FontRecord(
                    family=family,
                    postscript_name=family,
                    source="reportlab-standard",
                    aliases=(family,),
                )
        return None

    @staticmethod
    def _existing_directories(directories: Sequence[Path]) -> tuple[Path, ...]:
        result: list[Path] = []
        seen: set[Path] = set()
        for directory in directories:
            resolved = directory.expanduser().resolve()
            if resolved in seen or not resolved.is_dir():
                continue
            result.append(resolved)
            seen.add(resolved)
        return tuple(result)

    @staticmethod
    def _system_font_directories() -> tuple[Path, ...]:
        system = platform.system()
        home = Path.home()
        if system == "Darwin":
            return (
                home / "Library/Fonts",
                Path("/Library/Fonts"),
                Path("/System/Library/Fonts"),
            )
        if system == "Windows":
            windows = Path(os.environ.get("WINDIR", "C:/Windows"))
            return (windows / "Fonts",)
        return (
            home / ".fonts",
            home / ".local/share/fonts",
            Path("/usr/local/share/fonts"),
            Path("/usr/share/fonts"),
        )

    def _scan(self, directories: Sequence[Path], *, source: str) -> dict[str, FontRecord]:
        index: dict[str, FontRecord] = {}
        paths = sorted(
            (
                path
                for directory in directories
                for path in directory.rglob("*")
                if path.is_file() and path.suffix.lower() in self._FONT_SUFFIXES
            ),
            key=lambda item: str(item).casefold(),
        )
        for path in paths:
            record = self._font_record(path, source=source)
            if record is None:
                logger.debug("Skipping unreadable font file: %s", path)
                continue
            self._index_record(index, record, overwrite=False)
        return index

    @classmethod
    def _index_record(
        cls, index: dict[str, FontRecord], record: FontRecord, *, overwrite: bool
    ) -> None:
        for name in (record.family, record.postscript_name, *record.aliases):
            if not name:
                continue
            normalized = cls._normalize(name)
            if overwrite or normalized not in index:
                index[normalized] = record

    @staticmethod
    def _font_record(path: Path, *, source: str) -> FontRecord | None:
        resolved = path.expanduser().resolve()
        if not resolved.is_file() or resolved.suffix.lower() not in FontRegistry._FONT_SUFFIXES:
            return None
        try:
            with TTFont(resolved, lazy=True) as font:
                names = font["name"].names
                decoded: dict[int, list[str]] = {}
                for name in names:
                    if name.nameID not in {1, 4, 6, 16}:
                        continue
                    try:
                        value = name.toUnicode().strip()
                    except UnicodeError:
                        continue
                    if value and value not in decoded.setdefault(name.nameID, []):
                        decoded[name.nameID].append(value)
                family_values = decoded.get(16) or decoded.get(1) or decoded.get(4)
                if not family_values:
                    return None
                postscript_values = decoded.get(6, [])
                aliases = tuple(
                    dict.fromkeys(
                        value
                        for identifier in (1, 4, 6, 16)
                        for value in decoded.get(identifier, [])
                    )
                )
                return FontRecord(
                    family=family_values[0],
                    path=resolved,
                    postscript_name=postscript_values[0] if postscript_values else None,
                    source=source,
                    aliases=aliases,
                )
        except (OSError, KeyError, TTLibError):
            return None
