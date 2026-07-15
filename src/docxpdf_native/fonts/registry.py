from __future__ import annotations

import logging
import os
import platform
from collections.abc import Mapping, Sequence
from pathlib import Path

from fontTools.ttLib import TTCollection, TTFont, TTLibError
from pydantic import BaseModel, ConfigDict

from docxpdf_native.fonts.bundled import bundled_font_paths

logger = logging.getLogger(__name__)


class FontRecord(BaseModel):
    """A concrete font face known to the resolver."""

    model_config = ConfigDict(frozen=True)

    family: str
    path: Path | None = None
    postscript_name: str | None = None
    source: str
    aliases: tuple[str, ...] = ()
    font_number: int | None = None
    """Index into a TrueType/OpenType Collection (``.ttc``/``.otc``); ``None``
    for standalone font files, where reportlab and fontTools address the
    face by path alone."""
    bold: bool = False
    italic: bool = False
    """Style flags from the face's ``head.macStyle`` bits. Every face of a
    family carries the same family name, so these flags let indexing prefer
    the Regular face for the bare family name over whichever styled face a
    path-ordered scan happens to visit first."""

    @property
    def styled(self) -> bool:
        return self.bold or self.italic


class FontRegistry:
    """Deterministic, lazily scanned font index with explicit priority groups."""

    _SINGLE_FONT_SUFFIXES = frozenset({".ttf", ".otf"})
    _COLLECTION_SUFFIXES = frozenset({".ttc", ".otc"})
    _FONT_SUFFIXES = _SINGLE_FONT_SUFFIXES | _COLLECTION_SUFFIXES
    _NAME_TABLE_IDS = frozenset({1, 4, 6, 16})
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
        self._bundled_index: dict[str, FontRecord] | None = None

        for alias, path in (registered_fonts or {}).items():
            record = self._font_record_any(path, source="registered")
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
        bundled = self._bundled().get(normalized)
        if bundled is not None:
            return bundled
        return self._standard_font(normalized)

    def _bundled(self) -> dict[str, FontRecord]:
        """Lazily index the metric-compatible fonts shipped with the package.

        The bundled files are indexed in the deterministic order returned by
        :func:`bundled_font_paths` (Regular faces first) rather than by a
        directory scan, so the bare family name always resolves to the
        Regular face even though every face of a family shares the same
        family name.
        """
        if self._bundled_index is None:
            index: dict[str, FontRecord] = {}
            for path in bundled_font_paths():
                record = self._font_record(path, source="bundled")
                if record is None:
                    logger.debug("Skipping unreadable bundled font file: %s", path)
                    continue
                self._index_record(index, record, overwrite=False, prefer_regular=True)
            self._bundled_index = index
        return self._bundled_index

    @classmethod
    def resolve_face(cls, path: Path, family: str) -> FontRecord | None:
        """Recover the embeddable face matching ``family`` inside a font file.

        Callers such as the ReportLab backend and the fontTools text measurer
        only carry a flat ``family -> path`` mapping forward from resolution.
        For a TrueType/OpenType Collection that mapping alone is not enough to
        know *which* face inside the file to embed or measure, so this helper
        re-derives the collection index (``font_number``) deterministically
        from the file itself. For standalone ``.ttf``/``.otf`` files it simply
        confirms the family is present and returns ``font_number=None``.
        """
        resolved = Path(path).expanduser().resolve()
        normalized_family = cls._normalize(family)
        if resolved.suffix.lower() in cls._COLLECTION_SUFFIXES:
            candidates: tuple[FontRecord, ...] = cls._collection_records(resolved, source="embed")
        else:
            single = cls._font_record(resolved, source="embed")
            candidates = (single,) if single is not None else ()
        matches = tuple(
            record
            for record in candidates
            if any(
                name and cls._normalize(name) == normalized_family
                for name in (record.family, record.postscript_name, *record.aliases)
            )
        )
        if not matches:
            return None
        # A bare family name matches every face of the family; prefer the
        # Regular face, falling back to the first match in face order.
        return next((record for record in matches if not record.styled), matches[0])

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
            if path.suffix.lower() in self._COLLECTION_SUFFIXES:
                faces = self._collection_records(path, source=source)
                if not faces:
                    logger.debug("Skipping font collection with no embeddable faces: %s", path)
                for face in faces:
                    self._index_record(index, face, overwrite=False, prefer_regular=True)
                continue
            single = self._font_record(path, source=source)
            if single is None:
                logger.debug("Skipping unreadable font file: %s", path)
                continue
            self._index_record(index, single, overwrite=False, prefer_regular=True)
        return index

    @classmethod
    def _index_record(
        cls,
        index: dict[str, FontRecord],
        record: FontRecord,
        *,
        overwrite: bool,
        prefer_regular: bool = False,
    ) -> None:
        """Index a face under all of its names.

        With ``prefer_regular`` (used for scanned and bundled groups), a
        Regular face replaces a previously indexed styled face that shares
        the same name. Every face of a family carries the family name in its
        ``name`` table, and directory scans visit files in path order, so
        without this rule "Times New Roman" would resolve to whichever
        styled file sorts first ("Times New Roman Bold Italic.ttf" sorts
        before "Times New Roman.ttf") and body text would be measured and
        drawn with bold-italic metrics. First-wins order still applies
        between faces of equal styledness, keeping the index deterministic.
        """
        for name in (record.family, record.postscript_name, *record.aliases):
            if not name:
                continue
            normalized = cls._normalize(name)
            existing = index.get(normalized)
            regular_replaces_styled = (
                prefer_regular and existing is not None and existing.styled and not record.styled
            )
            if overwrite or existing is None or regular_replaces_styled:
                index[normalized] = record

    @classmethod
    def _font_record_any(cls, path: Path, *, source: str) -> FontRecord | None:
        """Resolve an explicitly registered font, allowing a collection path.

        ``registered_fonts`` maps an alias directly to a file; when that file
        is a ``.ttc``/``.otc`` we deterministically pick its first embeddable
        face (lowest ``font_number``) rather than rejecting the registration.
        """
        resolved = Path(path).expanduser().resolve()
        if resolved.suffix.lower() in cls._COLLECTION_SUFFIXES:
            records = cls._collection_records(resolved, source=source)
            return records[0] if records else None
        return cls._font_record(resolved, source=source)

    @classmethod
    def _read_names(cls, font: TTFont) -> tuple[str, str | None, tuple[str, ...]] | None:
        names = font["name"].names
        decoded: dict[int, list[str]] = {}
        for name in names:
            if name.nameID not in cls._NAME_TABLE_IDS:
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
                value for identifier in (1, 4, 6, 16) for value in decoded.get(identifier, [])
            )
        )
        return family_values[0], (postscript_values[0] if postscript_values else None), aliases

    @staticmethod
    def _style_flags(font: TTFont) -> tuple[bool, bool]:
        """Read the (bold, italic) flags from the face's ``head.macStyle``."""
        try:
            mac_style = int(font["head"].macStyle)
        except (KeyError, AttributeError, TypeError, ValueError):
            return False, False
        return bool(mac_style & 0b01), bool(mac_style & 0b10)

    @classmethod
    def _font_record(cls, path: Path, *, source: str) -> FontRecord | None:
        resolved = path.expanduser().resolve()
        if not resolved.is_file() or resolved.suffix.lower() not in cls._SINGLE_FONT_SUFFIXES:
            return None
        try:
            with TTFont(resolved, lazy=True) as font:
                names = cls._read_names(font)
                if names is None:
                    return None
                family, postscript_name, aliases = names
                bold, italic = cls._style_flags(font)
                return FontRecord(
                    family=family,
                    path=resolved,
                    postscript_name=postscript_name,
                    source=source,
                    aliases=aliases,
                    bold=bold,
                    italic=italic,
                )
        except (OSError, KeyError, TTLibError):
            return None

    @classmethod
    def _collection_records(cls, path: Path, *, source: str) -> tuple[FontRecord, ...]:
        """Index every embeddable face of a TrueType/OpenType Collection.

        ReportLab's TTFont backend can only embed TrueType outlines (a
        ``glyf`` table); OpenType/CFF-flavored faces -- common in TTC files
        such as macOS's Hiragino collections -- raise ``TTFError`` when
        registered. Those faces are skipped here so that anything the
        registry resolves is guaranteed embeddable, per a face's own
        deterministic index within the file (ascending ``font_number``).
        """
        resolved = path.expanduser().resolve()
        if not resolved.is_file() or resolved.suffix.lower() not in cls._COLLECTION_SUFFIXES:
            return ()
        records: list[FontRecord] = []
        try:
            with TTCollection(str(resolved), lazy=True) as collection:
                for font_number, font in enumerate(collection.fonts):
                    if "glyf" not in font:
                        logger.debug(
                            "Skipping non-TrueType face %d in collection (no glyf table): %s",
                            font_number,
                            resolved,
                        )
                        continue
                    try:
                        names = cls._read_names(font)
                    except (KeyError, UnicodeError):
                        names = None
                    if names is None:
                        continue
                    family, postscript_name, aliases = names
                    bold, italic = cls._style_flags(font)
                    records.append(
                        FontRecord(
                            family=family,
                            path=resolved,
                            postscript_name=postscript_name,
                            source=source,
                            aliases=aliases,
                            font_number=font_number,
                            bold=bold,
                            italic=italic,
                        )
                    )
        except (OSError, KeyError, TTLibError):
            return ()
        return tuple(records)
