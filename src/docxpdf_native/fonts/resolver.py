from __future__ import annotations

import logging

from docxpdf_native.abstractions import FontResolver
from docxpdf_native.exceptions import FontNotFoundError
from docxpdf_native.fonts.bundled import metric_compatible_substitution
from docxpdf_native.fonts.cjk import (
    FontCategory,
    contains_cjk_characters,
    detect_system_cjk_fonts,
    known_japanese_font_category,
)
from docxpdf_native.fonts.registry import FontRecord, FontRegistry
from docxpdf_native.models.fonts import (
    FontConfiguration,
    FontSubstitution,
    ResolvedFont,
)

logger = logging.getLogger(__name__)


class DefaultFontResolver(FontResolver):
    """Resolve configured and installed font faces with explicit fallback policy.

    Resolution order for a requested font name:

    1. An exact match in the registry (explicit registration, configured
       directories, environment variable, scanned system fonts, then the
       fonts bundled with the package).
    2. A user-configured explicit substitution
       (:attr:`FontConfiguration.substitutions`).
    3. The built-in Japanese font substitution map: well-known but almost
       never installed names such as "MS 明朝"/"MS Gothic" are mapped to a
       detected system CJK font of the matching serif/sans category. This
       applies even in strict mode -- it is a substitution, not a silent
       failure, and is recorded like any other.
    4. The built-in metric-compatible Latin substitution map: Word's default
       fonts Calibri and Cambria (including common styled variants) are
       mapped to the bundled, metrically compatible Carlito and Caladea so
       line wrapping stays close to the original. Also applies in strict
       mode and is recorded (``reason="builtin-metric-compatible"``).
    5. Strict mode stops here with :class:`FontNotFoundError`.
    6. Lenient mode falls back to ``default_font`` if configured, otherwise a
       detected system CJK font when the request looks like it wants CJK
       glyphs (``east_asia=True``, a known Japanese font name, or CJK
       characters in the name itself), otherwise ``Helvetica``.
    """

    def __init__(self, configuration: FontConfiguration, *, strict: bool = True) -> None:
        self._configuration = configuration
        self._strict = strict
        self._registry = FontRegistry(
            registered_fonts=configuration.registered_fonts,
            font_directories=configuration.font_directories,
            environment_variable=configuration.environment_variable,
            include_system_fonts=configuration.include_system_fonts,
        )
        self._substitutions: list[FontSubstitution] = []
        self._system_cjk_fonts = detect_system_cjk_fonts(self._registry)

    @property
    def substitutions(self) -> tuple[FontSubstitution, ...]:
        return tuple(self._substitutions)

    def resolve(
        self,
        font_name: str,
        *,
        east_asia: bool = False,
        paragraph_index: int | None = None,
        run_index: int | None = None,
    ) -> ResolvedFont:
        requested = " ".join(font_name.split())
        if not requested:
            raise FontNotFoundError.for_font(
                font_name,
                paragraph_index=paragraph_index,
                run_index=run_index,
            )

        exact = self._registry.resolve(requested)
        if exact is not None:
            return self._resolved(exact, substituted=False)

        configured = self._substitution_target(requested)
        if configured is not None:
            return self._resolve_substitution(
                requested,
                configured,
                reason="configured",
                east_asia=east_asia,
                paragraph_index=paragraph_index,
                run_index=run_index,
            )

        builtin_category = known_japanese_font_category(requested)
        if builtin_category is not None:
            builtin_record = self._system_cjk_fonts.get(builtin_category)
            if builtin_record is not None:
                return self._resolve_with_record(
                    requested,
                    builtin_record,
                    reason="builtin-cjk",
                    east_asia=east_asia,
                    paragraph_index=paragraph_index,
                    run_index=run_index,
                )

        metric_compatible = metric_compatible_substitution(requested)
        if metric_compatible is not None:
            metric_record = self._registry.resolve(metric_compatible)
            if metric_record is not None:
                return self._resolve_with_record(
                    requested,
                    metric_record,
                    reason="builtin-metric-compatible",
                    east_asia=east_asia,
                    paragraph_index=paragraph_index,
                    run_index=run_index,
                )

        if self._strict:
            raise FontNotFoundError.for_font(
                requested,
                paragraph_index=paragraph_index,
                run_index=run_index,
            )

        if self._configuration.default_font:
            return self._resolve_substitution(
                requested,
                self._configuration.default_font,
                reason="lenient-default",
                east_asia=east_asia,
                paragraph_index=paragraph_index,
                run_index=run_index,
            )

        if east_asia or builtin_category is not None or contains_cjk_characters(requested):
            cjk_record = self._best_system_cjk_font(builtin_category)
            if cjk_record is not None:
                return self._resolve_with_record(
                    requested,
                    cjk_record,
                    reason="lenient-default",
                    east_asia=east_asia,
                    paragraph_index=paragraph_index,
                    run_index=run_index,
                )
            logger.warning(
                "No system CJK font could be detected; falling back to Helvetica for "
                "%r, which has no Japanese/CJK glyphs and will render as tofu ('□').",
                requested,
            )

        return self._resolve_substitution(
            requested,
            "Helvetica",
            reason="lenient-default",
            east_asia=east_asia,
            paragraph_index=paragraph_index,
            run_index=run_index,
        )

    def _best_system_cjk_font(self, category: FontCategory | None) -> FontRecord | None:
        if category is not None:
            record = self._system_cjk_fonts.get(category)
            if record is not None:
                return record
        return self._system_cjk_fonts.get("sans") or self._system_cjk_fonts.get("serif")

    def _resolve_substitution(
        self,
        requested: str,
        selected_name: str,
        *,
        reason: str,
        east_asia: bool,
        paragraph_index: int | None,
        run_index: int | None,
    ) -> ResolvedFont:
        selected = self._registry.resolve(selected_name)
        if selected is None:
            raise FontNotFoundError.for_font(
                selected_name,
                paragraph_index=paragraph_index,
                run_index=run_index,
            )
        return self._resolve_with_record(
            requested,
            selected,
            reason=reason,
            east_asia=east_asia,
            paragraph_index=paragraph_index,
            run_index=run_index,
        )

    def _resolve_with_record(
        self,
        requested: str,
        record: FontRecord,
        *,
        reason: str,
        east_asia: bool,
        paragraph_index: int | None,
        run_index: int | None,
    ) -> ResolvedFont:
        self._substitutions.append(
            FontSubstitution(
                requested_font=requested,
                selected_font=record.family,
                reason=reason,
                paragraph_index=paragraph_index,
                run_index=run_index,
                east_asia=east_asia,
            )
        )
        return self._resolved(record, substituted=True)

    def _substitution_target(self, requested: str) -> str | None:
        normalized = requested.casefold()
        for source, target in self._configuration.substitutions.items():
            if " ".join(source.split()).casefold() == normalized:
                return target
        return None

    @staticmethod
    def _resolved(record: FontRecord, *, substituted: bool) -> ResolvedFont:
        return ResolvedFont(
            family=record.family,
            path=record.path,
            postscript_name=record.postscript_name,
            source=record.source,
            substituted=substituted,
            font_number=record.font_number,
        )
