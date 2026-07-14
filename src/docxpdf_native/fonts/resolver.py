from __future__ import annotations

from docxpdf_native.abstractions import FontResolver
from docxpdf_native.exceptions import FontNotFoundError
from docxpdf_native.fonts.registry import FontRecord, FontRegistry
from docxpdf_native.models.fonts import (
    FontConfiguration,
    FontSubstitution,
    ResolvedFont,
)


class DefaultFontResolver(FontResolver):
    """Resolve configured and installed font faces with explicit fallback policy."""

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

        if self._strict:
            raise FontNotFoundError.for_font(
                requested,
                paragraph_index=paragraph_index,
                run_index=run_index,
            )

        fallback = self._configuration.default_font or "Helvetica"
        return self._resolve_substitution(
            requested,
            fallback,
            reason="lenient-default",
            east_asia=east_asia,
            paragraph_index=paragraph_index,
            run_index=run_index,
        )

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
        self._substitutions.append(
            FontSubstitution(
                requested_font=requested,
                selected_font=selected.family,
                reason=reason,
                paragraph_index=paragraph_index,
                run_index=run_index,
                east_asia=east_asia,
            )
        )
        return self._resolved(selected, substituted=True)

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
        )
