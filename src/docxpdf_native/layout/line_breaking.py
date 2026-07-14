from __future__ import annotations

import math
from collections.abc import Callable, Sequence

from pydantic import BaseModel, ConfigDict, computed_field

from docxpdf_native.layout.japanese_breaking import (
    JapaneseLineBreakingRules,
    TextCluster,
    UnicodeText,
)

MeasureFunction = Callable[[str], float]


class LineFragment(BaseModel):
    model_config = ConfigDict(frozen=True)

    text: str
    width: float
    start: int
    end: int
    is_tab: bool = False


class BrokenLine(BaseModel):
    model_config = ConfigDict(frozen=True)

    fragments: tuple[LineFragment, ...]
    width: float
    start: int
    end: int
    hard_break: bool = False
    page_break_after: bool = False

    @computed_field  # type: ignore[prop-decorator]
    @property
    def text(self) -> str:
        return "".join(fragment.text for fragment in self.fragments)


class LineBreaker:
    """Greedy line breaking with explicit breaks, tabs and Japanese kinsoku."""

    def break_text(
        self,
        text: str,
        *,
        max_width: float,
        measure: MeasureFunction,
        tab_width: float = 36.0,
    ) -> tuple[BrokenLine, ...]:
        if not math.isfinite(max_width) or max_width <= 0:
            raise ValueError("max_width must be a positive finite number")
        if not math.isfinite(tab_width) or tab_width <= 0:
            raise ValueError("tab_width must be a positive finite number")

        clusters = UnicodeText.indexed_grapheme_clusters(text)
        result: list[BrokenLine] = []
        segment: list[TextCluster] = []
        segment_start = 0
        ended_with_break = False
        previous_break_was_cr = False
        for cluster in clusters:
            if cluster.text not in {"\n", "\r", "\f"}:
                segment.append(cluster)
                ended_with_break = False
                previous_break_was_cr = False
                continue
            if cluster.text == "\n" and previous_break_was_cr:
                segment_start = cluster.end
                ended_with_break = True
                previous_break_was_cr = False
                continue
            lines = self._break_segment(
                segment,
                empty_offset=segment_start,
                max_width=max_width,
                measure=measure,
                tab_width=tab_width,
            )
            result.extend(lines)
            flag = "page_break_after" if cluster.text == "\f" else "hard_break"
            result[-1] = result[-1].model_copy(update={flag: True})
            segment = []
            segment_start = cluster.end
            ended_with_break = True
            previous_break_was_cr = cluster.text == "\r"

        if segment or ended_with_break or not result:
            result.extend(
                self._break_segment(
                    segment,
                    empty_offset=segment_start,
                    max_width=max_width,
                    measure=measure,
                    tab_width=tab_width,
                )
            )
        return tuple(result)

    def _break_segment(
        self,
        clusters: Sequence[TextCluster],
        *,
        empty_offset: int,
        max_width: float,
        measure: MeasureFunction,
        tab_width: float,
    ) -> list[BrokenLine]:
        if not clusters:
            return [self._line((), empty_offset=empty_offset, measure=measure, tab_width=tab_width)]

        lines: list[BrokenLine] = []
        current: list[TextCluster] = []
        for cluster in clusters:
            tentative = [*current, cluster]
            tentative_width = self._width(tentative, measure=measure, tab_width=tab_width)
            if tentative_width <= max_width or not current:
                current = tentative
                continue

            split = self._last_break(
                tentative,
                max_width=max_width,
                measure=measure,
                tab_width=tab_width,
            )
            if split is None:
                lines.append(
                    self._line(
                        current,
                        empty_offset=empty_offset,
                        measure=measure,
                        tab_width=tab_width,
                    )
                )
                current = [cluster]
                continue
            lines.append(
                self._line(
                    tentative[:split],
                    empty_offset=empty_offset,
                    measure=measure,
                    tab_width=tab_width,
                )
            )
            current = tentative[split:]

        if current:
            lines.append(
                self._line(
                    current,
                    empty_offset=empty_offset,
                    measure=measure,
                    tab_width=tab_width,
                )
            )
        return lines

    def _last_break(
        self,
        clusters: Sequence[TextCluster],
        *,
        max_width: float,
        measure: MeasureFunction,
        tab_width: float,
    ) -> int | None:
        candidate: int | None = None
        for index in range(1, len(clusters)):
            if not JapaneseLineBreakingRules.can_break_between(
                clusters[index - 1].text, clusters[index].text
            ):
                continue
            if self._width(clusters[:index], measure=measure, tab_width=tab_width) <= max_width:
                candidate = index
        return candidate

    def _line(
        self,
        clusters: Sequence[TextCluster],
        *,
        empty_offset: int,
        measure: MeasureFunction,
        tab_width: float,
    ) -> BrokenLine:
        fragments: list[LineFragment] = []
        used = 0.0
        text_group: list[TextCluster] = []

        def flush_text_group() -> None:
            nonlocal used
            if not text_group:
                return
            grouped_text = "".join(item.text for item in text_group)
            width = self._measured_width(grouped_text, measure)
            fragments.append(
                LineFragment(
                    text=grouped_text,
                    width=width,
                    start=text_group[0].start,
                    end=text_group[-1].end,
                )
            )
            used += width
            text_group.clear()

        for cluster in clusters:
            if cluster.text != "\t":
                text_group.append(cluster)
                continue
            flush_text_group()
            width = self._tab_advance(current_width=used, tab_width=tab_width)
            fragments.append(
                LineFragment(
                    text=cluster.text,
                    width=width,
                    start=cluster.start,
                    end=cluster.end,
                    is_tab=True,
                )
            )
            used += width
        flush_text_group()
        start = clusters[0].start if clusters else empty_offset
        end = clusters[-1].end if clusters else empty_offset
        return BrokenLine(fragments=tuple(fragments), width=used, start=start, end=end)

    def _width(
        self,
        clusters: Sequence[TextCluster],
        *,
        measure: MeasureFunction,
        tab_width: float,
    ) -> float:
        used = 0.0
        text_group: list[str] = []
        for cluster in clusters:
            if cluster.text != "\t":
                text_group.append(cluster.text)
                continue
            if text_group:
                used += self._measured_width("".join(text_group), measure)
                text_group = []
            used += self._tab_advance(current_width=used, tab_width=tab_width)
        if text_group:
            used += self._measured_width("".join(text_group), measure)
        return used

    @staticmethod
    def _measured_width(text: str, measure: MeasureFunction) -> float:
        width = measure(text)
        if not math.isfinite(width) or width < 0:
            raise ValueError("text measurement must be a non-negative finite number")
        return width

    @staticmethod
    def _tab_advance(*, current_width: float, tab_width: float) -> float:
        return tab_width - math.fmod(current_width, tab_width)
