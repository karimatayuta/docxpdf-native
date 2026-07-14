from __future__ import annotations

import unicodedata

import pytest

from docxpdf_native.layout.japanese_breaking import JapaneseLineBreakingRules, UnicodeText
from docxpdf_native.layout.line_breaking import LineBreaker


def _cluster_width(text: str) -> float:
    return float(len(UnicodeText.grapheme_clusters(text)))


def test_unicode_text_normalizes_to_nfc() -> None:
    assert UnicodeText.normalize("か\N{COMBINING KATAKANA-HIRAGANA VOICED SOUND MARK}") == "が"


def test_grapheme_clusters_keep_combining_marks_and_variation_selectors() -> None:
    text = "A\N{COMBINING ACUTE ACCENT}葛\N{VARIATION SELECTOR-1}B"

    assert UnicodeText.grapheme_clusters(text) == ("Á", "葛\N{VARIATION SELECTOR-1}", "B")


def test_indexed_clusters_keep_original_offsets_after_nfc_composition() -> None:
    source = "A\N{COMBINING ACUTE ACCENT}B"

    clusters = UnicodeText.indexed_grapheme_clusters(source)

    assert tuple((cluster.text, cluster.start, cluster.end) for cluster in clusters) == (
        ("Á", 0, 2),
        ("B", 2, 3),
    )


def test_grapheme_clusters_keep_emoji_zwj_sequence() -> None:
    assert UnicodeText.grapheme_clusters("👩\u200d💻X") == ("👩\u200d💻", "X")


@pytest.mark.parametrize(
    "right",
    ["、", "。", "）", "」", "ぁ", "ー", ")", ",", "."],  # noqa: RUF001
)
def test_japanese_rule_forbids_line_start_characters(right: str) -> None:
    assert JapaneseLineBreakingRules.can_break_between("日", right) is False


@pytest.mark.parametrize(
    "left",
    ["（", "「", "『", "￥", "(", "["],  # noqa: RUF001
)
def test_japanese_rule_forbids_line_end_characters(left: str) -> None:
    assert JapaneseLineBreakingRules.can_break_between(left, "日") is False


def test_japanese_rule_allows_break_between_ordinary_cjk_characters() -> None:
    assert JapaneseLineBreakingRules.can_break_between("日", "本") is True


@pytest.mark.parametrize(("left", "right"), [("A", "B"), ("1", "2"), ("A", "2")])
def test_japanese_rule_suppresses_break_inside_alphanumeric_token(left: str, right: str) -> None:
    assert JapaneseLineBreakingRules.can_break_between(left, right) is False


def test_line_breaker_prefers_spaces_for_english_text() -> None:
    lines = LineBreaker().break_text("one two three", max_width=7.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("one ", "two ", "three")


def test_line_breaker_does_not_split_alphanumeric_token_when_it_can_move_whole() -> None:
    lines = LineBreaker().break_text("a abcde", max_width=5.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("a ", "abcde")


def test_line_breaker_forces_a_long_unbreakable_token_without_losing_text() -> None:
    lines = LineBreaker().break_text("abcdefgh", max_width=3.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("abc", "def", "gh")
    assert "".join(line.text for line in lines) == "abcdefgh"


def test_line_breaker_observes_japanese_line_start_prohibition() -> None:
    lines = LineBreaker().break_text("日本、語", max_width=2.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("日", "本、", "語")


def test_line_breaker_keeps_combining_sequence_intact() -> None:
    source = unicodedata.normalize("NFD", "é") + "x"
    lines = LineBreaker().break_text(source, max_width=1.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("é", "x")


def test_line_breaker_emits_empty_line_for_explicit_newline() -> None:
    lines = LineBreaker().break_text("a\n\nb", max_width=10.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("a", "", "b")
    assert tuple(line.hard_break for line in lines) == (True, True, False)


def test_line_breaker_treats_crlf_as_one_explicit_newline() -> None:
    lines = LineBreaker().break_text("a\r\nb", max_width=10.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("a", "b")
    assert tuple(line.hard_break for line in lines) == (True, False)


def test_line_breaker_marks_explicit_page_break() -> None:
    lines = LineBreaker().break_text("before\fafter", max_width=20.0, measure=_cluster_width)

    assert tuple(line.text for line in lines) == ("before", "after")
    assert lines[0].page_break_after is True


def test_line_breaker_advances_tab_to_next_stop() -> None:
    lines = LineBreaker().break_text("a\tb", max_width=20.0, measure=_cluster_width, tab_width=4.0)

    assert len(lines) == 1
    assert lines[0].width == 5.0
    assert tuple(fragment.width for fragment in lines[0].fragments) == (1.0, 3.0, 1.0)


def test_line_breaker_measures_cluster_spacing_across_a_text_run() -> None:
    def measure_with_spacing(text: str) -> float:
        cluster_count = len(UnicodeText.grapheme_clusters(text))
        return cluster_count + max(cluster_count - 1, 0) * 0.5

    lines = LineBreaker().break_text("abc", max_width=3.0, measure=measure_with_spacing)

    assert tuple(line.text for line in lines) == ("ab", "c")
    assert lines[0].width == 2.5


def test_line_breaker_rejects_invalid_dimensions() -> None:
    with pytest.raises(ValueError, match="max_width"):
        LineBreaker().break_text("text", max_width=0.0, measure=_cluster_width)
