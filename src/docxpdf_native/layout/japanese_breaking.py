from __future__ import annotations

import unicodedata

from pydantic import BaseModel, ConfigDict


class TextCluster(BaseModel):
    """An indivisible, normalized text cluster and its normalized source range."""

    model_config = ConfigDict(frozen=True)

    text: str
    start: int
    end: int


class UnicodeText:
    """Unicode normalization and a dependency-free grapheme approximation."""

    _ZERO_WIDTH_JOINER = "\u200d"

    @staticmethod
    def normalize(text: str) -> str:
        return unicodedata.normalize("NFC", text)

    @staticmethod
    def grapheme_clusters(text: str) -> tuple[str, ...]:
        return tuple(cluster.text for cluster in UnicodeText.indexed_grapheme_clusters(text))

    @staticmethod
    def indexed_grapheme_clusters(text: str) -> tuple[TextCluster, ...]:
        if not text:
            return ()

        clusters: list[TextCluster] = []
        current = text[0]
        current_start = 0
        for index, character in enumerate(text[1:], start=1):
            if UnicodeText._extends_cluster(character, current):
                current += character
                continue
            clusters.append(
                TextCluster(
                    text=UnicodeText.normalize(current),
                    start=current_start,
                    end=index,
                )
            )
            current = character
            current_start = index
        clusters.append(
            TextCluster(
                text=UnicodeText.normalize(current),
                start=current_start,
                end=len(text),
            )
        )
        return tuple(clusters)

    @staticmethod
    def _extends_cluster(character: str, current: str) -> bool:
        if current.endswith(UnicodeText._ZERO_WIDTH_JOINER):
            return True
        if character == UnicodeText._ZERO_WIDTH_JOINER:
            return True
        if unicodedata.combining(character) != 0:
            return True
        codepoint = ord(character)
        if 0xFE00 <= codepoint <= 0xFE0F or 0xE0100 <= codepoint <= 0xE01EF:
            return True
        if 0x1F3FB <= codepoint <= 0x1F3FF:
            return True
        return UnicodeText._regional_indicator_pair(character, current)

    @staticmethod
    def _regional_indicator_pair(character: str, current: str) -> bool:
        return (
            len(current) == 1
            and 0x1F1E6 <= ord(current) <= 0x1F1FF
            and 0x1F1E6 <= ord(character) <= 0x1F1FF
        )


class JapaneseLineBreakingRules:
    """A compact kinsoku set for horizontal Japanese text."""

    PROHIBITED_LINE_START = frozenset(
        "、。，．・：；？！‼⁇⁈⁉ヽヾゝゞ々ー〜～…‥"  # noqa: RUF001
        "）〕］｝〉》」』】〙〗〟’”｠»"  # noqa: RUF001
        "ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ"
        ",.!?:;)]}"
    )
    PROHIBITED_LINE_END = frozenset(
        "（〔［｛〈《「『【〘〖〝‘“｟«￥＄￡＠＃([{"  # noqa: RUF001
    )

    @staticmethod
    def can_break_between(left: str, right: str) -> bool:
        if not left or not right:
            return False
        if left[-1] in JapaneseLineBreakingRules.PROHIBITED_LINE_END:
            return False
        if right[0] in JapaneseLineBreakingRules.PROHIBITED_LINE_START:
            return False
        left_is_word = JapaneseLineBreakingRules._word_character(left[-1])
        right_is_word = JapaneseLineBreakingRules._word_character(right[0])
        if left_is_word and right_is_word:
            return False
        if right.isspace():
            return False
        if left.isspace():
            return True
        return JapaneseLineBreakingRules._is_east_asian(
            left[-1]
        ) or JapaneseLineBreakingRules._is_east_asian(right[0])

    @staticmethod
    def _word_character(character: str) -> bool:
        return character.isalnum() and not JapaneseLineBreakingRules._is_east_asian(character)

    @staticmethod
    def _is_east_asian(character: str) -> bool:
        codepoint = ord(character)
        return (
            0x2E80 <= codepoint <= 0x9FFF
            or 0xF900 <= codepoint <= 0xFAFF
            or 0x20000 <= codepoint <= 0x3134F
            or 0x3040 <= codepoint <= 0x30FF
            or 0xAC00 <= codepoint <= 0xD7AF
            or unicodedata.east_asian_width(character) in {"W", "F"}
        )
