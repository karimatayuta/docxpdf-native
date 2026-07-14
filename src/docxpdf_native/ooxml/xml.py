"""Bounded XML parsing based only on :mod:`xml.etree.ElementTree`."""

from __future__ import annotations

from io import BytesIO
from xml.etree import ElementTree
from xml.etree.ElementTree import Element

from docxpdf_native.exceptions import InvalidOoxmlError, ResourceLimitError


class SafeXmlParser:
    """Parse an XML part after applying size, declaration, and depth limits."""

    def __init__(self, *, max_part_size: int, max_depth: int) -> None:
        if max_part_size <= 0:
            raise ValueError("max_part_size must be greater than zero")
        if max_depth <= 0:
            raise ValueError("max_depth must be greater than zero")
        self._max_part_size = max_part_size
        self._max_depth = max_depth

    def parse(self, data: bytes, *, part_name: str) -> Element:
        """Return the root element or raise a domain-specific safe error."""

        size = len(data)
        if size > self._max_part_size:
            raise ResourceLimitError(
                f"XML part {part_name!r} is {size} bytes; limit is {self._max_part_size}"
            )
        if not data.strip():
            raise InvalidOoxmlError(f"XML part {part_name!r} is empty")
        lowered = data.lower()
        if b"<!doctype" in lowered or b"<!entity" in lowered:
            raise InvalidOoxmlError(
                f"DTD or entity declarations are not allowed in XML part {part_name!r}"
            )

        depth = 0
        root: Element | None = None
        try:
            for event, element in ElementTree.iterparse(BytesIO(data), events=("start", "end")):
                if event == "start":
                    depth += 1
                    if root is None:
                        root = element
                    if depth > self._max_depth:
                        raise ResourceLimitError(
                            f"XML depth in {part_name!r} exceeds limit {self._max_depth}"
                        )
                else:
                    depth -= 1
        except ResourceLimitError:
            raise
        except (ElementTree.ParseError, ValueError) as error:
            raise InvalidOoxmlError(f"Malformed XML in part {part_name!r}: {error}") from error

        if root is None:
            raise InvalidOoxmlError(f"XML part {part_name!r} is empty")
        return root
