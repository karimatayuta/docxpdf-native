"""Safe Office Open XML package and document parsing."""

from __future__ import annotations

from docxpdf_native.ooxml.namespaces import OoxmlNamespaces
from docxpdf_native.ooxml.xml import SafeXmlParser

__all__ = ["OoxmlNamespaces", "SafeXmlParser"]
