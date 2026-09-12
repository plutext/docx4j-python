"""The parts of a package. docx4j ``org.docx4j.openpackaging.parts``.

Everything is re-exported from :mod:`docx4j_py.openpackaging`; import from there
unless you want one module.
"""

from __future__ import annotations

from docx4j_py.openpackaging.parts.binary_part import (
    AlternativeFormatInputPart,
    BinaryPart,
    EmbeddedPackagePart,
    ImagePart,
    ObfuscatedFontPart,
    OleObjectBinaryPart,
    TrueTypeFontPart,
    VbaProjectBinaryPart,
)
from docx4j_py.openpackaging.parts.default_xml_part import (
    CustomXmlDataStoragePart,
    DefaultXmlPart,
    VMLPart,
)
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.parts_map import Parts
from docx4j_py.openpackaging.parts.registry import PartRegistry, default_part_registry
from docx4j_py.openpackaging.parts.relationships_part import (
    AddPartBehaviour,
    RelationshipsPart,
)
from docx4j_py.openpackaging.parts.xml_part import XmlPart

__all__ = [
    "AddPartBehaviour",
    "AlternativeFormatInputPart",
    "BinaryPart",
    "CustomXmlDataStoragePart",
    "DefaultXmlPart",
    "EmbeddedPackagePart",
    "ImagePart",
    "Namespaces",
    "ObfuscatedFontPart",
    "OleObjectBinaryPart",
    "Part",
    "PartRegistry",
    "Parts",
    "RelationshipsPart",
    "TrueTypeFontPart",
    "VMLPart",
    "VbaProjectBinaryPart",
    "XmlPart",
    "default_part_registry",
]
