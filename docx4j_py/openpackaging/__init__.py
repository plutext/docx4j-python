"""The engine: Open Packaging, parts, load and save. CR-002 Phase A.

``docx4j_py.wml`` gives WordprocessingML typed objects; this gives them a
document to live in. It reads and writes Office Open XML packages from zip
files, bytes, file objects and directories, and exposes their parts as typed
objects related by relationships exactly as docx4j does.

    >>> from docx4j_py.openpackaging import WordprocessingMLPackage
    >>> pkg = WordprocessingMLPackage.load("in.docx")  # doctest: +SKIP
    >>> body = pkg.main_document_part.contents.body  # doctest: +SKIP
    >>> pkg.save("out.docx")  # doctest: +SKIP

Two rules run through everything (CR-002 section 1):

**A part that is never touched is written back byte for byte.** Only a part
whose content was unmarshalled is re-marshalled on save, so a round trip is safe
with markup this model has never seen.

**A part that is unmarshalled loses nothing silently.** The fork's
skipped-content report is attached to every unmarshalled part and logged;
``LoadOptions(strict=True)`` raises instead.

Names follow ``org.docx4j.openpackaging`` in ``snake_case``. The departures are
listed in CR-002 section 12.
"""

from __future__ import annotations

from docx4j_py.openpackaging.api import create_package, load
from docx4j_py.openpackaging.content_types import (
    ContentTypeManager,
    ContentTypes,
    is_stored_uncompressed,
    is_xml_content_type,
)
from docx4j_py.openpackaging.exceptions import (
    Docx4JException,
    InvalidFormatException,
    InvalidOperationException,
    PartUnrecognisedException,
)
from docx4j_py.openpackaging.load import LoadOptions, load_package
from docx4j_py.openpackaging.mce import mce_preprocess, resolve_alternate_content
from docx4j_py.openpackaging.packages.opc_package import OpcPackage
from docx4j_py.openpackaging.packages.wordprocessingml_package import (
    PAGE_MARGINS,
    PAGE_SIZES,
    PageSizePaper,
    WordprocessingMLPackage,
)
from docx4j_py.openpackaging.part_name import (
    CONTENT_TYPES_NAME,
    PACKAGE_RELS_NAME,
    PartName,
)
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
from docx4j_py.openpackaging.parts.dml import (
    ChartColorStylePart,
    ChartExSpacePart,
    ChartPart,
    ChartShapePart,
    ChartStylePart,
    DiagramColorsPart,
    DiagramDataPart,
    DiagramDrawingPart,
    DiagramLayoutHeaderPart,
    DiagramLayoutPart,
    DiagramStylePart,
    DrawingPart,
    ThemeOverridePart,
    ThemePart,
)
from docx4j_py.openpackaging.parts.docprops import (
    CustomXmlDataStoragePropertiesPart,
    DocPropsCorePart,
    DocPropsCustomPart,
    DocPropsExtendedPart,
)
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.parts_map import Parts
from docx4j_py.openpackaging.parts.registry import PartRegistry, default_part_registry
from docx4j_py.openpackaging.parts.relationships_part import (
    AddPartBehaviour,
    RelationshipsPart,
    is_external,
)
from docx4j_py.openpackaging.parts.wml import (
    CommentsExtendedPart,
    CommentsPart,
    DocumentPart,
    DocumentSettingsPart,
    EndnotesPart,
    FontTablePart,
    FooterPart,
    FootnotesPart,
    GlossaryDocumentPart,
    HeaderPart,
    KeyMapCustomizationsPart,
    MainDocumentPart,
    NumberingDefinitionsPart,
    PeoplePart,
    StyleDefinitionsPart,
    VbaDataPart,
    WebSettingsPart,
)
from docx4j_py.openpackaging.parts.xml_part import XML_DECLARATION, XmlPart
from docx4j_py.openpackaging.save import save_package
from docx4j_py.openpackaging.stores import (
    DirectoryPartSink,
    DirectoryPartStore,
    MemoryPartSink,
    MemoryPartStore,
    PartSink,
    PartStore,
    ZipPartSink,
    ZipPartStore,
)

__all__ = [
    "CONTENT_TYPES_NAME",
    "PACKAGE_RELS_NAME",
    "PAGE_MARGINS",
    "PAGE_SIZES",
    "XML_DECLARATION",
    "AddPartBehaviour",
    "AlternativeFormatInputPart",
    "BinaryPart",
    "ChartColorStylePart",
    "ChartExSpacePart",
    "ChartPart",
    "ChartShapePart",
    "ChartStylePart",
    "CommentsExtendedPart",
    "CommentsPart",
    "ContentTypeManager",
    "ContentTypes",
    "CustomXmlDataStoragePart",
    "CustomXmlDataStoragePropertiesPart",
    "DefaultXmlPart",
    "DiagramColorsPart",
    "DiagramDataPart",
    "DiagramDrawingPart",
    "DiagramLayoutHeaderPart",
    "DiagramLayoutPart",
    "DiagramStylePart",
    "DirectoryPartSink",
    "DirectoryPartStore",
    "DocPropsCorePart",
    "DocPropsCustomPart",
    "DocPropsExtendedPart",
    "DocumentPart",
    "DocumentSettingsPart",
    "Docx4JException",
    "DrawingPart",
    "EmbeddedPackagePart",
    "EndnotesPart",
    "FontTablePart",
    "FooterPart",
    "FootnotesPart",
    "GlossaryDocumentPart",
    "HeaderPart",
    "ImagePart",
    "InvalidFormatException",
    "InvalidOperationException",
    "KeyMapCustomizationsPart",
    "LoadOptions",
    "MainDocumentPart",
    "MemoryPartSink",
    "MemoryPartStore",
    "Namespaces",
    "NumberingDefinitionsPart",
    "ObfuscatedFontPart",
    "OleObjectBinaryPart",
    "OpcPackage",
    "PageSizePaper",
    "Part",
    "PartName",
    "PartRegistry",
    "PartSink",
    "PartStore",
    "PartUnrecognisedException",
    "Parts",
    "PeoplePart",
    "RelationshipsPart",
    "StyleDefinitionsPart",
    "ThemeOverridePart",
    "ThemePart",
    "TrueTypeFontPart",
    "VMLPart",
    "VbaDataPart",
    "VbaProjectBinaryPart",
    "WebSettingsPart",
    "WordprocessingMLPackage",
    "XmlPart",
    "ZipPartSink",
    "ZipPartStore",
    "create_package",
    "default_part_registry",
    "is_external",
    "is_stored_uncompressed",
    "is_xml_content_type",
    "load",
    "load_package",
    "mce_preprocess",
    "resolve_alternate_content",
    "save_package",
]
