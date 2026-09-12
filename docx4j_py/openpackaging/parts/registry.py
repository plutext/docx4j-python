"""``PartRegistry``: content type -> part class. docx4j ``ContentTypeManager.newPartForContentType``.

CR-002 section 5.4. The table below is docx4j's, entry for entry, with its
resolution order preserved:

1. the **relationship type** first, for the four generic content types where the
   relationship is what says what the part is (``aFChunk``, ``package``,
   ``oleObject``, ``customXml``) --- docx4j tests these before the content type
   and so does this;
2. then the **content type**;
3. then the prefixes docx4j matches with ``startsWith`` (the DrawingML, PML and
   SML families);
4. then ``image/*`` and the ``+xml`` fallback;
5. and last, :class:`BinaryPart`, so that **nothing is ever dropped**.

Word's own ``application/xml`` is deliberately a custom XML data part rather than
a generic one, as in docx4j, because that string is both
``ContentTypes.APPLICATION_XML`` and
``ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGE`` and docx4j tests the
latter first.

Extend with ``registry.register(content_type, factory)`` or
``registry.register_for_relationship(relationship_type, factory)``; a package can
be loaded with a registry of its own through ``LoadOptions(registry=...)``.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docx4j_py.openpackaging.content_types import (
    IMAGE_CONTENT_TYPES_BY_EXTENSION,
    ContentTypes,
    is_xml_content_type,
)
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts import dml, docprops, wml
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
)
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part

__all__ = ["PartFactory", "PartRegistry", "default_part_registry"]

#: What the registry stores: ``(part_name, content_type) -> Part``.
PartFactory = Callable[[PartName, str], Part]

#: The content-type prefixes docx4j matches with ``startsWith``, and what they
#: mean when nothing more specific matched. Only the DrawingML one resolves to
#: a typed part here; PML and SML are Phase C (CR-002 section 1).
_DRAWING_PREFIX = "application/vnd.openxmlformats-officedocument.drawing"
_PRESENTATION_PREFIX = "application/vnd.openxmlformats-officedocument.presentation"
_SPREADSHEET_PREFIX = "application/vnd.openxmlformats-officedocument.spreadsheetml"


class PartRegistry:
    """Content type and relationship type to part class."""

    __slots__ = ("_by_content_type", "_by_relationship_type")

    def __init__(self) -> None:
        """An empty registry; :data:`default_part_registry` is the filled one."""
        self._by_content_type: dict[str, PartFactory] = {}
        self._by_relationship_type: dict[str, PartFactory] = {}

    def register(self, content_type: str, factory: PartFactory) -> PartRegistry:
        """Map a content type to a factory; returns self, so calls chain."""
        self._by_content_type[content_type] = factory
        return self

    def register_for_relationship(
        self, relationship_type: str, factory: PartFactory
    ) -> PartRegistry:
        """Map a relationship type to a factory. Consulted *before* the content type."""
        self._by_relationship_type[relationship_type] = factory
        return self

    def copy(self) -> PartRegistry:
        """A registry with the same tables, for a caller who wants to extend one."""
        other = PartRegistry()
        other._by_content_type = dict(self._by_content_type)
        other._by_relationship_type = dict(self._by_relationship_type)
        return other

    def create_part(
        self,
        part_name: PartName | str,
        content_type: str | None,
        rel: Any = None,
    ) -> Part:
        """Build the part for a name, a content type and the relationship that reached it."""
        name = PartName.of(part_name)
        rel_type = getattr(rel, "type_value", None) if rel is not None else None

        if rel_type:
            by_rel = self._by_relationship_type.get(rel_type)
            if by_rel is not None:
                return by_rel(name, content_type or ContentTypes.APPLICATION_XML)

        if content_type:
            factory = self._by_content_type.get(content_type)
            if factory is not None:
                return factory(name, content_type)
            if content_type.startswith("image/"):
                return ImagePart(name, content_type)
            if content_type.startswith(
                (_DRAWING_PREFIX, _PRESENTATION_PREFIX, _SPREADSHEET_PREFIX)
            ) or is_xml_content_type(content_type):
                return DefaultXmlPart(name, content_type, rel_type or "")
            if name.extension.lower() == "xml":
                return DefaultXmlPart(name, content_type, rel_type or "")
            return BinaryPart(name, content_type, rel_type or "")

        # No content type at all: Word would reject the package, but the bytes
        # are still kept rather than dropped.
        if rel_type == Namespaces.IMAGE:
            return ImagePart(
                name,
                IMAGE_CONTENT_TYPES_BY_EXTENSION.get(
                    name.extension.lower(), ContentTypes.OCTET_STREAM
                ),
            )
        if name.extension.lower() == "xml":
            return DefaultXmlPart(name, ContentTypes.APPLICATION_XML, rel_type or "")
        return BinaryPart(name, ContentTypes.OCTET_STREAM, rel_type or "")

    def __repr__(self) -> str:
        """``PartRegistry(48 content types, 4 relationship types)``."""
        return (
            f"PartRegistry({len(self._by_content_type)} content types, "
            f"{len(self._by_relationship_type)} relationship types)"
        )


def _build_default() -> PartRegistry:
    """docx4j's table."""
    registry = PartRegistry()

    # 1. relationship type first, as docx4j's newPartForContentType does
    registry.register_for_relationship(
        Namespaces.AF, lambda n, ct: AlternativeFormatInputPart(n, ct)
    )
    registry.register_for_relationship(
        Namespaces.EMBEDDED_PKG, lambda n, ct: EmbeddedPackagePart(n, ct)
    )
    registry.register_for_relationship(
        Namespaces.OLE_OBJECT, lambda n, ct: OleObjectBinaryPart(n, ct)
    )
    registry.register_for_relationship(
        Namespaces.CUSTOM_XML_DATA_STORAGE,
        lambda n, ct: CustomXmlDataStoragePart(n, ct),
    )

    # 2. WordprocessingML
    for main_ct in (
        ContentTypes.WORDPROCESSINGML_DOCUMENT,
        ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED,
        ContentTypes.WORDPROCESSINGML_TEMPLATE,
        ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED,
    ):
        registry.register(main_ct, lambda n, ct: wml.MainDocumentPart(n, ct))
    registry.register(
        ContentTypes.WORDPROCESSINGML_GLOSSARYDOCUMENT,
        lambda n, ct: wml.GlossaryDocumentPart(n),
    )
    registry.register(
        ContentTypes.WORDPROCESSINGML_STYLES,
        lambda n, ct: wml.StyleDefinitionsPart(n, ct),
    )
    registry.register(
        ContentTypes.WORDPROCESSINGML_STYLESWITHEFFECTS,
        lambda n, ct: wml.StyleDefinitionsPart(n, ct),
    )
    registry.register(
        ContentTypes.WORDPROCESSINGML_NUMBERING,
        lambda n, ct: wml.NumberingDefinitionsPart(n),
    )
    registry.register(ContentTypes.WORDPROCESSINGML_FONTTABLE, lambda n, ct: wml.FontTablePart(n))
    registry.register(
        ContentTypes.WORDPROCESSINGML_SETTINGS,
        lambda n, ct: wml.DocumentSettingsPart(n),
    )
    registry.register(
        ContentTypes.WORDPROCESSINGML_WEBSETTINGS, lambda n, ct: wml.WebSettingsPart(n)
    )
    registry.register(ContentTypes.WORDPROCESSINGML_HEADER, lambda n, ct: wml.HeaderPart(n))
    registry.register(ContentTypes.WORDPROCESSINGML_FOOTER, lambda n, ct: wml.FooterPart(n))
    registry.register(ContentTypes.WORDPROCESSINGML_FOOTNOTES, lambda n, ct: wml.FootnotesPart(n))
    registry.register(ContentTypes.WORDPROCESSINGML_ENDNOTES, lambda n, ct: wml.EndnotesPart(n))
    registry.register(ContentTypes.WORDPROCESSINGML_COMMENTS, lambda n, ct: wml.CommentsPart(n))
    registry.register(
        ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENDED,
        lambda n, ct: wml.CommentsExtendedPart(n),
    )
    registry.register(ContentTypes.WORDPROCESSINGML_PEOPLE, lambda n, ct: wml.PeoplePart(n))
    registry.register(ContentTypes.MS_WORD_KEYMAP, lambda n, ct: wml.KeyMapCustomizationsPart(n))
    registry.register(ContentTypes.OFFICEDOCUMENT_VBA_DATA, lambda n, ct: wml.VbaDataPart(n))
    registry.register(ContentTypes.VML_DRAWING, lambda n, ct: wml.VMLPart(n))
    # commentsIds and commentsExtensible: w16cid / w16cex are outside this
    # build's schema closure (CR-001 section 13.5), so they are trees.
    registry.register(
        ContentTypes.WORDPROCESSINGML_COMMENTS_IDS,
        lambda n, ct: DefaultXmlPart(n, ct, Namespaces.COMMENTS_IDS),
    )
    registry.register(
        ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENSIBLE,
        lambda n, ct: DefaultXmlPart(n, ct, Namespaces.COMMENTS_EXTENSIBLE),
    )

    # 3. DrawingML
    registry.register(ContentTypes.OFFICEDOCUMENT_THEME, lambda n, ct: dml.ThemePart(n))
    registry.register(
        ContentTypes.OFFICEDOCUMENT_THEME_OVERRIDE,
        lambda n, ct: dml.ThemeOverridePart(n),
    )
    registry.register(ContentTypes.DRAWINGML_CHART, lambda n, ct: dml.ChartPart(n))
    registry.register(ContentTypes.DRAWINGML_DRAWING, lambda n, ct: dml.DrawingPart(n))
    registry.register(ContentTypes.DRAWINGML_CHART_SHAPES, lambda n, ct: dml.ChartShapePart(n))
    registry.register(ContentTypes.DRAWINGML_DIAGRAM_DATA, lambda n, ct: dml.DiagramDataPart(n))
    registry.register(
        ContentTypes.DRAWINGML_DIAGRAM_LAYOUT, lambda n, ct: dml.DiagramLayoutPart(n)
    )
    registry.register(
        ContentTypes.DRAWINGML_DIAGRAM_LAYOUT_HEADER,
        lambda n, ct: dml.DiagramLayoutHeaderPart(n),
    )
    registry.register(ContentTypes.DRAWINGML_DIAGRAM_STYLE, lambda n, ct: dml.DiagramStylePart(n))
    registry.register(
        ContentTypes.DRAWINGML_DIAGRAM_COLORS, lambda n, ct: dml.DiagramColorsPart(n)
    )
    registry.register(
        ContentTypes.DRAWINGML_DIAGRAM_DRAWING, lambda n, ct: dml.DiagramDrawingPart(n)
    )
    registry.register(ContentTypes.CHART_STYLE, lambda n, ct: dml.ChartStylePart(n))
    registry.register(ContentTypes.CHART_COLOR_STYLE, lambda n, ct: dml.ChartColorStylePart(n))
    registry.register(ContentTypes.CHART_EX, lambda n, ct: dml.ChartExSpacePart(n))

    # 4. document properties and custom XML
    registry.register(
        ContentTypes.PACKAGE_COREPROPERTIES, lambda n, ct: docprops.DocPropsCorePart(n)
    )
    registry.register(
        ContentTypes.OFFICEDOCUMENT_EXTENDEDPROPERTIES,
        lambda n, ct: docprops.DocPropsExtendedPart(n),
    )
    registry.register(
        ContentTypes.OFFICEDOCUMENT_CUSTOMPROPERTIES,
        lambda n, ct: docprops.DocPropsCustomPart(n),
    )
    registry.register(
        ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGEPROPERTIES,
        lambda n, ct: docprops.CustomXmlDataStoragePropertiesPart(n),
    )
    # docx4j tests OFFICEDOCUMENT_CUSTOMXML_DATASTORAGE ('application/xml')
    # before APPLICATION_XML, so an application/xml part is a custom XML data
    # part. The relationship-type entry above usually settles it first.
    registry.register(
        ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGE,
        lambda n, ct: CustomXmlDataStoragePart(n, ct),
    )

    # 5. fonts, OLE, VBA
    registry.register(ContentTypes.OFFICEDOCUMENT_FONT, lambda n, ct: ObfuscatedFontPart(n))
    registry.register(ContentTypes.TRUETYPE_FONT, lambda n, ct: TrueTypeFontPart(n))
    registry.register(
        ContentTypes.OFFICEDOCUMENT_OLE_OBJECT, lambda n, ct: OleObjectBinaryPart(n, ct)
    )
    registry.register(
        ContentTypes.OFFICEDOCUMENT_ACTIVEX_OBJECT,
        lambda n, ct: OleObjectBinaryPart(n, ct),
    )
    registry.register(
        ContentTypes.OFFICEDOCUMENT_VBA_PROJECT,
        lambda n, ct: VbaProjectBinaryPart(n, ct),
    )
    registry.register(
        ContentTypes.OFFICEDOCUMENT_VBA_PROJECT_SIGNATURE,
        lambda n, ct: BinaryPart(n, ct, Namespaces.VBA_PROJECT_SIGNATURE),
    )
    registry.register(
        ContentTypes.PRESENTATIONML_FONT_DATA,
        lambda n, ct: BinaryPart(n, ct, Namespaces.FONT),
    )
    registry.register(
        ContentTypes.DIGITAL_SIGNATURE_ORIGIN_PART,
        lambda n, ct: BinaryPart(n, ct, Namespaces.DIGITAL_SIGNATURE_ORIGIN),
    )
    registry.register(
        ContentTypes.SPREADSHEETML_PRINTER_SETTINGS,
        lambda n, ct: BinaryPart(n, ct, Namespaces.SPREADSHEETML_PRINTER_SETTINGS),
    )
    return registry


#: docx4j's table. Extend it with :meth:`PartRegistry.copy` and ``register``.
default_part_registry = _build_default()
