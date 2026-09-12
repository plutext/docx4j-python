"""The document properties parts. docx4j ``DocPropsCorePart`` and friends.

CR-002 sections 5.4 and 9. Three parts, three generated packages:
``docx4j_py.docprops.core`` (with Dublin Core), ``.extended`` and ``.custom``.

``/customXml/itemPropsN.xml`` is *not* here: its ``ds:datastoreItem`` lives in
``…/officeDocument/2006/customXml``, which is not in this build's schema closure
(the generated ``docx4j_py.customxml`` is the *schema library*, ``sl:``), so
:class:`CustomXmlDataStoragePropertiesPart` is an lxml tree. It is read for its
``ds:itemID`` and written back unchanged, which is all anything does with it.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.default_xml_part import DefaultXmlPart
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.xml_part import XmlPart

__all__ = [
    "DS_NS",
    "CustomXmlDataStoragePropertiesPart",
    "DocPropsCorePart",
    "DocPropsCustomPart",
    "DocPropsExtendedPart",
]

#: The custom XML data storage properties namespace.
DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"

_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_CUP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"


class DocPropsCorePart(XmlPart[Any]):
    """``/docProps/core.xml``, ``cp:coreProperties``. docx4j ``DocPropsCorePart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.docprops.core.CoreProperties"
    bool_format = "words"

    def __init__(self, part_name: PartName | str = "/docProps/core.xml") -> None:
        """Build the core properties part."""
        super().__init__(
            part_name,
            ContentTypes.PACKAGE_COREPROPERTIES,
            Namespaces.PROPERTIES_CORE,
            f"{{{_CP}}}coreProperties",
        )


class DocPropsExtendedPart(XmlPart[Any]):
    """``/docProps/app.xml``. docx4j ``DocPropsExtendedPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.docprops.extended.Properties"
    bool_format = "words"
    #: Word writes ``<Properties xmlns="...extended-properties" xmlns:vt="...">``.
    default_namespace = _EP

    def __init__(self, part_name: PartName | str = "/docProps/app.xml") -> None:
        """Build the extended properties part."""
        super().__init__(
            part_name,
            ContentTypes.OFFICEDOCUMENT_EXTENDEDPROPERTIES,
            Namespaces.PROPERTIES_EXTENDED,
            f"{{{_EP}}}Properties",
        )


class DocPropsCustomPart(XmlPart[Any]):
    """``/docProps/custom.xml``. docx4j ``DocPropsCustomPart``."""

    __slots__ = ()
    model_class_path = "docx4j_py.docprops.custom.Properties"
    bool_format = "words"

    def __init__(self, part_name: PartName | str = "/docProps/custom.xml") -> None:
        """Build the custom properties part."""
        super().__init__(
            part_name,
            ContentTypes.OFFICEDOCUMENT_CUSTOMPROPERTIES,
            Namespaces.PROPERTIES_CUSTOM,
            f"{{{_CUP}}}Properties",
        )


class CustomXmlDataStoragePropertiesPart(DefaultXmlPart):
    """``/customXml/itemPropsN.xml``. docx4j ``CustomXmlDataStoragePropertiesPart``."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the custom XML properties part."""
        super().__init__(
            part_name,
            ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGEPROPERTIES,
            Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES,
        )

    @property
    def item_id(self) -> str | None:
        """The ``ds:itemID`` of the ``ds:datastoreItem`` root, or None.

        Read off a *throwaway* parse of the root element, not off
        :attr:`tree`: the loader asks every custom XML part for this, and a part
        whose tree has been built is no longer written back byte for byte.
        """
        import io

        from lxml import etree

        try:
            for _event, element in etree.iterparse(
                io.BytesIO(self.bytes_for_save),
                events=("start",),
                resolve_entities=False,
                recover=True,
            ):
                return element.get(f"{{{DS_NS}}}itemID") or element.get("itemID")
        except etree.XMLSyntaxError:
            return None
        return None
