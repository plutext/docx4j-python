"""XML parts the model does not type, held as lxml trees.

CR-002 section 5.4. Three of them:

:class:`DefaultXmlPart`
    docx4j ``DefaultXmlPart``: XML the registry has no class for. Nothing is
    lost --- a part that is never parsed is written back byte for byte, and one
    that is parsed is an lxml tree, which holds everything.
:class:`CustomXmlDataStoragePart`
    docx4j ``CustomXmlDataStoragePart``: a customer's own schema, which by
    definition no generated binding covers. Indexed by the ``ds:itemID`` of its
    sibling properties part.
:class:`VMLPart`
    docx4j ``VMLPart``: Word writes an unqualified ``<xml>`` root and VML is not
    in this model's schema closure (CR-001 section 13.5), so it is a tree. The
    same departure docx4j-core-ts made.
"""

from __future__ import annotations

from lxml import etree

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part

__all__ = ["CustomXmlDataStoragePart", "DefaultXmlPart", "VMLPart"]

_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


class DefaultXmlPart(Part):
    """An XML part held as an lxml tree, parsed on first access."""

    __slots__ = ("_bytes", "_tree")

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = ContentTypes.APPLICATION_XML,
        relationship_type: str = "",
    ) -> None:
        """Build the part; the tree is parsed when it is first asked for."""
        super().__init__(part_name, content_type, relationship_type)
        self._tree: etree._Element | None = None
        self._bytes: bytes | None = None

    @property
    def is_loaded(self) -> bool:
        """Whether a tree or bytes have been set."""
        return self._tree is not None or self._bytes is not None

    @property
    def is_parsed(self) -> bool:
        """Whether the tree exists; a part that is not parsed is copied verbatim."""
        return self._tree is not None

    @property
    def tree(self) -> etree._Element:
        """The root element, parsed on first access."""
        if self._tree is None:
            data = self._bytes if self._bytes is not None else self._source_bytes()
            if data is None:
                raise Docx4JException(
                    f"Part {self.part_name} has no content: set a tree, or load it "
                    "from a container"
                )
            self._tree = etree.fromstring(data, _PARSER)
            self._bytes = None
        return self._tree

    def set_tree(self, tree: etree._Element) -> None:
        """Replace the tree."""
        self._tree = tree
        self._bytes = None

    def set_bytes(self, data: bytes) -> None:
        """Replace the content with raw XML; parsed on the next access."""
        self._bytes = bytes(data)
        self._tree = None

    def set_xml(self, xml: str) -> None:
        """Replace the content with an XML string."""
        self.set_bytes(xml.encode("utf-8"))

    @property
    def xml(self) -> bytes:
        """The bytes as they will be written."""
        if self._tree is not None:
            from docx4j_py.openpackaging.parts.xml_part import XML_DECLARATION

            return XML_DECLARATION + etree.tostring(self._tree, encoding="utf-8")
        if self._bytes is not None:
            return self._bytes
        data = self._source_bytes()
        if data is None:
            raise Docx4JException(f"Part {self.part_name} has no content and no source container")
        return data

    @property
    def bytes_for_save(self) -> bytes:
        """:attr:`xml`."""
        return self.xml

    def get_xml(self) -> str:
        """:attr:`xml` decoded."""
        return self.xml.decode("utf-8")


class CustomXmlDataStoragePart(DefaultXmlPart):
    """A custom XML data part: the customer's own XML. docx4j ``CustomXmlDataStoragePart``."""

    __slots__ = ("item_id",)

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGE,
    ) -> None:
        """Build the part with the ``customXml`` relationship type."""
        super().__init__(part_name, content_type, Namespaces.CUSTOM_XML_DATA_STORAGE)
        #: The ``ds:itemID`` of the sibling properties part, lower-cased.
        self.item_id: str | None = None


class VMLPart(DefaultXmlPart):
    """A VML drawing, ``/word/vmlDrawing1.vml``. docx4j ``VMLPart``."""

    __slots__ = ()

    def __init__(self, part_name: PartName | str) -> None:
        """Build the part with the ``vmlDrawing`` relationship type."""
        super().__init__(part_name, ContentTypes.VML_DRAWING, Namespaces.VML)
