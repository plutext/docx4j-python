"""``XmlMapping``: the ``w:dataBinding`` of a content control.

CR-003 section 3.7. Office JS's ``Word.XmlMapping``, hung off every
:class:`~docx4j_py.model.content.ContentControl`, over the ``w:dataBinding``
Word writes when a control is bound to a node of a custom XML part.

Two things the TypeScript engine learned and CR-003 section 4 makes rules:

* the binding is read from ``w:dataBinding`` **and** ``w15:dataBinding`` ---
  Word writes the second on a repeating section and on a rich-text control bound
  to a container, and 3 of the 20 bindings in ``samples/invoice2013.docx`` are
  that spelling;
* :meth:`XmlMapping.set_mapping` **resolves before it writes**: an XPath that
  selects nothing in any candidate part leaves the control's binding exactly as
  it was and answers False, which is what Word does.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.model.customxml.parts import normalise_id
from docx4j_py.model.customxml.xpath import canonical_xpath_of, format_prefix_mappings
from docx4j_py.wml import el
from docx4j_py.wml.sdt import W15_NS, W_NS, sdt_property

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.controls import ContentControl
    from docx4j_py.model.customxml.nodes import CustomXmlNode
    from docx4j_py.model.customxml.parts import CustomXmlPart

__all__ = ["XmlMapping"]


class XmlMapping:
    """A content control's binding to a custom XML node. Office JS ``Word.XmlMapping``."""

    __slots__ = ("control",)

    def __init__(self, control: ContentControl) -> None:
        """Build the view over a control's ``w:dataBinding``."""
        #: The control this is the mapping of.
        self.control = control

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two mappings of the same control are equal."""
        return isinstance(other, XmlMapping) and other.control == self.control

    def __hash__(self) -> int:
        """Hashes by the control."""
        return hash(self.control)

    def __repr__(self) -> str:
        """``<XmlMapping '/invoice[1]/total[1]' in {5D7B…}>``."""
        if not self.is_mapped:
            return "<XmlMapping unmapped>"
        return f"<XmlMapping {self.xpath!r} in {self.store_item_id}>"

    # -- the binding -------------------------------------------------------

    @property
    def data_binding(self) -> Any:
        """The ``w:dataBinding`` or ``w15:dataBinding`` element, or None. Extension."""
        return sdt_property(self.control.sdt_pr, "dataBinding", (W_NS, W15_NS))

    @property
    def is_mapped(self) -> bool:
        """Whether the control carries a binding at all. Office JS ``isMapped``."""
        return self.data_binding is not None

    @property
    def xpath(self) -> str:
        """The bound XPath, or ``""``. Office JS ``xpath``."""
        return getattr(self.data_binding, "xpath", None) or ""

    @property
    def prefix_mappings(self) -> str:
        """The ``xmlns:ns0='…'`` string the XPath needs, or ``""``. Office JS ``prefixMappings``."""
        return getattr(self.data_binding, "prefix_mappings", None) or ""

    @property
    def store_item_id(self) -> str:
        """The ``w:storeItemID`` of the part the binding names, or ``""``."""
        return getattr(self.data_binding, "store_item_id", None) or ""

    @property
    def custom_xml_part(self) -> CustomXmlPart | None:
        """The part the binding names, or None when the package has no such part.

        The three well-known docProps store item ids are **not** special-cased
        (CR-003 section 4): a binding to one resolves only when the document
        really carries that custom XML part.
        """
        collection = self._collection()
        item_id = self.store_item_id
        if collection is None or not item_id:
            return None
        return collection.get_item(item_id)

    @property
    def custom_xml_node(self) -> CustomXmlNode | None:
        """The node the binding resolves to, or None. Office JS ``customXmlNode``.

        Evaluated on every read, as Word does: an edit to the data is visible
        through the same mapping straight away.
        """
        part = self.custom_xml_part
        if part is None or not self.xpath:
            return None
        return part.select_single_node(self.xpath, self.prefix_mappings)

    # -- writing it --------------------------------------------------------

    def set_mapping(
        self,
        xpath: str,
        prefix_mappings: str | None = None,
        part: CustomXmlPart | None = None,
    ) -> bool:
        """Bind the control to an XPath. Office JS ``setMapping``.

        The candidates are `part` alone when one is given, and otherwise the
        control's current part first, then the customer's own parts, then Word's
        property stores last. The first that **resolves** the XPath is written;
        if none does, nothing is written and this answers False, which is what
        Word's XML Mapping pane does with a path that matches no node.

        Returns:
            Whether a binding was written.
        """
        for candidate in [part] if part is not None else self._candidates():
            if candidate is None:
                continue
            if candidate.select_single_node(xpath, prefix_mappings) is None:
                continue
            self._write(xpath, prefix_mappings or "", candidate.id)
            return True
        return False

    def set_mapping_by_node(self, node: CustomXmlNode) -> bool:
        """Bind the control to a node, by its canonical XPath. Office JS ``setMappingByNode``.

        Always succeeds: the node is there, so its own path resolves.
        """
        if node.attribute is not None:
            canonical = canonical_xpath_of((node.element, node.attribute))
        else:
            canonical = canonical_xpath_of(node.element)
        self._write(canonical.xpath, canonical.prefix_mappings, node.owner_part.id)
        return True

    def delete(self) -> None:
        """Remove the binding, keeping the content the control shows. Office JS ``delete``."""
        self.control.remove_property("dataBinding", (W_NS, W15_NS))

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a mapping."""
        return {
            "is_mapped": self.is_mapped,
            "xpath": self.xpath,
            "prefix_mappings": self.prefix_mappings,
            "store_item_id": self.store_item_id,
            "resolves": self.custom_xml_node is not None,
        }

    # -- helpers -----------------------------------------------------------

    def _collection(self) -> Any:
        from docx4j_py.model.customxml.parts import custom_xml_parts_of

        package = getattr(self.control.parent_body, "package", None)
        return custom_xml_parts_of(package) if package is not None else None

    def _candidates(self) -> list[CustomXmlPart]:
        collection = self._collection()
        if collection is None:
            return []
        current = self.custom_xml_part
        rest = [part for part in collection.items if part is not current]
        out: list[CustomXmlPart] = [current] if current is not None else []
        out.extend(part for part in rest if not part.built_in)
        out.extend(part for part in rest if part.built_in)
        return out

    def _write(self, xpath: str, prefix_mappings: str, store_item_id: str) -> None:
        """Write the binding, in the namespace the control already uses.

        A control Word bound with ``w15:dataBinding`` --- a repeating section, a
        rich-text control bound to a container --- keeps that spelling, because
        replacing it with ``w:dataBinding`` would leave Word with two bindings
        or a container it no longer repeats. Everything else gets
        ``w:dataBinding``, which is what Word writes for a text control.
        """
        existing = self.data_binding
        namespace = W_NS
        if existing is not None:
            from docx4j_py.traversal import element_name

            name = element_name(existing) or ""
            if name.startswith(f"{{{W15_NS}}}"):
                namespace = W15_NS
        self.control.remove_property("dataBinding", (W_NS, W15_NS))
        if namespace == W15_NS:
            from docx4j_py.w15 import el as w15_el

            binding = w15_el.dataBinding(xpath=xpath, store_item_id=store_item_id)
        else:
            binding = el.dataBinding(xpath=xpath, store_item_id=store_item_id)
        if prefix_mappings:
            binding.prefix_mappings = prefix_mappings
        self.control.put_property(binding)


def mappings_string(mappings: dict[str, str]) -> str:
    """Word's spelling of a prefix mapping; re-exported for the template module."""
    return format_prefix_mappings(mappings)


def same_part(store_item_id: str, part: CustomXmlPart) -> bool:
    """Whether a ``w:storeItemID`` names a part, braces and case aside."""
    return normalise_id(store_item_id) == normalise_id(part.id)
