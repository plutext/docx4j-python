"""A data-bound control's bound value (WD REQ-509). CR-011 section 6.

docx4j's ``org.docx4j.model.text.BoundValues``: ``BindingTraverser``'s rule,
read-only. The binding's XPath is evaluated, with its own prefix mappings,
against the custom XML part its ``w:storeItemID`` names (through the content
API's :mod:`docx4j_py.model.customxml`, whose parts parse without being
modified) or against the core or application properties by their well-known
store item ids (their bytes as they would be saved, parsed privately); the
value is the string value of the first node selected. A binding resolves when
the part exists and the XPath selects a node; one that does not resolve gives
None, and the control's surface is read instead.

The content API's ``XmlMapping.custom_xml_part`` deliberately does not
special-case the two property stores (CR-003 section 4); docx4j's reading of
a bound value does, because Word shows them, so this does too.
"""

from __future__ import annotations

from typing import Any

__all__ = ["CORE_PROPERTIES_STORE_ITEM_ID", "EXTENDED_PROPERTIES_STORE_ITEM_ID", "BoundValues"]

#: docx4j ``BindingHandler.CORE_PROPERTIES_STOREITEMID``.
CORE_PROPERTIES_STORE_ITEM_ID = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"
#: docx4j ``BindingHandler.EXTENDED_PROPERTIES_STOREITEMID``.
EXTENDED_PROPERTIES_STORE_ITEM_ID = "{6668398D-A668-4E3E-A5EB-62B293D839F1}"


class BoundValues:
    """Bound values of one package, each store's tree read once."""

    __slots__ = ("_roots", "package")

    def __init__(self, package: Any) -> None:
        """A reader over a package."""
        self.package = package
        self._roots: dict[str, Any] = {}

    def value(self, binding: Any) -> str | None:
        """The bound value of a ``w:dataBinding``, or None where the binding does not resolve."""
        xpath = getattr(binding, "xpath", None) if binding is not None else None
        store = getattr(binding, "store_item_id", None) if binding is not None else None
        if self.package is None or not xpath or not store:
            return None
        root = self._root(store)
        if root is None:
            return None
        from docx4j_py.model.customxml.xpath import parse_prefix_mappings, select

        try:
            nodes = select(root, xpath, parse_prefix_mappings(getattr(binding, "prefix_mappings", None)))
        except Exception:  # noqa: BLE001 - a binding that will not evaluate does not resolve
            return None
        if not nodes:
            return None
        node = nodes[0]
        if isinstance(node, str):
            return str(node)  # an attribute, a text node or a string result
        if isinstance(node, (bool, float)):
            return None
        return "".join(node.itertext())

    def _root(self, store: str) -> Any:
        key = store.upper()
        if key in self._roots:
            return self._roots[key]
        root = None
        try:
            if key == CORE_PROPERTIES_STORE_ITEM_ID:
                root = _parse(getattr(self.package, "doc_props_core_part", None))
            elif key == EXTENDED_PROPERTIES_STORE_ITEM_ID:
                root = _parse(getattr(self.package, "doc_props_extended_part", None))
            else:
                from docx4j_py.model.customxml.parts import custom_xml_parts_of

                part = custom_xml_parts_of(self.package).get_item(store)
                root = part.tree if part is not None else None
        except Exception:  # noqa: BLE001 - a store that will not read resolves nothing
            root = None
        self._roots[key] = root
        return root


def _parse(part: Any) -> Any:
    """A properties part's tree, from its bytes as they would be saved: nothing is unmarshalled."""
    if part is None:
        return None
    from lxml import etree

    parser = etree.XMLParser(resolve_entities=False, huge_tree=True)
    return etree.fromstring(part.xml, parser)
