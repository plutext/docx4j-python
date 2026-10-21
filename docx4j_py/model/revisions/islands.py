"""Text boxes in a lossless tree, as typed islands. CR-010 section 3.5.

docx4j keeps both branches of ``mc:AlternateContent`` as typed JAXB. This
package keeps them, and the DrawingML and VML shapes a text box lives in, as
**wildcard** ``AnyElement`` trees (CR-002's lossless MCE): the ``w:txbxContent``
of a text box is reached through ``wps:txbx`` or ``v:textbox``, elements this
build has no classes for, so everything below them is a wildcard too. A typed
engine cannot see the revisions in there --- which is why CR-003 Phase F's
listing missed them.

An **island** is one such ``w:txbxContent``, made typed for as long as it is
being resolved:

1. the wildcard subtree is written out as XML (with docx4j's prefixes) and
   parsed into ``TxbxContent`` by the runtime parser, as ``wml(...)`` parses a
   fragment --- strictly: anything skipped raises, so nothing is lost;
2. it is resolved like any other story;
3. it is written back in place: serialised, and parsed again as a wildcard
   tree by the same parser path that made the original (inside an element no
   class is bound to, so that ``w:txbxContent`` stays a wildcard), and the new
   node takes the old one's place and ``tail``. The rest of the branch is not
   touched at all: not re-parsed, not re-serialised.

A text box with no revision markup in it is never made an island, so a
document whose text boxes nobody edited keeps them exactly as they were.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from docx4j_py.child import MC_CHOICE, is_any_element, link_parents, source_bindings
from docx4j_py.namespaces import PREFIXES, WML_NS

__all__ = ["REVISION_LOCALS", "Island", "holds_revisions"]

_TXBX = f"{{{WML_NS}}}txbxContent"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_WRAPPER_NS = "urn:docx4j-python:revisions:island"

#: The WordprocessingML elements that are revision markup, by local name: a
#: wildcard text box holding none of them is left alone.
REVISION_LOCALS: frozenset[str] = frozenset(
    {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "moveFromRangeStart",
        "moveToRangeStart",
        "rPrChange",
        "pPrChange",
        "sectPrChange",
        "tblPrChange",
        "tblGridChange",
        "tblPrExChange",
        "trPrChange",
        "tcPrChange",
        "cellIns",
        "cellDel",
        "cellMerge",
        "numberingChange",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
    }
)

_REVISION_QNAMES: frozenset[str] = frozenset(f"{{{WML_NS}}}{local}" for local in REVISION_LOCALS)

_NSMAP = {prefix: uri for prefix, uri in PREFIXES.items() if prefix != "xml"}


def holds_revisions(node: Any) -> bool:
    """True when a wildcard subtree holds any revision markup."""
    stack = [node]
    while stack:
        current = stack.pop()
        if is_any_element(current):
            if current.qname in _REVISION_QNAMES:
                return True
            stack.extend(c for c in current.children or () if not isinstance(c, str))
        elif not isinstance(current, str):
            # a typed object inside a wildcard: ask the serialised form
            from docx4j_py.fragments import to_xml

            try:
                xml = to_xml(current)
            except Exception:  # noqa: BLE001 - cannot tell: assume it may
                return True
            if any(f":{local}" in xml for local in REVISION_LOCALS):
                return True
    return False


def _to_lxml(node: Any, parent: etree._Element | None = None) -> etree._Element:
    """A wildcard subtree as an lxml element, its attributes, text and tails kept.

    An ``mc:Choice`` is declared what its ``Requires`` (and its own
    ``mc:Ignorable`` and ``mc:MustUnderstand``) were bound to in its source
    (:func:`~docx4j_py.child.source_bindings`), so that the typed parse
    records the same bindings and selects the same branch (CR-011 Q3). Made
    with ``SubElement``, which keeps a declaration whose namespace is in scope
    under another prefix (``shape2010``), where ``append`` would strip it.
    """
    if not is_any_element(node):
        from docx4j_py.fragments import to_xml

        element = etree.fromstring(to_xml(node).encode("utf-8"))
        if parent is not None:
            parent.append(element)
        return element
    nsmap: dict[str, str] | None = None
    if node.qname == MC_CHOICE:
        bound = {prefix: uri for prefix, uri in source_bindings(node).items() if uri is not None}
        nsmap = bound or None
    if parent is None:
        element = etree.Element(node.qname, nsmap={**_NSMAP, **(nsmap or {})})
    else:
        element = etree.SubElement(parent, node.qname, nsmap=nsmap)
    for key, value in (node.attributes or {}).items():
        if value is not None:
            element.set(key, str(value))
    element.text = node.text
    for child in node.children or ():
        if isinstance(child, str):
            continue
        _to_lxml(child, element)
    if parent is not None:
        element.tail = node.tail
    return element


class Island:
    """A wildcard ``w:txbxContent`` made typed, and where to write it back."""

    __slots__ = ("dirty", "holder", "node", "typed")

    def __init__(self, node: Any, holder: list) -> None:
        """Parse the wildcard `node`, which stands in the list `holder`.

        Raises:
            ValueError: the subtree would lose content parsed as ``TxbxContent``.
        """
        from docx4j_py.runtime import context, parser
        from docx4j_py.wml import TxbxContent

        #: The wildcard node, as it stands in the tree.
        self.node = node
        #: The wildcard list it stands in (its parent's ``children``).
        self.holder = holder
        #: True once resolving has changed the typed tree.
        self.dirty = False
        data = etree.tostring(_to_lxml(node))
        xml_parser = parser(lenient=True, skipped_report=True)
        typed = xml_parser.from_bytes(data, TxbxContent)
        lost = list(getattr(xml_parser.skipped, "items", ()) or ())
        if lost:
            raise ValueError(f"a text box would lose {len(lost)} item(s) as typed content: {lost[:3]}")
        link_parents(typed, context=context())
        #: The typed ``w:txbxContent`` the story is over.
        self.typed = typed

    def write_back(self) -> None:
        """Put the typed content back into the wildcard tree, in place; the rest of the branch untouched."""
        if not self.dirty:
            return
        from docx4j_py.fragments import to_xml
        from docx4j_py.mce import AlternateContentFallback
        from docx4j_py.runtime import parser

        xml = to_xml(self.typed, name=_TXBX)
        # inside an element no class is bound to, the parser keeps w:txbxContent
        # a wildcard, exactly as it did when it read the document
        document = (
            f'<mc:Fallback xmlns:mc="{_MC}"><x:island xmlns:x="{_WRAPPER_NS}">'
            f"{xml}</x:island></mc:Fallback>"
        )
        wrapper = parser(lenient=True, skipped_report=True).from_bytes(
            document.encode("utf-8"), AlternateContentFallback
        )
        fresh = wrapper.any_element[0].children[0]
        fresh.tail = self.node.tail
        for index, item in enumerate(self.holder):
            if item is self.node:
                self.holder[index] = fresh
                break
        self.node = fresh
        self.dirty = False


def is_wildcard_txbx(node: Any) -> bool:
    """True for a ``w:txbxContent`` kept as a wildcard."""
    return is_any_element(node) and node.qname == _TXBX
