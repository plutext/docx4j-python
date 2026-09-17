"""``CustomXmlNode``: Office JS ``Word.CustomXmlNode`` over an lxml node.

CR-003 section 3.7. A node in a custom XML part's tree --- an element, one of
its attributes, or a run of text --- with Office JS's member names and the DOM's
semantics, over lxml rather than a DOM implementation, because lxml is what
CR-002 already holds a custom XML part as.

The two things worth knowing:

* **Reading never marks the part.** Every mutation calls
  :meth:`~docx4j_py.model.customxml.parts.CustomXmlPart.touch`, which is what
  makes the part re-marshal; a read leaves it byte for byte (CR-003 section 3.7,
  and :class:`~docx4j_py.openpackaging.parts.default_xml_part.DefaultXmlPart`'s
  parsed-versus-modified split).
* **lxml has no text-node objects**, so a text node is this view over the
  element whose ``text`` or ``tail`` it is. That is enough for
  ``select_nodes("text()")`` to come back as nodes and for their value to be
  read and written, which is what the node model is for.

:class:`CustomXmlPrefixMappingCollection` is Office JS's
``CustomXmlPrefixMappingCollection``: the part's own namespaces, under the
``ns0``, ``ns1``, … names Word uses when it writes a binding.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal

from lxml import etree

from docx4j_py.model.content.errors import BindingError
from docx4j_py.model.customxml.xpath import (
    canonical_xpath_of,
    format_prefix_mappings,
)

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.customxml.parts import CustomXmlPart

__all__ = [
    "CustomXmlNode",
    "CustomXmlPrefixMapping",
    "CustomXmlPrefixMappingCollection",
    "NodeType",
    "node_for",
]

#: Office JS ``Word.CustomXmlNodeType``, in its own spelling.
NodeType = Literal[
    "Element",
    "Attribute",
    "Text",
    "CData",
    "Comment",
    "ProcessingInstruction",
    "Document",
    "Other",
]

class CustomXmlPrefixMapping:
    """One ``prefix`` to ``namespace_uri`` pair. Office JS ``CustomXmlPrefixMapping``."""

    __slots__ = ("namespace_uri", "prefix")

    def __init__(self, prefix: str, namespace_uri: str) -> None:
        """Build the pair."""
        #: The prefix, without the ``xmlns:``.
        self.prefix = prefix
        #: The namespace it stands for.
        self.namespace_uri = namespace_uri

    def to_dict(self) -> dict[str, str]:
        """``{"prefix": …, "namespace_uri": …}``."""
        return {"prefix": self.prefix, "namespace_uri": self.namespace_uri}

    def __eq__(self, other: object) -> bool:
        """Two mappings of the same prefix and URI are equal."""
        return (
            isinstance(other, CustomXmlPrefixMapping)
            and other.prefix == self.prefix
            and other.namespace_uri == self.namespace_uri
        )

    def __hash__(self) -> int:
        """Hashes by the pair."""
        return hash((self.prefix, self.namespace_uri))

    def __repr__(self) -> str:
        """``<CustomXmlPrefixMapping ns0='urn:invoice'>``."""
        return f"<CustomXmlPrefixMapping {self.prefix}={self.namespace_uri!r}>"


class CustomXmlPrefixMappingCollection:
    """The prefixes a part's XPaths may use. Office JS ``CustomXmlPrefixMappingCollection``.

    Seeded from the namespaces the document element declares, under Word's own
    ``ns0``, ``ns1``, … convention where a declaration has no prefix of its own:
    XPath 1.0 has no default namespace, so an unprefixed ``xmlns`` would be
    unusable as it stands.
    """

    __slots__ = ("_mappings",)

    def __init__(self, mappings: dict[str, str] | None = None) -> None:
        """Build the collection from a prefix-to-URI mapping."""
        self._mappings: dict[str, str] = dict(mappings or {})

    @property
    def items(self) -> list[CustomXmlPrefixMapping]:
        """Every mapping, in the order they were added."""
        return [
            CustomXmlPrefixMapping(prefix, uri) for prefix, uri in self._mappings.items() if prefix
        ]

    def add_namespace(self, prefix: str, namespace_uri: str) -> None:
        """Add or replace a prefix. Office JS ``addNamespace``."""
        self._mappings[prefix] = namespace_uri

    def lookup_namespace(self, prefix: str) -> str:
        """The URI a prefix stands for, or ``""``. Office JS ``lookupNamespace``."""
        return self._mappings.get(prefix, "")

    def lookup_prefix(self, namespace_uri: str) -> str:
        """The first prefix that stands for a URI, or ``""``. Office JS ``lookupPrefix``."""
        for prefix, uri in self._mappings.items():
            if uri == namespace_uri and prefix:
                return prefix
        return ""

    def to_map(self) -> dict[str, str]:
        """A copy as a plain mapping, for :func:`lxml.etree._Element.xpath`. Extension."""
        return dict(self._mappings)

    def to_string(self) -> str:
        """Word's ``xmlns:ns0='…'`` spelling of the whole collection. Extension."""
        return format_prefix_mappings(self._mappings)

    def __len__(self) -> int:
        """How many prefixes there are."""
        return len([prefix for prefix in self._mappings if prefix])

    def __iter__(self) -> Any:
        """Iterates the mappings."""
        return iter(self.items)

    def __repr__(self) -> str:
        """``<CustomXmlPrefixMappingCollection 1 prefix>``."""
        count = len(self)
        return f"<CustomXmlPrefixMappingCollection {count} prefix{'es' if count != 1 else ''}>"


def namespaces_of(element: Any) -> dict[str, str]:
    """Every namespace in scope at an element, under Word's ``ns0`` convention.

    lxml's ``nsmap`` keys a default declaration with ``None``; XPath cannot use
    one, so it is given the first free ``nsN`` name, which is what Word writes
    when it binds a control to a node in a namespaced part.
    """
    if element is None:
        return {}
    out: dict[str, str] = {}
    unnamed: list[str] = []
    for prefix, uri in (getattr(element, "nsmap", None) or {}).items():
        if prefix is None:
            unnamed.append(uri)
        else:
            out[prefix] = uri
    index = 0
    for uri in unnamed:
        if uri in out.values():
            continue
        while f"ns{index}" in out:
            index += 1
        out[f"ns{index}"] = uri
    return out


def node_for(result: Any, part: CustomXmlPart) -> CustomXmlNode | None:
    """Wrap one lxml XPath result as a node, or None for a value.

    lxml answers an attribute or a ``text()`` step with a "smart string" that
    knows where it came from, which is how an attribute and a text node get
    their identity here.
    """
    if result is None:
        return None
    if isinstance(result, str):
        owner = getattr(result, "getparent", lambda: None)()
        if owner is None:
            return None
        if getattr(result, "is_attribute", False):
            return CustomXmlNode(owner, part, attribute=result.attrname)
        if getattr(result, "is_tail", False):
            return CustomXmlNode(owner, part, text_kind="tail")
        return CustomXmlNode(owner, part, text_kind="text")
    if isinstance(result, (float, bool, int)):
        return None
    return CustomXmlNode(result, part)


class CustomXmlNode:
    """A node of a custom XML part. Office JS ``Word.CustomXmlNode``."""

    __slots__ = ("attribute", "element", "owner_part", "text_kind")

    def __init__(
        self,
        element: Any,
        owner_part: CustomXmlPart,
        *,
        attribute: str | None = None,
        text_kind: str | None = None,
    ) -> None:
        """Build the view over an element, one of its attributes, or its text.

        Args:
            element: the lxml element; for an attribute or a text node, the
                element it belongs to.
            owner_part: the part, which is what a mutation marks.
            attribute: the attribute's qualified name (``{uri}local``).
            text_kind: ``"text"`` or ``"tail"`` for a text node.
        """
        #: The lxml element this node is, or the one it belongs to.
        self.element = element
        #: The part this node lives in. Office JS ``ownerPart``.
        self.owner_part = owner_part
        #: The attribute's qualified name, or None.
        self.attribute = attribute
        #: ``"text"``, ``"tail"`` or None.
        self.text_kind = text_kind

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same node are equal (CR-003 section 3.1)."""
        return (
            isinstance(other, CustomXmlNode)
            and other.element is self.element
            and other.attribute == self.attribute
            and other.text_kind == self.text_kind
        )

    def __hash__(self) -> int:
        """Hashes by the node's identity."""
        return hash((id(self.element), self.attribute, self.text_kind))

    def __str__(self) -> str:
        """The node's text."""
        return self.text

    def __repr__(self) -> str:
        """``<CustomXmlNode /invoice[1]/customer[1]/name[1] 'Acme Ltd'>``."""
        text = self.text.replace("\n", " ")
        preview = text[:30] + "…" if len(text) > 30 else text
        return f"<CustomXmlNode {self.xpath} {preview!r}>"

    # -- what it is --------------------------------------------------------

    @property
    def node_type(self) -> str:
        """``"Element"``, ``"Attribute"``, ``"Text"``, ``"Comment"``, ….

        Office JS ``nodeType``, in its own spelling.
        """
        if self.attribute is not None:
            return "Attribute"
        if self.text_kind is not None:
            return "Text"
        tag = getattr(self.element, "tag", None)
        if tag is etree.Comment:
            return "Comment"
        if tag is etree.ProcessingInstruction:
            return "ProcessingInstruction"
        if isinstance(tag, str):
            return "Element"
        return "Other"

    @property
    def base_name(self) -> str:
        """The local name, without a prefix. Office JS ``baseName``."""
        name = self.attribute if self.attribute is not None else getattr(self.element, "tag", "")
        if self.text_kind is not None or not isinstance(name, str):
            return ""
        return name.rpartition("}")[2] if name.startswith("{") else name

    @property
    def namespace_uri(self) -> str:
        """The namespace, or ``""``. Office JS ``namespaceUri``."""
        name = self.attribute if self.attribute is not None else getattr(self.element, "tag", "")
        if self.text_kind is not None or not isinstance(name, str) or not name.startswith("{"):
            return ""
        return name[1:].partition("}")[0]

    @property
    def xpath(self) -> str:
        """The node's canonical path, Word's fully positional form. Office JS ``xpath``."""
        return self._canonical().xpath

    @property
    def prefix_mappings(self) -> str:
        """The ``xmlns:ns0='…'`` string :attr:`xpath` needs. Extension."""
        return self._canonical().prefix_mappings

    def _canonical(self) -> Any:
        if self.attribute is not None:
            return canonical_xpath_of((self.element, self.attribute))
        return canonical_xpath_of(self.element)

    # -- the value ---------------------------------------------------------

    @property
    def text(self) -> str:
        """Everything below this node as text. Office JS ``text``.

        The DOM's ``textContent``: an attribute's value, a text node's string,
        and for an element every descendant's text run together.
        """
        if self.attribute is not None:
            return self.element.get(self.attribute) or ""
        if self.text_kind is not None:
            return (getattr(self.element, self.text_kind, None)) or ""
        return "".join(self.element.itertext())

    @text.setter
    def text(self, value: str) -> None:
        """Replace the node's text; the part is marked and a report is recorded."""
        value = "" if value is None else str(value)
        with self.owner_part.recording("custom_xml_node.text") as change:
            change.touched(self.xpath)
            change.text(before=self.text, after=value)
            self._write_text(value)

    def _write_text(self, value: str) -> None:
        """The write itself, inside :attr:`text`'s report."""
        if self.attribute is not None:
            self.element.set(self.attribute, value)
        elif self.text_kind is not None:
            setattr(self.element, self.text_kind, value)
        else:
            for child in list(self.element):
                self.element.remove(child)
            self.element.text = value
        self.owner_part.touch()

    @property
    def node_value(self) -> str:
        """The DOM's ``nodeValue``: :attr:`text` for every node kind here."""
        return self.text

    @node_value.setter
    def node_value(self, value: str) -> None:
        """As :attr:`text`, under its own operation name."""
        with self.owner_part.recording("custom_xml_node.node_value"):
            self.text = value

    @property
    def xml(self) -> str:
        """The node as XML. Office JS ``xml``."""
        return self.get_xml()

    def get_xml(self) -> str:
        """The node as XML, with no declaration. Extension (Office JS has ``xml``)."""
        if self.attribute is not None:
            name = self.base_name
            value = (self.text or "").replace("&", "&amp;").replace("<", "&lt;")
            return f'{name}="{value.replace(chr(34), "&quot;")}"'
        if self.text_kind is not None:
            return self.text
        return etree.tostring(self.element, encoding="unicode")

    def set_xml(self, xml: str) -> None:
        """Replace this element with the one `xml` describes. Extension.

        Raises:
            BindingError: `xml` is not one well-formed element, or this node is
                the document element (replace the part with ``set_xml``).
        """
        with self.owner_part.recording("custom_xml_node.set_xml") as change:
            self._write_xml(xml, change)

    def _write_xml(self, xml: str, change: Any) -> None:
        """The write itself, inside :meth:`set_xml`'s report."""
        parent = self.element.getparent() if self.attribute is None else None
        if self.attribute is not None or self.text_kind is not None:
            self.text = xml
            return
        if parent is None:
            raise BindingError(
                "this is the document element; replacing it replaces the part",
                code="binding.document_element",
                hint="call part.set_xml(xml) to replace the whole part",
            )
        created = _parse_element(xml)
        parent.replace(self.element, created)
        self.element = created
        self.owner_part.touch()
        change.touched(self.xpath)

    # -- the tree ----------------------------------------------------------

    @property
    def parent_node(self) -> CustomXmlNode | None:
        """The node above, or None for the document element. Office JS ``parentNode``."""
        if self.attribute is not None or self.text_kind == "text":
            return CustomXmlNode(self.element, self.owner_part)
        parent = self.element.getparent()
        if parent is None:
            return None
        return CustomXmlNode(parent, self.owner_part)

    @property
    def child_nodes(self) -> list[CustomXmlNode]:
        """Every child, text nodes included. Office JS ``childNodes``."""
        if self.attribute is not None or self.text_kind is not None:
            return []
        out: list[CustomXmlNode] = []
        if self.element.text:
            out.append(CustomXmlNode(self.element, self.owner_part, text_kind="text"))
        for child in self.element:
            out.append(CustomXmlNode(child, self.owner_part))
            if child.tail:
                out.append(CustomXmlNode(child, self.owner_part, text_kind="tail"))
        return out

    @property
    def child_elements(self) -> list[CustomXmlNode]:
        """The element children alone, which is what a data store is read for. Extension."""
        if self.attribute is not None or self.text_kind is not None:
            return []
        return [
            CustomXmlNode(child, self.owner_part)
            for child in self.element
            if isinstance(getattr(child, "tag", None), str)
        ]

    @property
    def attributes(self) -> list[CustomXmlNode]:
        """This element's attributes, as nodes. Office JS ``attributes``."""
        if self.attribute is not None or self.text_kind is not None:
            return []
        return [
            CustomXmlNode(self.element, self.owner_part, attribute=name)
            for name in self.element.attrib
        ]

    @property
    def first_child(self) -> CustomXmlNode | None:
        """The first child, or None. Office JS ``firstChild``."""
        children = self.child_nodes
        return children[0] if children else None

    @property
    def last_child(self) -> CustomXmlNode | None:
        """The last child, or None. Office JS ``lastChild``."""
        children = self.child_nodes
        return children[-1] if children else None

    @property
    def next_sibling(self) -> CustomXmlNode | None:
        """The node after this one, or None. Office JS ``nextSibling``."""
        return self._sibling(1)

    @property
    def previous_sibling(self) -> CustomXmlNode | None:
        """The node before this one, or None. Office JS ``previousSibling``."""
        return self._sibling(-1)

    def _sibling(self, step: int) -> CustomXmlNode | None:
        parent = self.parent_node
        if parent is None or self.attribute is not None:
            return None
        siblings = parent.child_nodes
        for index, node in enumerate(siblings):
            if node == self:
                position = index + step
                if 0 <= position < len(siblings):
                    return siblings[position]
                return None
        return None

    def has_child_nodes(self) -> bool:
        """Whether this node has any children. Office JS ``hasChildNodes``."""
        return bool(self.child_nodes)

    # -- selecting ---------------------------------------------------------

    def select_nodes(
        self, xpath: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> list[CustomXmlNode]:
        """Every node an XPath selects, **relative to this one**.

        Office JS ``selectNodes``. `namespace_mappings` is Word's
        ``xmlns:ns0='urn:invoice'`` string, a mapping, or None for the part's
        own namespace manager.
        """
        return self.owner_part.select_from(self.element, xpath, namespace_mappings)

    def select_single_node(
        self, xpath: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> CustomXmlNode | None:
        """The first node an XPath selects, or None. Office JS ``selectSingleNode``."""
        found = self.select_nodes(xpath, namespace_mappings)
        return found[0] if found else None

    # -- editing -----------------------------------------------------------

    def append_child_node(
        self,
        xml: str | None = None,
        namespace_uri: str | None = None,
        node_type: str | None = None,
        node_value: str | None = None,
        index: int | None = None,
    ) -> CustomXmlNode:
        """Add a child, from XML or from Office JS's four arguments.

        ``append_child_node("<price>20</price>")`` takes the XML of the node;
        ``append_child_node("price", "", "Element", "20")`` takes Office JS's
        name, namespace, type and value. `index` counts **element** children, as
        Office JS's does, and past the end appends.

        Raises:
            BindingError: neither form was given what it needs.
        """
        with self.owner_part.recording("custom_xml_node.append_child_node") as change:
            out = self._append(xml, namespace_uri, node_type, node_value, index)
            change.touched(out.xpath)
            change.text(after=out.text)
            return out

    def _append(
        self,
        xml: str | None,
        namespace_uri: str | None,
        node_type: str | None,
        node_value: str | None,
        index: int | None,
    ) -> CustomXmlNode:
        """The write itself, inside :meth:`append_child_node`'s report."""
        created = self._create(xml, namespace_uri, node_type, node_value)
        if isinstance(created, tuple):
            name, value = created
            self.element.set(name, value)
            self.owner_part.touch()
            return CustomXmlNode(self.element, self.owner_part, attribute=name)
        if isinstance(created, str):  # a text node
            self.element.text = (self.element.text or "") + created
            self.owner_part.touch()
            return CustomXmlNode(self.element, self.owner_part, text_kind="text")
        children = list(self.element)
        if index is None or index >= len(children):
            self.element.append(created)
        else:
            self.element.insert(max(0, index), created)
        self.owner_part.touch()
        return CustomXmlNode(created, self.owner_part)

    def insert_node_before(
        self, xml: str, next_sibling: CustomXmlNode | None = None
    ) -> CustomXmlNode:
        """Insert a node before a child of this one, or append. Office JS ``insertNodeBefore``."""
        with self.owner_part.recording("custom_xml_node.insert_node_before") as change:
            created = _parse_element(xml)
            if next_sibling is None:
                self.element.append(created)
            else:
                self.element.insert(self.element.index(next_sibling.element), created)
            self.owner_part.touch()
            out = CustomXmlNode(created, self.owner_part)
            change.touched(out.xpath)
            change.text(after=out.text)
            return out

    def remove_child(self, child: CustomXmlNode) -> None:
        """Remove a child node or an attribute. Office JS ``removeChild``."""
        with self.owner_part.recording("custom_xml_node.remove_child") as change:
            change.touched(child.xpath)
            change.text(before=child.text)
            self._remove(child)

    def _remove(self, child: CustomXmlNode) -> None:
        """The removal itself, inside :meth:`remove_child`'s report."""
        if child.attribute is not None:
            self.element.attrib.pop(child.attribute, None)
        elif child.text_kind is not None:
            setattr(child.element, child.text_kind, None)
        else:
            tail = child.element.tail
            self.element.remove(child.element)
            if tail:
                # lxml carries a removed element's tail away with it; the DOM
                # leaves that text where it was
                previous = child.element.getprevious()
                if previous is not None:
                    previous.tail = (previous.tail or "") + tail
                else:
                    self.element.text = (self.element.text or "") + tail
        self.owner_part.touch()

    #: Office JS calls it ``removeChild``; the TypeScript engine's node model
    #: spells it ``removeChildNode`` as well. Both are here.
    remove_child_node = remove_child

    def replace_child_node(self, old_node: CustomXmlNode, xml: str) -> CustomXmlNode:
        """Replace a child with the node `xml` describes. Office JS ``replaceChildNode``."""
        with self.owner_part.recording("custom_xml_node.replace_child_node") as change:
            created = _parse_element(xml)
            change.text(before=old_node.text)
            self.element.replace(old_node.element, created)
            self.owner_part.touch()
            out = CustomXmlNode(created, self.owner_part)
            change.touched(out.xpath)
            change.text(after=out.text)
            return out

    def delete(self) -> None:
        """Remove this node from its parent. Office JS ``delete``.

        Raises:
            BindingError: this is the document element, which only
                :meth:`~docx4j_py.model.customxml.CustomXmlPart.delete` removes.
        """
        with self.owner_part.recording("custom_xml_node.delete"):
            self._delete()

    def _delete(self) -> None:
        """The removal itself, inside :meth:`delete`'s report."""
        parent = self.parent_node
        if parent is None:
            raise BindingError(
                "the document element cannot be deleted; delete the part instead",
                code="binding.document_element",
                hint="call part.delete(), or replace the content with part.set_xml(xml)",
            )
        parent.remove_child(self)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a node."""
        return {
            "xpath": self.xpath,
            "prefix_mappings": self.prefix_mappings,
            "base_name": self.base_name,
            "namespace_uri": self.namespace_uri,
            "node_type": self.node_type,
            "text": self.text,
        }

    # -- helpers -----------------------------------------------------------

    def _create(
        self,
        xml: str | None,
        namespace_uri: str | None,
        node_type: str | None,
        node_value: str | None,
    ) -> Any:
        """The element, ``(name, value)`` attribute pair or text a caller asked for."""
        if node_type is None and namespace_uri is None:
            if xml is None:
                raise BindingError(
                    "pass the XML of the node, or its name, namespace, type and value",
                    code="binding.node_undescribed",
                    hint="append_child_node('<price>20</price>')",
                )
            return _parse_element(xml)
        if not xml:
            raise BindingError(
                "a node needs a name",
                code="binding.node_unnamed",
                hint="append_child_node('price', '', 'Element', '20')",
            )
        qualified = f"{{{namespace_uri}}}{xml}" if namespace_uri else xml
        if node_type == "Attribute":
            return (qualified, node_value or "")
        if node_type == "Text":
            return node_value or ""
        created = etree.Element(qualified)
        if node_value:
            created.text = node_value
        return created


def _parse_element(xml: str) -> Any:
    """One well-formed element out of an XML string.

    Raises:
        BindingError: it is not well formed, or not one element.
    """
    try:
        return etree.fromstring(xml)
    except Exception as error:
        raise BindingError(
            f"not a well-formed XML element: {xml!r}",
            code="binding.not_well_formed",
            hint="pass one element, namespaces declared on it if it needs them",
        ) from error


def require_node(
    found: list[CustomXmlNode], xpath: str, where: str, *, one: bool = True
) -> CustomXmlNode:
    """The one node an XPath selected, or a :class:`BindingError` saying why not.

    Used by everything that edits through an XPath: CR-003 section 3.7 asks for
    an error that names the expression, because that is what tells a caller
    which binding is wrong.
    """
    if not found:
        raise BindingError(
            f"{xpath!r} selects nothing in {where}",
            code="binding.no_match",
            hint="call select_nodes() to see what the part holds",
        )
    if one and len(found) > 1:
        raise BindingError(
            f"{xpath!r} selects {len(found)} nodes in {where}",
            code="binding.many_matches",
            hint=f"add a positional predicate, as in {found[0].xpath!r}",
        )
    return found[0]
