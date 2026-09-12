"""XML fragments in and out: ``wml(...)`` and ``to_xml(...)``.

CR-001 section 6.2, the analogue of docx4j's ``XmlUtils.unmarshalString`` with
``W_NAMESPACE_DECLARATION`` and of ``builders/wml.mts``'s ``wml``:

    >>> from docx4j_py.wml import wml, to_xml
    >>> p = wml("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
    >>> to_xml(p)
    '<w:p xmlns:w="...">...'

The author writes what is *inside* ``document.xml`` and declares nothing: the
fragment is wrapped in a synthetic root that declares docx4j's whole prefix
table (:mod:`docx4j_py.namespaces`), so ``w14:paraId``, ``r:id``, ``wp:inline``
and ``mc:AlternateContent`` all parse. Declarations inside the fragment win, as
they always do in XML.

How the class is chosen
-----------------------

Unlike the TypeScript runtime, which can only unmarshal a *global* element and
therefore has to wrap a ``w:tbl`` in a ``w:body``, the xsdata parser is told
which class to build and does not check the root element's name, so every
fragment is parsed directly as the class the ``el`` table gives for its
qualified name --- ``w:tbl`` as ``Tbl``, ``w:pPr`` as ``PPr``, ``w:r`` as ``R``.

Where one element name is several classes by scope (``w:sdt``, ``w:customXml``,
``w:t``, ...), the ``el`` table's default is used and ``wrapper=`` overrides it
with the container the fragment is destined for: ``wml(sdt_xml, wrapper='p')``
resolves ``w:sdt`` the way a paragraph would, which is ``SdtRun``. That is the
same knob ``builders/wml.mts`` calls ``wrapper``, decided by the caller rather
than guessed from the first decisive descendant.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from docx4j_py.child import link_parents
from docx4j_py.namespaces import MC_IGNORABLE_PREFIXES, PREFIXES, WML_NS, declarations
from docx4j_py.runtime import context, parser, serializer

__all__ = ["FragmentError", "to_xml", "wml"]

_MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
_WRAPPER = f"{{{WML_NS}}}fragment"


class FragmentError(ValueError):
    """A fragment could not be parsed, or lost content on the way in."""


# ---------------------------------------------------------------------------
# resolving an element name to a class
# ---------------------------------------------------------------------------


def _el_module(uri: str) -> Any:
    from importlib import import_module

    from docx4j_py.el_index import EL_MODULES

    name = EL_MODULES.get(uri)
    if name is None:
        raise FragmentError(f"no generated bindings for the namespace {uri!r}")
    return import_module(name)


def class_for(qname: str) -> type:
    """The class the ``el`` table builds for an element qualified name."""
    if not qname.startswith("{"):
        raise FragmentError(f"{qname!r} is in no namespace; OOXML elements always are")
    uri = qname[1:].split("}", 1)[0]
    table = _el_module(uri).QNAME_TO_CLASS
    try:
        return table[qname]
    except KeyError:
        raise FragmentError(f"no class is bound to the element {qname}") from None


def _scope_class(wrapper: type | str) -> type:
    if isinstance(wrapper, type):
        return wrapper
    name = wrapper if wrapper.startswith("{") else f"{{{WML_NS}}}{wrapper}"
    return class_for(name)


def class_for_in(qname: str, wrapper: type | str) -> type:
    """The class `wrapper` uses for a child element named `qname`.

    The metadata of the containing class answers directly: a single-valued
    element field carries its own name, and the alternatives of a compound
    field carry theirs in the field's ``choices``. This is how ``w:sdt`` inside
    a ``w:p`` comes back as ``SdtRun`` and inside a ``w:body`` as ``SdtBlock``.
    """
    scope = _scope_class(wrapper)
    try:
        meta = context().build(scope)
    except Exception as exc:  # noqa: BLE001
        raise FragmentError(f"{scope.__name__} has no metadata to resolve {qname} in") from exc

    for vars_ in meta.elements.values():
        for var in vars_:
            if var.qname == qname and var.types:
                return var.types[0]
    for choice in meta.choices:
        alt = choice.elements.get(qname)
        if alt is not None and alt.types:
            return alt.types[0]
    raise FragmentError(f"{scope.__name__} holds no {qname}")


# ---------------------------------------------------------------------------
# parsing
# ---------------------------------------------------------------------------

_parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


def _wrap(xml: str) -> etree._Element:
    """Parse the fragment inside a synthetic root carrying every prefix."""
    text = xml.strip()
    if not text:
        raise FragmentError("empty fragment")
    if text.startswith("<?xml"):
        text = text[text.index("?>") + 2 :].lstrip()
    document = f"<w:fragment {declarations()}>{text}</w:fragment>"
    try:
        return etree.fromstring(document.encode("utf-8"), _parser)
    except etree.XMLSyntaxError as exc:
        raise FragmentError(f"not well-formed XML: {exc}") from exc


def _parse_one(node: etree._Element, wrapper: type | str | None, lenient: bool) -> Any:
    qname = node.tag
    if not isinstance(qname, str):
        raise FragmentError("a comment or processing instruction is not a fragment")
    clazz = class_for_in(qname, wrapper) if wrapper is not None else class_for(qname)

    xml_parser = parser(lenient=True, skipped_report=True)
    obj = xml_parser.from_bytes(etree.tostring(node), clazz)
    report = xml_parser.skipped
    if report is not None and not lenient:
        lost = list(getattr(report, "items", ()) or ())
        if lost:
            raise FragmentError(
                f"{len(lost)} item(s) of the fragment were skipped: "
                + "; ".join(str(item) for item in lost[:5])
                + " (pass lenient=True to accept the loss)"
            )
    link_parents(obj, context=context())
    return obj


def _all(xml: str, *, wrapper: type | str | None = None, lenient: bool = False) -> list[Any]:
    root = _wrap(xml)
    children = [child for child in root if isinstance(child.tag, str)]
    if not children:
        raise FragmentError("the fragment holds no element")
    return [_parse_one(child, wrapper, lenient) for child in children]


def wml(xml: str, *, wrapper: type | str | None = None, lenient: bool = False) -> Any:
    """Parse one WordprocessingML fragment and return the typed object.

    Args:
        xml: one element, as it is written inside ``document.xml``; namespace
            declarations are optional.
        wrapper: the element name (``'p'``, ``'body'``, ``'tc'``) or the class
            of the container the fragment is destined for, when the element
            name alone does not settle the class. The ``el`` table's default is
            used without it.
        lenient: accept content the bindings do not know. Off by default: a
            fragment that loses anything raises :class:`FragmentError`, which is
            CR-001's "lenient parsing that silently discards content is not
            acceptable" at the fragment level.

    Returns:
        The parsed object, with every parent pointer below it linked.

    Raises:
        FragmentError: the XML is not well formed, holds no element, holds more
            than one (use :func:`wml.all`), names an element with no binding, or
            lost content.
    """
    objects = _all(xml, wrapper=wrapper, lenient=lenient)
    if len(objects) != 1:
        raise FragmentError(
            f"the fragment holds {len(objects)} elements; use wml.all() for siblings"
        )
    return objects[0]


wml.all = _all  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# serialising
# ---------------------------------------------------------------------------


def _ignorable_prefixes(root: etree._Element) -> set[str]:
    """Prefixes named by any ``mc:Ignorable`` in the tree; they must survive."""
    found: set[str] = set()
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        value = node.get(_MC_IGNORABLE)
        if value:
            found.update(value.split())
    return found


def to_xml(obj: Any, *, pretty: bool = False, name: str | None = None) -> str:
    """Serialise one object as a fragment, with docx4j's prefixes.

    The inverse of :func:`wml`. Three things happen that the raw serialiser does
    not do:

    * **the root gets its element name.** 72 of the generated classes have no
      element name of their own (CR-001 section 13.3) --- ``RT`` is ``w:t`` only
      because ``R.content`` says so --- and the serialiser, given one as the
      root, writes its *class* name. The ``el`` table supplies the real one.
    * **the declarations are trimmed.** The serialiser is handed the whole
      prefix table so that it never invents ``ns0``; what the fragment does not
      use is then removed, except any prefix an ``mc:Ignorable`` in the fragment
      names, which has to stay declared or the markup is invalid.
    * **booleans are numeric**, ``w:val="1"`` as Word writes them (CR-001
      decided question 5).

    Args:
        obj: the object to serialise.
        pretty: indent the output.
        name: force the root element's qualified name.
    """
    from docx4j_py.traversal import element_name

    xml = serializer().render(obj, {p: u for p, u in PREFIXES.items() if p != "xml"})
    root = etree.fromstring(xml.encode("utf-8"), _parser)

    wanted = name or element_name(obj)
    if wanted is None:
        raise FragmentError(
            f"{type(obj).__name__} stands for more than one element name, so a fragment "
            'of it has no name of its own; pass name="{namespace}local"'
        )
    if root.tag != wanted:
        renamed = etree.Element(wanted, nsmap=root.nsmap)
        for key, value in root.attrib.items():
            renamed.set(key, value)
        renamed.text = root.text
        for child in root:
            renamed.append(child)
        root = renamed

    etree.cleanup_namespaces(root, keep_ns_prefixes=sorted(_ignorable_prefixes(root)))
    return etree.tostring(root, pretty_print=pretty, encoding="unicode")


#: The prefixes :func:`wml` declares on its synthetic root, for the record.
DECLARED_PREFIXES = tuple(p for p in PREFIXES if p != "xml")
assert MC_IGNORABLE_PREFIXES.issubset(set(DECLARED_PREFIXES))
