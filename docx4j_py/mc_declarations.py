"""Where a save declares the prefixes ``mc`` attributes name. docx4j CR-052's ``McDeclarations``.

CR-011 Q3 and Q11. ``mc:Choice/@Requires``, a Choice's own ``mc:Ignorable``
and ``mc:MustUnderstand``, and a part root's ``mc:Ignorable`` name namespace
**prefixes**, so each prefix must be declared where it is named. The
serialiser writes the prefix table's declarations and the cleanup keeps only
those an element or attribute *name* uses, so a prefix nothing but such a
value names lost its declaration on a re-marshalled save (the probe's ``zz9``
and ``shape2010``, Word's own ``cx1`` on a chartex chart), and Word 365 repairs
such a file (docx4j CR-052 phase 0).

The rule is docx4j CR-052's section 4 (its D3 and D4 (c), as Word 365 places
them over 2,728 of its saves):

* **the binding**: the one the Choice carries from its source
  (:class:`~docx4j_py.child.BoundPrefixes`, recorded at parse), else the
  prefix table's (:data:`~docx4j_py.namespaces.PREFIXES`); a prefix that
  neither binds is left undeclared, with a warning;
* **the place**: a prefix the table binds as the source did goes on the
  root, as Word puts its fixed set; any other --- a private prefix (``zz9``),
  an alias (``shape2010`` for the wps namespace), a table prefix the source
  bound otherwise --- goes on the ``mc:Choice`` that names it, ahead of
  ``Requires`` (``<mc:Choice xmlns:zz9="..." Requires="zz9">``); a prefix the
  root's ``mc:Ignorable`` names with the same binding is declared once, on the
  root; nothing goes on ``mc:Fallback``;
* **the clash**: where the Choice's content uses the namespace the prefix has
  in scope (from the table or the root), a declaration on the Choice would
  rebind that content, so the content's binding stands, with a warning.

The records are read from the object tree, paired with the rendered
``mc:Choice`` elements in document order and checked by their values, only
when some binding recorded in this process is not the table's
(:func:`~docx4j_py.child.other_bindings_seen`): until then the table alone
gives every answer the records would. Every marshal path uses it:
``XmlPart._marshal`` and :func:`docx4j_py.fragments.to_xml`.

Not covered, as in docx4j's phase 1: a source-only binding the root keeps
because its own content uses it, and the ``mc:AlternateContent`` element's
own ``mc:Ignorable``.
"""

from __future__ import annotations

import logging
from typing import Any

from lxml import etree

from docx4j_py.namespaces import PREFIXES

__all__ = ["Plan", "finish", "plan"]

logger = logging.getLogger("docx4j_py.openpackaging")

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_CHOICE = f"{{{_MC_NS}}}Choice"
#: The rendered Choice's attributes that name prefixes, in the record's order.
_NAMING = ("Requires", f"{{{_MC_NS}}}Ignorable", f"{{{_MC_NS}}}MustUnderstand")

#: A namespace no document uses, for a declaration lxml would otherwise strip.
_PLACEHOLDER = "urn:docx4j-python:declare:"


class Plan:
    """What a marshal declares beyond the cleanup: on the root, and on each ``mc:Choice``."""

    __slots__ = ("choices", "placeholders", "root")

    def __init__(self) -> None:
        #: Prefix -> URI, to declare on the root.
        self.root: dict[str, str] = {}
        #: Each Choice element, with what to declare on it.
        self.choices: list[tuple[etree._Element, dict[str, str]]] = []
        #: The placeholder declarations written, with what each stands for.
        self.placeholders: dict[bytes, bytes] = {}

    def apply(self) -> None:
        """Rebuild each Choice with its declarations; :func:`finish` puts the URIs in."""
        for element, declared in self.choices:
            _declare_on(element, declared, self.placeholders)


def plan(root: etree._Element, obj: Any, on_root: dict[str, str], where: Any = None) -> Plan:
    """Where each prefix the rendered `root`'s Choices name, and the tree leaves unbound, is to be declared.

    Args:
        root: the rendered tree, its declarations cleaned up.
        obj: the object it was rendered from, whose Choices carry the records.
        on_root: what the root is about to declare besides its own (the part
            root's ``mc:Ignorable`` prefixes): in scope everywhere.
        where: the part, for the warnings.
    """
    out = Plan()
    elements = list(root.iter(_MC_CHOICE))
    if not elements:
        return out
    records: list[dict[str, str | None]] | None = None
    if _other_bindings_seen():
        records = _records(obj, elements)
    for index, element in enumerate(elements):
        named = [prefix for key in _NAMING for prefix in (element.get(key) or "").split()]
        if not named:
            continue
        recorded = records[index] if records is not None else {}
        scope = element.nsmap
        here: dict[str, str] = {}
        for prefix in dict.fromkeys(named):
            wanted = recorded.get(prefix) or PREFIXES.get(prefix)
            in_scope = scope.get(prefix)
            if in_scope is None:
                in_scope = out.root.get(prefix) or on_root.get(prefix)
            if wanted is None:
                if in_scope is None:
                    logger.warning(
                        "%s: an mc:Choice names the prefix %r, which neither its source nor the "
                        "prefix table binds; it is left undeclared",
                        where,
                        prefix,
                    )
                continue
            if in_scope == wanted:
                continue
            if in_scope is None:
                if PREFIXES.get(prefix) == wanted:
                    out.root[prefix] = wanted
                else:
                    here[prefix] = wanted
                continue
            if _uses(element, in_scope):
                logger.warning(
                    "%s: an mc:Choice names the prefix %r, bound to %s in its source, but its content "
                    "uses %s, which the prefix has in scope; the content's binding stands",
                    where,
                    prefix,
                    wanted,
                    in_scope,
                )
                continue
            here[prefix] = wanted
        if here:
            out.choices.append((element, here))
    return out


def finish(data: bytes, applied: Plan) -> bytes:
    """The serialised bytes with each placeholder declaration given its URI."""
    for placeholder, declaration in applied.placeholders.items():
        data = data.replace(placeholder, declaration, 1)
    return data


def _other_bindings_seen() -> bool:
    from docx4j_py.child import other_bindings_seen

    return other_bindings_seen()


def _records(obj: Any, elements: list[etree._Element]) -> list[dict[str, str | None]] | None:
    """Each ``mc:Choice``'s source bindings, in document order, paired with `elements`; None where they do not pair."""
    from docx4j_py.child import MC_CHOICE, is_any_element, iter_children, requires_of, source_bindings

    found: list[Any] = []
    stack: list[Any] = [obj]
    while stack:
        node = stack.pop()
        if node is None or isinstance(node, str):
            continue
        if is_any_element(node):
            if node.qname == MC_CHOICE:
                found.append(node)
            kids = list(node.children or ())
        else:
            if type(node).__name__ == "AlternateContentChoice":
                found.append(node)
            kids = [child for _qname, child in iter_children(node, mce="all", wildcards=True)]
        kids.reverse()
        stack.extend(kids)
    if len(found) != len(elements):
        logger.warning("the mc:Choice records do not pair with the rendered tree; the prefix table is used")
        return None
    records: list[dict[str, str | None]] = []
    for choice, element in zip(found, elements, strict=True):
        if (requires_of(choice) or "") != (element.get("Requires") or ""):
            logger.warning("the mc:Choice records do not pair with the rendered tree; the prefix table is used")
            return None
        records.append(source_bindings(choice))
    return records


def _uses(element: etree._Element, uri: str) -> bool:
    """Whether anything under `element` is in the namespace `uri`: an element's name or an attribute's."""
    prefix = f"{{{uri}}}"
    for node in element.iterdescendants():
        tag = node.tag
        if isinstance(tag, str) and tag.startswith(prefix):
            return True
        if any(key.startswith(prefix) for key in node.attrib):
            return True
    return False


def _escape(value: str) -> str:
    """An attribute value's characters as XML writes them in double quotes."""
    return value.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")


def _declare_on(element: etree._Element, declared: dict[str, str], placeholders: dict[bytes, bytes]) -> None:
    """The element made again with `declared` among its declarations, in its place.

    lxml fixes an element's declarations when it is made, so the ``mc:Choice``
    is rebuilt: its name, its attributes, text, children and tail, and these
    declarations, which lxml writes ahead of ``Requires`` as Word does. lxml
    strips a declaration whose namespace is already in scope under another
    prefix (``shape2010`` beside the root's ``wps``) and one it finds unused,
    so each is made with a placeholder namespace of its own and its URI put
    back in the serialised bytes (:func:`finish`).
    """
    parent = element.getparent()
    if parent is None:
        return
    nsmap: dict[str, str] = {}
    for prefix, uri in declared.items():
        stand_in = f"{_PLACEHOLDER}{len(placeholders)}"
        nsmap[prefix] = stand_in
        placeholders[f'xmlns:{prefix}="{stand_in}"'.encode()] = f'xmlns:{prefix}="{_escape(uri)}"'.encode()
    fresh = parent.makeelement(element.tag, attrib=dict(element.attrib), nsmap=nsmap)
    fresh.text = element.text
    fresh.tail = element.tail
    for child in list(element):
        fresh.append(child)
    parent.replace(element, fresh)
