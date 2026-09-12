"""Walking a tree and reading its text: ``TraversalUtil``, ``ClassFinder``, ``TextUtils``.

CR-001 section 6.2. Three docx4j utilities, none of them with any per-class
code: the enumeration of a node's children is
:func:`docx4j_py.child.iter_children`, which reads the ``XmlContext`` metadata,
and it is the same enumeration :func:`docx4j_py.child.link_parents` uses.

``walk(root, visitor)``
    docx4j ``TraversalUtil``: pre-order over every typed node. The visitor is
    called with ``(node, parent, element name)``; returning ``False`` stops the
    walk descending into that node, everything else continues.
``iter_nodes(root)``
    the same traversal as a generator, for when a callback is in the way.
``find(root, P)`` / ``find(root, (P, Tbl))``
    docx4j ``ClassFinder``, with ``isinstance`` rather than ``getClass() ==``,
    so the element-specific subclasses of CR-001 decided question 2 match their
    base: ``find(root, CTTrackChange)`` finds every ``RunIns`` and ``RunDel``.
``text_of(obj)``
    docx4j ``TextUtils``.

``element_name(obj)`` is the piece that makes ``text_of`` possible at all: 72 of
the generated classes (CR-001 section 13.3) have no element name of their own,
so the name a node is written under comes from the *parent's* metadata. Where
there is no parent to ask --- the root of the walk --- the generated ``el``
module's ``CLASS_TO_QNAME`` answers instead.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from importlib import import_module
from typing import Any

from docx4j_py.child import MCE_NS, iter_children
from docx4j_py.namespaces import WML_NS

__all__ = [
    "element_name",
    "find",
    "iter_nodes",
    "text_of",
    "walk",
]


# ---------------------------------------------------------------------------
# element names
# ---------------------------------------------------------------------------

_class_qnames: dict[str, dict[type, str]] = {}


def _module_qnames(module: str) -> dict[type, str]:
    """``CLASS_TO_QNAME`` of the ``el`` module belonging to a model module."""
    try:
        return _class_qnames[module]
    except KeyError:
        pass
    try:
        table = import_module(f"{module}.el").CLASS_TO_QNAME
    except ModuleNotFoundError:
        # a permanent answer: ``builtins``, ``docx4j_py.child`` and the other
        # modules in a class's MRO have no ``el``
        _class_qnames[module] = {}
        return {}
    except Exception:  # noqa: BLE001 - transient (a partially imported module)
        return {}
    _class_qnames[module] = table
    return table


def element_name(obj_or_class: Any) -> str | None:
    """The qualified name a node is written under, as far as it can be known.

    Two sources, in this order:

    1. the generated ``el`` table (``CLASS_TO_QNAME``), which knows the one
       element name a class stands for where there is only one --- the answer
       for the type classes (``Tbl`` is ``w:tbl``) and for the 72 intermediate
       choice classes, which have no name of their own at all (``RT`` is
       ``w:t`` only because ``R.content`` says so);
    2. the class's own ``Meta``, when it declares a ``namespace`` as well as a
       ``name``: that is exactly what xsdata emits for a **global element**
       declaration (``Document`` is ``{...}document``, ``RunIns`` is
       ``{...}ins``) and never for a type.

    ``XmlMeta.qname`` is deliberately *not* consulted: xsdata builds a type
    class's qname against whatever parent namespace it first met the class in
    and caches it, so the same class reads ``CT_PPr`` before a parse and
    ``{...}CT_PPr`` after one. ``Meta`` is stable.

    Returns None for a class that stands for several element names
    (``TblWidth`` is ``w:left``, ``w:right``, ``w:tblCellSpacing`` and five
    more): there is no answer without a scope.
    """
    cls = obj_or_class if isinstance(obj_or_class, type) else obj_or_class.__class__
    for klass in cls.__mro__:
        name = _module_qnames(klass.__module__).get(klass)
        if name is not None:
            return name
        meta = vars(klass).get("Meta")
        if meta is not None:
            namespace = getattr(meta, "namespace", None)
            local = getattr(meta, "name", None)
            if namespace and local:
                return f"{{{namespace}}}{local}"
    return None


# ---------------------------------------------------------------------------
# traversal
# ---------------------------------------------------------------------------


def walk(
    root: Any,
    visitor: Callable[[Any, Any, str | None], bool | None],
    *,
    context: Any = None,
    mce: str = "resolve",
) -> None:
    """Visit `root` and everything below it, pre-order (docx4j ``TraversalUtil``).

    Args:
        root: the node to start at; it is visited first.
        visitor: called ``visitor(node, parent, element_name)``. Return ``False``
            to stop descending into that node; anything else (``None``
            included) descends.
        context: an ``XmlContext`` to read the metadata from; the shared one by
            default.
        mce: how to treat ``mc:AlternateContent`` --- ``"resolve"`` (the
            default, the view Word takes), ``"all"`` or ``"none"``. See
            :func:`docx4j_py.child.iter_children` and CR-002 section 5.6.

    A node reached twice --- which a hand-built tree can arrange --- is visited
    once.
    """
    if context is None:
        from docx4j_py.runtime import context as shared

        context = shared()

    seen: set[int] = set()
    stack: list[tuple[Any, Any, str | None]] = [(root, None, element_name(root))]
    while stack:
        node, parent, name = stack.pop()
        key = id(node)
        if key in seen:
            continue
        seen.add(key)
        if visitor(node, parent, name) is False:
            continue
        children = list(iter_children(node, context=context, mce=mce))
        children.reverse()
        stack.extend((child, node, qname) for qname, child in children)


def iter_nodes(root: Any, *, context: Any = None, mce: str = "resolve") -> Iterator[Any]:
    """Yield `root` and every node below it, pre-order.

    The generator form of :func:`walk`, for the common case of a filter that
    never wants to prune. `mce` is as :func:`walk`'s.
    """
    if context is None:
        from docx4j_py.runtime import context as shared

        context = shared()

    seen: set[int] = set()
    stack: list[Any] = [root]
    while stack:
        node = stack.pop()
        key = id(node)
        if key in seen:
            continue
        seen.add(key)
        yield node
        children = [child for _q, child in iter_children(node, context=context, mce=mce)]
        children.reverse()
        stack.extend(children)


def find(
    root: Any,
    cls_or_tuple: type | tuple[type, ...],
    *,
    context: Any = None,
    mce: str = "resolve",
) -> list[Any]:
    """Every node below `root`, `root` included, that is an instance of the class.

    docx4j's ``ClassFinder``. ``isinstance`` rather than an exact class test, so
    a base class finds the element-specific subclasses xsdata generates
    (CR-001 decided question 2): ``find(doc, CTTrackChange)`` returns the
    ``RunIns`` and ``RunDel`` nodes, and ``find(doc, Text)`` the ``RT`` ones.

    ``mce`` is as :func:`walk`'s: by default the branch of an
    ``mc:AlternateContent`` a consumer would ignore is not searched, which is
    the answer a caller asking "what is in this document" wants. Pass
    ``mce="all"`` to search both branches.
    """
    return [
        node
        for node in iter_nodes(root, context=context, mce=mce)
        if isinstance(node, cls_or_tuple)
    ]


# ---------------------------------------------------------------------------
# TextUtils
# ---------------------------------------------------------------------------


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


#: A run's children that contribute text, and what they contribute.
#: ``w:t`` and ``w:sym`` are computed; the rest are constants. docx4j's
#: ``TextUtils`` writes ``\t`` for a tab and a newline for a break or a carriage
#: return; ``w:noBreakHyphen`` and ``w:softHyphen`` are the Unicode characters
#: Word means by them.
_RUN_TEXT: dict[str, str] = {
    _w("tab"): "\t",
    _w("br"): "\n",
    _w("cr"): "\n",
    _w("noBreakHyphen"): "‑",
    _w("softHyphen"): "­",
}

#: Deliberately absent from ``_RUN_TEXT``: ``w:delText`` is deleted text and
#: ``w:instrText`` / ``w:delInstrText`` are field instructions, neither of which
#: is document text. ``w:del`` is likewise absent from ``_RUN_CONTAINERS``, so a
#: tracked deletion contributes nothing.
_EXCLUDED: frozenset[str] = frozenset(
    {_w("delText"), _w("instrText"), _w("delInstrText"), _w("del")}
)

_RUN = _w("r")
_PARAGRAPH = _w("p")
_TEXT = _w("t")
_SYM = _w("sym")

#: Elements whose children are runs, or more of these. The list is docx4j's
#: (and ``builders/wml.mts``'s): hyperlinks, content controls, smart tags,
#: custom XML, simple fields, bidirectional overrides and *insertions*.
_RUN_CONTAINERS: frozenset[str] = frozenset(
    {
        _w("hyperlink"),
        _w("sdt"),
        _w("sdtContent"),
        _w("smartTag"),
        _w("customXml"),
        _w("ins"),
        _w("moveTo"),
        _w("dir"),
        _w("bdo"),
        _w("fldSimple"),
    }
)


def _sym_text(node: Any) -> str:
    """``w:sym`` as a character, from its hexadecimal ``w:char``."""
    for attr in ("char_value", "char", "char_attribute"):
        value = getattr(node, attr, None)
        if isinstance(value, str) and len(value) == 4:
            try:
                return chr(int(value, 16))
            except ValueError:
                return ""
    # an AnyElement: the attribute is in the dict, under its qualified name
    attributes = getattr(node, "attributes", None)
    if isinstance(attributes, dict):
        value = attributes.get(_w("char")) or attributes.get("char")
        if isinstance(value, str) and len(value) == 4:
            try:
                return chr(int(value, 16))
            except ValueError:
                return ""
    return ""


def _text_value(node: Any) -> str:
    """The characters of a ``w:t``, typed (``value``) or wildcard (``text``)."""
    value = getattr(node, "value", None)
    if isinstance(value, str):
        return value
    text = getattr(node, "text", None)
    return text if isinstance(text, str) else ""


def _item_text(qname: str | None, node: Any) -> str:
    if qname == _TEXT:
        return _text_value(node)
    if qname == _SYM:
        return _sym_text(node)
    if qname is None:
        return ""
    return _RUN_TEXT.get(qname, "")


#: The markup-compatibility wrappers. Under ``mce="resolve"`` the child
#: enumeration never yields them --- the chosen branch's content stands in
#: their place --- so these only matter under ``mce="all"``, where they are
#: transparent: the text of *every* branch is read, which is exactly the
#: double counting ``"resolve"`` exists to avoid.
_MCE_CONTAINERS: frozenset[str] = frozenset(
    {
        f"{{{MCE_NS}}}AlternateContent",
        f"{{{MCE_NS}}}Choice",
        f"{{{MCE_NS}}}Fallback",
    }
)


def _children(node: Any, context: Any, mce: str):
    return iter_children(node, context=context, mce=mce, wildcards=True)


def _run_text(run: Any, context: Any, mce: str) -> str:
    out: list[str] = []
    for qname, child in _children(run, context, mce):
        if qname in _MCE_CONTAINERS:
            out.append(_run_text(child, context, mce))
        else:
            out.append(_item_text(qname, child))
    return "".join(out)


def _inline_text(container: Any, context: Any, mce: str) -> str:
    out: list[str] = []
    for qname, child in _children(container, context, mce):
        if qname in _EXCLUDED:
            continue
        if qname == _RUN:
            out.append(_run_text(child, context, mce))
        elif qname in _RUN_CONTAINERS or qname == _PARAGRAPH or qname in _MCE_CONTAINERS:
            out.append(_inline_text(child, context, mce))
    return "".join(out)


def _block_texts(node: Any, qname: str | None, context: Any, out: list[str], mce: str) -> None:
    if qname in _EXCLUDED:
        return
    if qname == _RUN:
        out.append(_run_text(node, context, mce))
        return
    if qname == _PARAGRAPH or qname in _RUN_CONTAINERS:
        out.append(_inline_text(node, context, mce))
        return
    for child_qname, child in _children(node, context, mce):
        _block_texts(child, child_qname, context, out, mce)


def text_of(obj: Any, *, context: Any = None, mce: str = "resolve") -> str:
    """The text of a run, a paragraph or anything that holds them.

    docx4j's ``TextUtils``, and the same rules as ``builders/wml.mts``'s
    ``textOf``:

    * ``w:t`` contributes its value, ``w:tab`` a tab, ``w:br`` and ``w:cr`` a
      newline, ``w:noBreakHyphen`` U+2011, ``w:softHyphen`` U+00AD and ``w:sym``
      the character its ``w:char`` names;
    * runs inside hyperlinks, content controls, smart tags, custom XML, simple
      fields and *insertions* are read;
    * ``w:del`` (a tracked deletion) and its ``w:delText`` are skipped, and so
      are the field instructions ``w:instrText`` and ``w:delInstrText``;
    * every paragraph, and every run found outside one, is a line: the pieces
      are joined with a newline.

    ``mc:AlternateContent`` is **resolved** by default (CR-002 section 5.6):
    only the branch a consumer that understands
    :data:`docx4j_py.namespaces.UNDERSTOOD` would take contributes, so a
    document whose text box is written once as a ``wps`` shape and once as a
    VML fallback is not counted twice, which is the view Word's own text
    extraction takes. ``mce="all"`` counts every branch; ``mce="none"`` counts
    none of them.

    Unlike ``walk`` and ``find``, this reads wildcard (``AnyElement``) content
    as well as model objects: the content of an ``mc:Choice`` is a wildcard
    tree in this model, and the text in it is real text.
    """
    if context is None:
        from docx4j_py.runtime import context as shared

        context = shared()

    out: list[str] = []
    _block_texts(obj, element_name(obj), context, out, mce)
    return "\n".join(out)
