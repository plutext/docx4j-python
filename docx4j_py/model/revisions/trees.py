"""The tree as the revision code walks it. docx4j ``org.docx4j.model.revisions.Trees``.

CR-010 section 3.1. The live lists that hold run-level and block-level content,
the names of their members, their text, and a walk of everything under a node
--- typed objects **and** the ``AnyElement`` wildcard trees CR-002's lossless
MCE keeps --- for what a removal takes with it.

Where docx4j asks a JAXB class, this asks the generated class: the run-level
revisions are ``RunIns``, ``RunDel`` and the ``MoveFrom`` / ``MoveTo`` classes
(``Ins1`` / ``Del1`` are the ``w:ins`` / ``w:del`` *inside* an ``m:r``, and
``Ins2`` / ``Del2`` an ``m:ctrlPr``'s, which are not); a holder's list is its
``content`` field, through ``sdt_content`` for a content control. No model
class is imported at module level (CR-010 section 3.1, CLAUDE.md): the class
tables are built on first use.
"""

from __future__ import annotations

import dataclasses
import enum
import functools
from collections.abc import Callable, Iterable
from typing import Any

from docx4j_py.child import (
    ChildList,
    is_alternate_content,
    is_any_element,
    iter_children,
    mce_branch,
)
from docx4j_py.model.values import integer
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name

__all__ = [
    "MATH_NS",
    "R_NS",
    "W_NS",
    "Revision",
    "View",
    "attribute",
    "block_items",
    "branch_items",
    "children",
    "fit",
    "has_text",
    "index_of",
    "is_empty",
    "is_holder",
    "is_math",
    "item_text",
    "marker_id",
    "name",
    "paragraph_text",
    "revision_of",
    "run_items",
    "text",
    "walk",
]

W_NS = WML_NS
#: The classes of an ``mc:AlternateContent``'s branches.
_MC_BRANCHES = frozenset({"AlternateContentChoice", "AlternateContentFallback"})
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_XML_NS = "http://www.w3.org/XML/1998/namespace"


class Revision(enum.Enum):
    """The four run-level revision holders."""

    INS = "ins"
    DEL = "del"
    MOVE_FROM = "moveFrom"
    MOVE_TO = "moveTo"


class View(enum.Enum):
    """Which text a revision shows in."""

    ACCEPTED = "accepted"
    """The document with every change accepted: inserted text counts, deleted text does not."""
    ORIGINAL = "original"
    """The document with every change rejected."""


# ---------------------------------------------------------------------------
# the class tables, built on first use
# ---------------------------------------------------------------------------


@functools.cache
def _tables() -> dict[str, Any]:
    from docx4j_py import math as m
    from docx4j_py import wml as w

    revisions = {
        w.RunIns: Revision.INS,
        w.RunDel: Revision.DEL,
        w.MoveFrom1: Revision.MOVE_FROM,
        w.MoveFrom2: Revision.MOVE_FROM,
        w.MoveTo1: Revision.MOVE_TO,
        w.MoveTo2: Revision.MOVE_TO,
    }
    run_lists = {
        w.P,
        w.RunIns,
        w.RunDel,
        w.MoveFrom1,
        w.MoveFrom2,
        w.MoveTo1,
        w.MoveTo2,
        w.PHyperlink,
        w.CtBdoHyperlink,
        w.CtCustomXmlRunHyperlink,
        w.CtDirHyperlink,
        w.CtSdtContentRunHyperlink,
        w.CtSimpleFieldHyperlink,
        w.CtSmartTagRunHyperlink,
        w.CTSmartTagRun,
        w.CTCustomXmlRun,
        w.CTSimpleField,
        w.CtBdo,
        w.CtDir,
        w.CTSdtContentRun,
    }
    block_lists = {
        w.Body1,
        w.Hdr,
        w.Ftr,
        w.CTFtnEdn,
        w.CommentsComment,
        w.Tc,
        w.Tbl,
        w.Tr,
        w.CTTxbxContent,
        w.CTCustomXmlBlock,
        w.CTCustomXmlRow,
        w.CTCustomXmlCell,
        w.SdtContentBlock,
        w.CTSdtContentRow,
        w.CTSdtContentCell,
    }
    return {
        "revisions": revisions,
        "run_lists": frozenset(run_lists),
        "run_sdt": frozenset({w.SdtRun}),
        "block_lists": frozenset(block_lists),
        "block_sdt": frozenset({w.SdtBlock, w.CTSdtRow, w.CTSdtCell}),
        "holders": frozenset(
            {
                w.SdtBlock,
                w.SdtRun,
                w.CTSdtRow,
                w.CTSdtCell,
                w.CTCustomXmlRun,
                w.CTCustomXmlBlock,
                w.CTCustomXmlRow,
                w.CTCustomXmlCell,
            }
        ),
        "math_module": m.CTR.__module__,
        "R": w.R,
        "P": w.P,
        "Tbl": w.Tbl,
        "Tr": w.Tr,
        "Tc": w.Tc,
        "CTR": m.CTR,
        "CTCtrlPr": m.CTCtrlPr,
        "CTOMathArg": m.CTOMathArg,
        "CTMathRunTrackChange": w.CTMathRunTrackChange,
        "CTTxbxContent": w.CTTxbxContent,
        "Body": w.Body1,
        "SectPr": w.SectPr,
        "SdtContentBlock": w.SdtContentBlock,
    }


def cls(name_: str) -> type:
    """A model class by its short name, from the table above."""
    return _tables()[name_]


_kind_cache: dict[type, str | None] = {}


def _kind(klass: type) -> str | None:
    """``"run"``, ``"run_sdt"``, ``"block"``, ``"block_sdt"`` or None, through the MRO."""
    try:
        return _kind_cache[klass]
    except KeyError:
        pass
    tables = _tables()
    found: str | None = None
    for base in klass.__mro__:
        if base in tables["run_lists"]:
            found = "run"
        elif base in tables["run_sdt"]:
            found = "run_sdt"
        elif base in tables["block_lists"]:
            found = "block"
        elif base in tables["block_sdt"]:
            found = "block_sdt"
        else:
            continue
        break
    _kind_cache[klass] = found
    return found


# ---------------------------------------------------------------------------
# names
# ---------------------------------------------------------------------------

_names: dict[type, str] = {}


def name(item: Any) -> str:
    """A list member's local element name, ``""`` when the class stands for none.

    A wildcard node's comes from its ``qname``; a typed one's from the
    generated ``el`` table (:func:`~docx4j_py.traversal.element_name`), cached
    by class.
    """
    klass = item.__class__
    found = _names.get(klass)
    if found is not None:
        return found
    if item is None:
        return ""
    if is_any_element(item):
        qname = item.qname
        return qname.rpartition("}")[2] if isinstance(qname, str) else ""
    qname = element_name(klass) or ""
    found = _names[klass] = qname.rpartition("}")[2]
    return found


def attribute(item: Any, local: str, namespace: str = W_NS) -> Any:
    """An attribute of a typed or a wildcard node: ``w:id``, ``w:name``, ``w:author``."""
    if is_any_element(item):
        attributes = item.attributes or {}
        found = attributes.get(f"{{{namespace}}}{local}")
        return found if found is not None else attributes.get(local)
    return getattr(item, local, None)


def marker_id(item: Any) -> int | None:
    """The ``w:id`` of a marker, typed or wildcard, as Word reads it (CR-021): ``"3.0"`` is 3; a refused form None."""
    return integer(attribute(item, "id"))


def is_math(value: Any) -> bool:
    """True for an Office Math object (``m:oMath``, ``m:r``, ``m:sSup``, ...)."""
    return value is not None and value.__class__.__module__ == _tables()["math_module"]


def revision_of(item: Any) -> Revision | None:
    """The run-level revision a list member is --- ``w:ins``, ``w:del``, ``w:moveFrom``, ``w:moveTo`` --- or None."""
    return _tables()["revisions"].get(item.__class__)


# ---------------------------------------------------------------------------
# live content lists
# ---------------------------------------------------------------------------


def branch_items(ac: Any) -> list | None:
    """The live list of the branch of an ``mc:AlternateContent`` a reader of text takes, or None.

    :func:`docx4j_py.child.mce_branch`'s branch (WD REQ-302 to REQ-305): its
    ``any_element`` list, or a wildcard branch's ``children``. docx4j
    ``Trees.blockItems`` and ``runItems`` answer an element so since its
    CR-051 phase 1 (CR-012 Phase A, section 4.1 item 6).
    """
    branch = mce_branch(ac)
    if branch is None:
        return None
    if is_any_element(branch):
        return branch.children
    return branch.any_element


def run_items(value: Any) -> list | None:
    """The live run-level list a run holder keeps, or None.

    A paragraph, a hyperlink (in any of its scopes), an inline content control
    (its ``w:sdtContent``'s list), a smart tag, run-level custom XML, a simple
    field, ``w:dir`` / ``w:bdo``, and the four run-level revisions. A
    paragraph-level ``mc:AlternateContent`` answers with the branch Word reads
    (:func:`branch_items`).
    """
    if value is None:
        return None
    if is_alternate_content(value):
        return branch_items(value)
    kind = _kind(value.__class__)
    if kind == "run":
        return value.content
    if kind == "run_sdt":
        inner = value.sdt_content
        return inner.content if inner is not None else None
    return None


def block_items(value: Any) -> list | None:
    """The live block-level list (or rows, or cells) a container keeps, or None.

    A body, a header, a footer, a note, a comment, a cell, a text box, a block,
    row or cell content control (its ``w:sdtContent``'s list) or custom XML
    element, a table (its rows) and a row (its cells). A block-level
    ``mc:AlternateContent`` answers with the branch Word reads
    (:func:`branch_items`); the row and cell finders pass a row- or cell-level
    one by (:mod:`.tables`).
    """
    if value is None:
        return None
    if is_alternate_content(value):
        return branch_items(value)
    kind = _kind(value.__class__)
    if kind == "block":
        return value.content
    if kind == "block_sdt":
        inner = value.sdt_content
        return inner.content if inner is not None else None
    return None


def children(value: Any) -> list | None:
    """Either kind of list: what a container or a run holder keeps."""
    found = block_items(value)
    return found if found is not None else run_items(value)


def is_holder(value: Any) -> bool:
    """A content control or custom XML element: what ``w:customXmlInsRangeStart`` marks inserted."""
    return value.__class__ in _tables()["holders"]


def list_of(parent: Any) -> list | None:
    """The live list a parent object keeps its children in, or None.

    An ``mc:Choice`` or ``mc:Fallback`` keeps them in ``any_element`` (docx4j
    ``McSplice.contentOf``), so a change straight in a kept branch resolves in
    that branch's list.
    """
    if parent is None:
        return None
    if parent.__class__.__module__ == "docx4j_py.mce" and parent.__class__.__name__ in _MC_BRANCHES:
        return parent.any_element
    found = children(parent)
    if found is not None:
        return found
    content = getattr(parent, "content", None)
    return content if isinstance(content, list) else None


def index_of(items: list | None, item: Any) -> int:
    """Where `item` stands in `items`, by identity; -1 when it does not."""
    if items is None:
        return -1
    for index, candidate in enumerate(items):
        if candidate is item:
            return index
    return -1


# ---------------------------------------------------------------------------
# moving members between lists
# ---------------------------------------------------------------------------

_accepted_cache: dict[type, tuple[frozenset[type], dict[str, type]] | None] = {}


def _accepted(owner_class: type) -> tuple[frozenset[type], dict[str, type]] | None:
    """The classes a ``content`` list of `owner_class` holds, and the class per name."""
    try:
        return _accepted_cache[owner_class]
    except KeyError:
        pass
    from docx4j_py.runtime import context

    found: tuple[frozenset[type], dict[str, type]] | None = None
    try:
        meta = context().build(owner_class)
    except Exception:  # noqa: BLE001 - not a model class: nothing to say
        meta = None
    if meta is not None:
        for var in meta.get_element_vars():
            if var.name != "content" or not var.elements:
                continue
            types: set[type] = set()
            by_name: dict[str, type] = {}
            for qname, alternative in var.elements.items():
                for tp in alternative.types:
                    types.add(tp)
                    by_name.setdefault(qname, tp)
            found = (frozenset(types), by_name)
            break
    _accepted_cache[owner_class] = found
    return found


def fit(item: Any, items: list) -> Any:
    """`item`, as the class the list `items` keeps it as.

    The generated model has one class per element per scope for some names
    (``w:moveFromRangeStart`` is ``MoveFromRangeStart1`` between paragraphs and
    ``MoveFromRangeStart2`` inside one; a hyperlink in a content control is
    ``CtSdtContentRunHyperlink``), so content moved from one list to another ---
    a join, an unwrapped revision or control --- is re-typed where the target
    holds it as another class. Docx4j's JAXB elements have no such scoping.
    """
    owner = items.owner if isinstance(items, ChildList) else None
    if owner is None or is_any_element(item) or not dataclasses.is_dataclass(item):
        return item
    accepted = _accepted(owner.__class__)
    if accepted is None:
        return item
    types, by_name = accepted
    klass = item.__class__
    if klass in types:
        return item
    qname = element_name(item)
    target = by_name.get(qname) if qname is not None else None
    if target is None or target is klass:
        return item
    from docx4j_py.child import deep_copy_as

    return deep_copy_as(item, target)


def splice(items: list, index: int, count: int, moved: Iterable[Any]) -> list:
    """Replace ``items[index:index + count]`` with `moved`, each fitted to `items`.

    `moved` must already be out of the list it came from: a ``ChildList``
    orphans what it loses, so detaching after attaching would clear the parent
    this sets.
    """
    fitted = [fit(item, items) for item in moved]
    items[index : index + count] = fitted
    return fitted


# ---------------------------------------------------------------------------
# text, Office JS's way
# ---------------------------------------------------------------------------

_CONSTANT_TEXT = {
    "tab": "\t",
    "br": "\n",
    "cr": "\n",
    "noBreakHyphen": "‑",
    "softHyphen": "­",
}


def item_text(item: Any) -> str | None:
    """The text of a run's item, Office JS's way; None for an item that is no text.

    ``w:t`` and ``w:delText`` (and an equation's ``m:t``) give their value, a
    tab ``\\t``, a break and a carriage return ``\\n``, the two hyphens their
    characters and ``w:sym`` the character its ``w:char`` names.
    """
    local = name(item)
    if local in ("t", "delText"):
        value = getattr(item, "value", None)
        return value if isinstance(value, str) else ""
    constant = _CONSTANT_TEXT.get(local)
    if constant is not None:
        return constant
    if local == "sym":
        code = getattr(item, "char", None) or getattr(item, "char_value", None)
        if isinstance(code, str) and len(code) == 4:
            try:
                return chr(int(code, 16))
            except ValueError:
                return ""
        return ""
    return None


def has_text(run: Any) -> bool:
    """True when a run holds text: a comment reference alone does not (check 32)."""
    content = getattr(run, "content", None)
    if not content:
        return False
    return any(item_text(item) is not None for item in content)


def text(item: Any, view: View) -> str:
    """The text of run-level content in a view: runs, math, and what run holders hold."""
    out: list[str] = []
    append_text(item, view, out)
    return "".join(out)


def append_text(item: Any, view: View, out: list[str]) -> None:
    """docx4j ``Trees.appendText``."""
    tables = _tables()
    klass = item.__class__
    if klass is tables["R"]:
        for child in item.content:
            found = item_text(child)
            if found is not None:
                out.append(found)
        return
    if klass is tables["CTR"]:
        for child in item.content:
            if isinstance(child, tables["CTMathRunTrackChange"]):
                inserted = name(child) == "ins"
                if inserted == (view is View.ACCEPTED):
                    for inner in child.content:
                        found = item_text(inner)
                        if found is not None:
                            out.append(found)
                continue
            found = item_text(child)
            if found is not None:
                out.append(found)
        return
    if is_math(item):

        def enter(node: Any) -> bool:
            if node is not item and node.__class__ is tables["CTR"]:
                append_text(node, view, out)
                return False
            return node.__class__ is not tables["CTCtrlPr"]

        walk(item, enter, wildcards=False)
        return
    revision = revision_of(item)
    if revision is not None:
        if view is View.ACCEPTED:
            shown = revision in (Revision.INS, Revision.MOVE_TO)
        else:
            shown = revision in (Revision.DEL, Revision.MOVE_FROM)
        if not shown:
            return
    items = run_items(item)
    if items is not None:
        for child in items:
            append_text(child, view, out)


def paragraph_text(paragraph: Any, view: View) -> str:
    """A paragraph's text in a view."""
    return text(paragraph, view)


# ---------------------------------------------------------------------------
# the walk
# ---------------------------------------------------------------------------

_rel_fields_cache: dict[type, tuple[str, ...]] = {}


def _rel_fields(klass: type) -> tuple[str, ...]:
    """The fields of a class that are ``r:`` attributes: relationship ids."""
    try:
        return _rel_fields_cache[klass]
    except KeyError:
        pass
    found: tuple[str, ...] = ()
    try:
        found = tuple(
            field.name
            for field in dataclasses.fields(klass)
            if field.metadata.get("type") == "Attribute" and field.metadata.get("namespace") == R_NS
        )
    except TypeError:
        found = ()
    _rel_fields_cache[klass] = found
    return found


def walk(
    root: Any,
    enter: Callable[[Any], bool | None] | None = None,
    relationship: Callable[[str], None] | None = None,
    *,
    wildcards: bool = True,
) -> None:
    """Every node under `root` (itself included), pre-order. docx4j ``Trees.walk``.

    `enter` is called with each node --- a typed object, or with `wildcards` a
    wildcard ``AnyElement`` --- and returning ``False`` skips what is under it.
    `relationship` is called with every ``r:`` attribute's value met on the
    way, typed (``r:id``, ``r:embed``, ...) or on a wildcard node, which is
    where a picture's ``r:embed`` is in this model. Both branches of
    ``mc:AlternateContent`` are walked.
    """
    stack: list[Any] = [root]
    seen: set[int] = set()
    rel_prefix = f"{{{R_NS}}}"
    while stack:
        node = stack.pop()
        key = id(node)
        if key in seen:
            continue
        seen.add(key)
        if is_any_element(node):
            if not wildcards:
                kids = [c for c in node.children or () if not isinstance(c, str)]
                kids.reverse()
                stack.extend(kids)
                continue
            if enter is not None and enter(node) is False:
                continue
            if relationship is not None:
                for attr_name, value in (node.attributes or {}).items():
                    if isinstance(attr_name, str) and attr_name.startswith(rel_prefix) and value:
                        relationship(value)
            kids = [c for c in node.children or () if not isinstance(c, str)]
            kids.reverse()
            stack.extend(kids)
            continue
        if isinstance(node, str) or not dataclasses.is_dataclass(node):
            continue
        if enter is not None and enter(node) is False:
            continue
        if relationship is not None:
            for field_name in _rel_fields(node.__class__):
                value = getattr(node, field_name, None)
                if isinstance(value, str) and value:
                    relationship(value)
        kids = [child for _qname, child in iter_children(node, mce="all", wildcards=wildcards)]
        kids.reverse()
        stack.extend(kids)


def is_empty(value: Any) -> bool:
    """True when a model object holds nothing: every field None, or an empty list."""
    if value is None:
        return True
    for field in dataclasses.fields(value):
        found = getattr(value, field.name, None)
        if found is None:
            continue
        if isinstance(found, list) and not found:
            continue
        return False
    return True


def xml_space_of(item: Any) -> Any:
    """``xml:space`` of a text item, typed or wildcard."""
    if is_any_element(item):
        return (item.attributes or {}).get(f"{{{_XML_NS}}}space")
    return getattr(item, "space", None)
