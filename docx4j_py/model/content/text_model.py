"""A paragraph's text as segments, and the block structure under a container.

This is the piece every view rests on. A paragraph's text is not a string in the
tree: it is spread over runs, and over the hyperlinks, content controls, smart
tags and tracked insertions those runs sit in. :func:`segments_of` flattens it
into one list of :class:`Segment`, each mapping a slice of the accepted-view
text to the item that carries it and to the lists that hold that item and its
run --- which is everything an edit needs to change the text in place, split a
run at an offset, or apply formatting to exactly one span.

What counts as text (docx4j ``TextUtils``, and the same rules as
:func:`docx4j_py.traversal.text_of`): ``w:t``, ``w:tab`` (a tab), ``w:br`` and
``w:cr`` (a newline), ``w:noBreakHyphen`` (U+2011), ``w:softHyphen`` (U+00AD)
and ``w:sym`` (the character its ``w:char`` names). ``w:delText``,
``w:instrText`` and ``w:delInstrText`` are not text --- except that a
``w:delText`` *is* the text of the original view, inside the ``w:del`` or
``w:moveFrom`` that removed it, which is the only place it is read.

What is descended into: ``w:hyperlink``, ``w:smartTag``, ``w:customXml``, a
run-level ``w:sdt``, ``w:fldSimple``, ``w:dir``, ``w:bdo``, and the revision
holder the view shows --- ``w:ins`` and ``w:moveTo`` in the accepted view,
``w:del`` and ``w:moveFrom`` in the original one. What is **not**: a
``w:drawing``, and therefore a text box. ``Body.text`` is Word's main story
(CR-003 sections 3.13 and decided question 8).

**Word's branch** (CR-011 section 15, Phase F; Q4 revised 2026-10-22). A run- or
paragraph-level ``mc:AlternateContent``, which the tree keeps whole, is read as
the branch Word shows, :func:`docx4j_py.child.mce_branch` (docx4j's TEXT rule
over its published understood list, as :func:`~docx4j_py.traversal.text_of`
and the text model read it): a run-level element's branch as the run's own
items, a paragraph-level one's runs as the paragraph's own runs, with the same
characters and the same revision filter. Office JS reads the same (core-ts
check 49). Such a segment carries the element as :attr:`Segment.branch` and is
**not editable**: an edit whose span holds its text, or an insertion strictly
inside its text, is refused (``address.read_only``, :func:`refuse_branch_text`),
since it would change one branch alone. An insertion at either end of the
element's text goes beside the element (:meth:`Segment.site`).

Offsets are **code points** (CR-003 decided question 3), and a split never falls
inside a grapheme cluster (section 3.12): :func:`is_grapheme_boundary` is
``unicodedata``'s combining classes plus the zero-width joiner, the variation
selectors, the emoji modifiers, regional-indicator pairs and the Indic
virama conjuncts, and the ``regex`` module's ``\\X`` is used for
:func:`grapheme_clusters` when it happens to be installed.
"""

from __future__ import annotations

import dataclasses
import re
import sys
import unicodedata
from typing import Any

from docx4j_py.child import (
    is_alternate_content,
    is_any_element,
    iter_children,
    mce_branch,
)
from docx4j_py.model.content.errors import ContentError
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of

__all__ = [
    "RUN_HOLDER_NAMES",
    "TEXT_ITEM_NAMES",
    "VIEWS",
    "Reading",
    "Segment",
    "block_children_of",
    "block_list_of",
    "branch_extents",
    "branch_text",
    "canonical_view",
    "cells_of",
    "find_all",
    "grapheme_clusters",
    "is_grapheme_boundary",
    "item_text",
    "kept_content_under",
    "markup_text",
    "reading_of",
    "refuse_branch_text",
    "require_view",
    "rows_of",
    "runs_of",
    "search_pattern",
    "segments_of",
    "set_text",
    "snap_back",
    "snap_forward",
    "split_at",
    "text_of_view",
    "touches_branch_text",
]


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


W_P = _w("p")
W_R = _w("r")
W_T = _w("t")
W_SYM = _w("sym")
W_SDT = _w("sdt")
W_DEL_TEXT = _w("delText")
W_TBL = _w("tbl")
W_TR = _w("tr")
W_TC = _w("tc")
W_CUSTOM_XML = _w("customXml")

#: What may wrap a row or a cell without being one: an OpenDoPE repeat puts its
#: ``w:tr`` inside a ``w:sdt``, and a legacy binding inside a ``w:customXml``
#: (CR-003 section 4: rows and cells descend into row- and cell-level controls).
_ROW_WRAPPERS = frozenset({W_SDT, W_CUSTOM_XML})

#: A run's children that contribute a constant string.
_CONSTANT_TEXT: dict[str, str] = {
    _w("tab"): "\t",
    _w("br"): "\n",
    _w("cr"): "\n",
    _w("noBreakHyphen"): "‑",
    _w("softHyphen"): "­",
}

#: Every run child that carries text. ``w:delText``, ``w:instrText`` and
#: ``w:delInstrText`` are deliberately absent: deleted text and field
#: instructions are not document text.
TEXT_ITEM_NAMES: frozenset[str] = frozenset({W_T, W_SYM, *_CONSTANT_TEXT})

#: Run-level containers whose items are read as part of the paragraph's text.
RUN_HOLDER_NAMES: frozenset[str] = frozenset(
    {
        _w("hyperlink"),
        _w("smartTag"),
        _w("customXml"),
        _w("fldSimple"),
        _w("dir"),
        _w("bdo"),
        W_SDT,
        _w("sdtContent"),
    }
)

#: The four run-level revision holders, and the view each one shows in.
_REVISIONS: dict[str, str] = {
    _w("ins"): "accepted",
    _w("moveTo"): "accepted",
    _w("del"): "original",
    _w("moveFrom"): "original",
}


def _sym_text(node: Any) -> str:
    """``w:sym`` as a character, from its hexadecimal ``w:char``."""
    for attribute in ("char_value", "char", "char_attribute"):
        value = getattr(node, attribute, None)
        if isinstance(value, str) and len(value) == 4:
            try:
                return chr(int(value, 16))
            except ValueError:
                return ""
    return ""


def _qname(node: Any) -> str | None:
    """A node's element name: a wildcard's own ``qname``, else :func:`element_name`'s."""
    if is_any_element(node):
        return node.qname
    return element_name(node)


def _wildcard_text(qname: str | None, node: Any) -> str | None:
    """The text a wildcard run child contributes (a kept branch's ``AnyElement``), as :func:`item_text`."""
    if qname in (W_T, W_DEL_TEXT):
        value = getattr(node, "text", None)
        return value if isinstance(value, str) else ""
    if qname == W_SYM:
        attributes = getattr(node, "attributes", None) or {}
        value = attributes.get(_w("char")) or attributes.get("char")
        if isinstance(value, str) and len(value) == 4:
            try:
                return chr(int(value, 16))
            except ValueError:
                return ""
        return ""
    if qname is None:
        return None
    return _CONSTANT_TEXT.get(qname)


def _branch_items(branch: Any) -> list | None:
    """The live list a branch holds: a typed one's ``any_element``, a wildcard one's children."""
    if branch is None:
        return None
    if is_any_element(branch):
        children = getattr(branch, "children", None)
        return children if isinstance(children, list) else None
    items = getattr(branch, "any_element", None)
    return items if isinstance(items, list) else None


def _wildcard_items(node: Any) -> list | None:
    """A wildcard holder's or run's children, the list a kept branch's untyped content keeps."""
    if is_any_element(node):
        children = getattr(node, "children", None)
        return children if isinstance(children, list) else None
    return None


def item_text(qname: str | None, item: Any) -> str | None:
    """The text a run child contributes, or None when it contributes none."""
    if qname == W_T:
        value = getattr(item, "value", None)
        return value if isinstance(value, str) else ""
    if qname == W_SYM:
        return _sym_text(item)
    if qname is None:
        return None
    return _CONSTANT_TEXT.get(qname)


@dataclasses.dataclass(frozen=True, slots=True)
class Segment:
    """One text-bearing item of a run, with its place in the paragraph's text.

    Segments are computed fresh on every use --- nothing is cached, and an edit
    invalidates the indices --- so a caller that mutates the tree recomputes
    them before the next step.
    """

    item: Any
    """The item: a ``w:t``, ``w:tab``, ``w:br``, ``w:cr``, ``w:sym``, ..."""
    owner: list
    """The run's ``content`` list, which holds :attr:`item`."""
    index: int
    """The item's index in :attr:`owner`."""
    run: Any
    """The ``w:r``."""
    run_owner: list
    """The list holding the run: a paragraph's, a hyperlink's, an insertion's."""
    run_index: int
    """The run's index in :attr:`run_owner`."""
    text: str
    """The characters this item contributes."""
    start: int
    """The offset of the first character, in code points from the paragraph's start."""
    end: int
    """One past the last: ``start + len(text)``."""
    editable: bool
    """True for a ``w:t``: text that can be changed in place. False in a kept branch."""
    revision: str | None = None
    """``"ins"``, ``"del"``, ``"moveFrom"``, ``"moveTo"`` when the run is in one."""
    branch: Any = None
    """The outermost run- or paragraph-level ``mc:AlternateContent`` this text is
    read from, as the branch Word shows (CR-011 Phase F); None for the
    paragraph's own text. Such text reads, and is not edited in place."""
    outer: Any = None
    """For text read from a branch, the item outside every branch that holds it:
    the ``w:r`` a run-level element is in, or a paragraph-level element itself."""
    outer_owner: list | None = None
    """The list holding :attr:`outer`."""
    outer_index: int = -1
    """:attr:`outer`'s index in :attr:`outer_owner`."""

    def site(self) -> tuple[Any, list, int]:
        """``(item, list, index)``: what an insertion beside this segment goes next to.

        The run and the list holding it; for text read from a kept branch, the
        item outside every branch that holds it (:attr:`outer`), so that what is
        inserted beside the element's text goes beside the element and into no
        branch (CR-011 Phase F).
        """
        if self.branch is None or self.outer_owner is None:
            return self.run, self.run_owner, self.run_index
        return self.outer, self.outer_owner, self.outer_index

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the text, the offsets and the item's name."""
        return {
            "text": self.text,
            "start": self.start,
            "end": self.end,
            "item": element_name(self.item),
            "editable": self.editable,
            "revision": self.revision,
            "branch": self.branch is not None,
        }


def segments_of(container: Any, *, view: str = "accepted") -> list[Segment]:
    """The text segments of a paragraph (or any run holder), in order.

    A run- or paragraph-level ``mc:AlternateContent`` is read as the branch
    Word shows (:func:`docx4j_py.child.mce_branch`; CR-011 Phase F): its
    segments carry the element as :attr:`Segment.branch` and are not editable.

    Args:
        container: a ``w:p``, or anything :func:`run_items_of` knows.
        view: ``"accepted"`` (the default: ``w:ins`` and ``w:moveTo`` count,
            ``w:del`` and ``w:moveFrom`` do not) or ``"original"``, the other
            way round.

    Returns:
        One :class:`Segment` per text-bearing run child, in document order,
        with ``start``/``end`` running over the whole text.
    """
    walker = _Walker(view)
    walker.visit(run_items_of(container), None, None, None)
    return walker.out


class _Walker:
    """:func:`segments_of`'s walk: run-level items, runs, and Word's branch of a kept element."""

    __slots__ = ("out", "position", "view")

    def __init__(self, view: str) -> None:
        self.view = view
        self.out: list[Segment] = []
        self.position = 0

    def visit(self, items: list | None, revision: str | None, kept: Any, outer: Any) -> None:
        """A run-level list: runs, revision holders, run holders, kept elements."""
        if not items:
            return
        for run_index, element in enumerate(items):
            if isinstance(element, str):
                continue
            name = element_name(element) if kept is None else _qname(element)
            if name == W_R:
                content = getattr(element, "content", None)
                if not isinstance(content, list):
                    content = _wildcard_items(element) if kept is not None else None
                    if content is None:
                        continue
                self.run(content, element, items, run_index, revision, kept, outer)
                continue
            if is_alternate_content(element):
                # a paragraph-level element kept whole: its runs, as the branch Word shows
                self.visit(
                    _branch_items(mce_branch(element)),
                    revision,
                    element if kept is None else kept,
                    (element, items, run_index) if outer is None else outer,
                )
                continue
            shown = _REVISIONS.get(name or "")
            if shown is not None:
                if shown != self.view:
                    continue
                kind = (name or "").rpartition("}")[2]
                self.visit(self._holder_items(element, kept), kind, kept, outer)
            elif name in RUN_HOLDER_NAMES:
                self.visit(self._holder_items(element, kept), revision, kept, outer)

    @staticmethod
    def _holder_items(element: Any, kept: Any) -> list | None:
        found = run_items_of(element)
        if found is None and kept is not None:
            found = _wildcard_items(element)
        return found

    def run(
        self,
        content: list,
        run: Any,
        items: list,
        run_index: int,
        revision: str | None,
        kept: Any,
        outer: Any,
    ) -> None:
        """A run's children, or a run-level element's branch read as them."""
        for index, item in enumerate(content):
            if isinstance(item, str):
                continue
            if is_alternate_content(item):
                # a run-level element kept whole: the branch Word shows, as the run's own items
                inner = _branch_items(mce_branch(item))
                if inner:
                    self.run(
                        inner,
                        run,
                        items,
                        run_index,
                        revision,
                        item if kept is None else kept,
                        (run, items, run_index) if outer is None else outer,
                    )
                continue
            if kept is None:
                qname = element_name(item)
                if qname == W_DEL_TEXT:
                    value = getattr(item, "value", None)
                    text: str | None = value if isinstance(value, str) else ""
                else:
                    text = item_text(qname, item)
            else:
                qname = _qname(item)
                text = _wildcard_text(qname, item) if is_any_element(item) else (
                    _typed_text(qname, item)
                )
            if qname == W_DEL_TEXT and revision not in ("del", "moveFrom"):
                # deleted text is text in the original view only, and only inside
                # the revision that deleted it
                continue
            if text is None:
                continue
            position = self.position
            self.out.append(
                Segment(
                    item=item,
                    owner=content,
                    index=index,
                    run=run,
                    run_owner=items,
                    run_index=run_index,
                    text=text,
                    start=position,
                    end=position + len(text),
                    editable=kept is None and qname in (W_T, W_DEL_TEXT),
                    revision=revision,
                    branch=kept,
                    outer=None if outer is None else outer[0],
                    outer_owner=None if outer is None else outer[1],
                    outer_index=-1 if outer is None else outer[2],
                )
            )
            self.position = position + len(text)


def _typed_text(qname: str | None, item: Any) -> str | None:
    """A typed run child's text, ``w:delText`` included (its view is the caller's)."""
    if qname == W_DEL_TEXT:
        value = getattr(item, "value", None)
        return value if isinstance(value, str) else ""
    return item_text(qname, item)


def branch_text(element: Any, *, view: str = "accepted", run_level: bool = False) -> str:
    """The text a kept ``mc:AlternateContent`` gives its paragraph: the branch Word shows (CR-011 Phase F).

    `run_level` for an element inside a ``w:r``, whose branch holds run
    children; otherwise its branch holds runs. What :func:`segments_of` reads
    for it, for a walker that counts offsets on its own (the comment markers',
    a revision's collapsed place).
    """
    walker = _Walker(view)
    if run_level:
        walker.run([element], None, [], -1, None, None, None)
    else:
        walker.visit([element], None, None, None)
    return "".join(segment.text for segment in walker.out)


def branch_extents(segments: list[Segment]) -> list[tuple[int, int, Any]]:
    """``(start, end, element)`` for each kept element whose branch gives the paragraph text, in order."""
    found: dict[int, list[Any]] = {}
    order: list[list[Any]] = []
    for segment in segments:
        if segment.branch is None or not segment.text:
            continue
        entry = found.get(id(segment.branch))
        if entry is None:
            entry = [segment.start, segment.end, segment.branch]
            found[id(segment.branch)] = entry
            order.append(entry)
        else:
            entry[1] = segment.end
    return [(entry[0], entry[1], entry[2]) for entry in order]


def touches_branch_text(segments: list[Segment], start: int, end: int) -> bool:
    """Whether an edit of ``[start, end)`` would change text read from a kept branch (CR-011 Phase F).

    A span touches it when it holds any of that text; an insertion point
    (``start == end``) when it lies strictly inside one element's text. An
    insertion at either end of the element's text goes beside the element. The
    offsets are snapped to grapheme boundaries first, outwards, as an edit
    snaps them (CR-003 section 3.12).
    """
    extents = branch_extents(segments)
    if not extents:
        return False
    text = "".join(segment.text for segment in segments)
    start = snap_back(text, start)
    if end > start:
        end = snap_forward(text, end)
        return any(a < end and b > start for a, b, _ in extents)
    return any(a < start < b for a, b, _ in extents)


def refuse_branch_text(container: Any, start: int, end: int, operation: str = "this edit", *, segments: list[Segment] | None = None) -> None:
    """Refuse an edit of ``[start, end)`` that would change text read from a kept branch.

    CR-011 Phase F (section 15 item 3), CR-012 Phase A's rule: text a run- or
    paragraph-level ``mc:AlternateContent`` gives its paragraph is read as the
    branch Word shows, and an edit there would change that branch alone. Called
    by every verb before it changes anything.

    Raises:
        AddressError: ``address.read_only``.
    """
    found = segments if segments is not None else segments_of(container)
    if not touches_branch_text(found, start, end):
        return
    from docx4j_py.model.content.errors import AddressError

    where = f"offset {start}" if end <= start else f"{start}:{end}"
    raise AddressError(
        f"{operation} refused: {where} is in text read from a branch of an mc:AlternateContent "
        "the document keeps whole (the branch Word shows): an edit there would change that "
        "branch alone (CR-011 section 15, CR-012)",
        code="address.read_only",
        hint="edit beside the element (an insertion at either end of its text goes outside it), "
        "or load with LoadOptions(mce_preprocess=True), which replaces each element with the "
        "branch Word takes",
    )


def text_of_view(container: Any, *, view: str = "accepted") -> str:
    """The text of one run holder in one of the two views."""
    return "".join(segment.text for segment in segments_of(container, view=view))


#: The views ``get_text`` and ``text_budget`` read (CR-003 section 3.8), and
#: since CR-011 Phase B the text model's ``markup`` view.
VIEWS = ("accepted", "original", "markup")


def canonical_view(view: str) -> str:
    """A view's name, ``current`` read as ``accepted`` (WD REQ-401's synonym, CR-011 Phase B)."""
    return "accepted" if view == "current" else view


def require_view(view: str, views: tuple[str, ...] = VIEWS) -> None:
    """Refuse a view ``get_text`` cannot read, rather than return neither.

    The filter shows a revision holder only in the view it belongs to, so any
    other name would keep the text outside every insertion, deletion and move.
    ``"markup"`` is the text model's view (CR-011 Q10, Phase B:
    :func:`markup_text`), which in 0.1.0 a paragraph's and a body's
    ``view="markup"`` was not.

    Raises:
        ContentError: ``text.view_invalid`` for a view not in `views`.
    """
    if view in views:
        return
    taken = " or ".join(repr(v) for v in views)
    raise ContentError(
        f"get_text takes view={taken}, not {view!r}",
        code="text.view_invalid",
        hint="'accepted' reads the document as if every change were accepted",
    )


def markup_text(target: Any, *, text_boxes: bool = True) -> str:
    """The text model's ``markup`` view of a paragraph or a body, as plain text (CR-011 Phase B).

    Every run as stored, its revision written ``{++...++}`` or ``{--...--}``,
    each comment ``{>>...<<}`` at its reference with its replies (WD REQ-403,
    REQ-504, REQ-702); a line per paragraph, hidden text left out, a text box
    after its anchoring paragraph unless `text_boxes` is off. The model's, not
    Office JS's: :func:`docx4j_py.model.text.extract` is the block model under
    it.
    """
    from docx4j_py.model.text import extract, plain_text

    text = plain_text(extract(target, view="markup", text_boxes=text_boxes))
    return text.removesuffix("\n")


@dataclasses.dataclass(frozen=True, slots=True)
class Reading:
    """A body read by the text model (CR-011 Phase D): the result, and how its addresses map onto the body's.

    ``result`` is :func:`docx4j_py.model.text.extract`'s; ``prefix`` is what
    its blocks' ordinal addresses start with (``body``, ``header:rId4``, or
    ``content`` for a body that is not a whole story); ``package`` is the
    package it was read from, None for a detached body.
    """

    result: Any
    prefix: str
    package: Any

    def address_of(self, body_prefix: str, block: Any) -> str:
        """A block's address in the content API's form: its paraId first, else its ordinal.

        A block of the body's own story (its model address a plain ordinal
        under :attr:`prefix`) takes ``w14:`` and its ``w14:paraId`` where it
        has one, else the ordinal under the body's own prefix (`body_prefix`), which
        :meth:`~docx4j_py.model.content.body.Body.element_at` resolves (CR-003
        section 3.4). Any other block --- a text box's, a note's, a chart's, a
        chunk's --- keeps the model's address (REQ-519, ``body/3/textbox:0/1``),
        which ``element_at`` does not take (CR-011 Q9: read only).
        """
        address = block.address
        prefix = self.prefix
        if address.startswith(prefix + "/") and address[len(prefix) + 1 :].replace("/", "").isdigit():
            if block.para_id:
                return f"w14:{block.para_id}"
            return body_prefix + address[len(prefix) :]
        return address


def reading_of(body: Any, **options: Any) -> Reading:
    """The text model's reading of a :class:`~docx4j_py.model.content.body.Body` (CR-011 Phase D).

    The main document's body is read as the package is (docx4j
    ``Texts.of(pkg)``: the main story, its text boxes and its notes, the notes'
    numbers in reference order); a header's or a footer's as its part
    (``Texts.of(part)``); any other body --- a cell, a content control, a
    comment, a notes part --- as content in its package's view. Every part is
    read with ``read_contents``, so nothing is unmarshalled by the reading
    (REQ-204). `options` are :class:`~docx4j_py.model.text.TextOptions`' by name.
    """
    from docx4j_py.model.text import extract

    package = body.package
    part = body.part
    contents = getattr(part, "contents", None) if part is not None else None
    if package is not None and part is not None:
        if part is getattr(package, "main_document_part", None) and body.container is getattr(contents, "body", None):
            return Reading(extract(package, **options), "body", package)
        if type(part).__name__ in ("HeaderPart", "FooterPart") and body.container is contents:
            result = extract(part, package=package, **options)
            prefix = result.blocks[0].address.split("/", 1)[0] if result.blocks else body.prefix
            return Reading(result, prefix, package)
    return Reading(extract(body, **options), "content", package)


def runs_of(container: Any, *, view: str = "accepted") -> list[Any]:
    """Every ``w:r`` of a run holder, nested ones included, in document order.

    Runs with no text are included --- a run holding only a ``w:drawing`` is a
    run --- which is why this is not derived from :func:`segments_of`.
    """
    out: list[Any] = []

    def visit(items: list | None) -> None:
        if not items:
            return
        for element in items:
            name = element_name(element)
            if name == W_R:
                out.append(element)
                continue
            shown = _REVISIONS.get(name or "")
            if shown is not None:
                if shown == view:
                    visit(run_items_of(element))
            elif name in RUN_HOLDER_NAMES:
                visit(run_items_of(element))

    visit(run_items_of(container))
    return out


# ---------------------------------------------------------------------------
# block structure
# ---------------------------------------------------------------------------

_sole_list_cache: dict[type, str | None] = {}


def _holds_blocks(cls: type) -> bool:
    """True when instances of `cls` are themselves containers of block content."""
    try:
        fields = dataclasses.fields(cls)
    except TypeError:
        return False
    for field in fields:
        if field.name != "content":
            continue
        factory = field.default_factory
        return factory is not dataclasses.MISSING and isinstance(factory, type)
    return False


def _sole_list_field(value: Any) -> str | None:
    """The one repeated element field of a class that has no ``content`` list.

    ``w:footnotes`` keeps its notes in ``footnote`` and ``w:comments`` its
    comments in ``comment``; both are a container of blocks all the same, so a
    ``Body`` over such a part works without naming either field here.

    The field's own items have to be containers of blocks --- a ``w:footnote``
    and a ``w:comment`` have a ``content`` list, a ``w:style`` does not --- so
    that ``w:styles``, which likewise keeps one repeated element, is correctly
    reported as holding no block content at all.
    """
    cls = value.__class__
    try:
        return _sole_list_cache[cls]
    except KeyError:
        pass
    found: str | None = None
    try:
        fields = dataclasses.fields(cls)
    except TypeError:
        fields = ()  # type: ignore[assignment]
    for field in fields:
        if field.metadata.get("type") not in (None, "Element"):
            continue
        factory = field.default_factory
        if factory is dataclasses.MISSING or not isinstance(factory, type):
            continue
        if not issubclass(factory, list):
            continue
        if found is not None:  # more than one: no answer
            found = None
            break
        if not all(_holds_blocks(item) for item in _field_types(field)):
            continue
        found = field.name
    _sole_list_cache[cls] = found
    return found


def _field_types(field: Any) -> list[type]:
    """The classes a repeated element field holds, from its type annotation."""
    import typing

    annotation = field.type
    if isinstance(annotation, str):
        try:
            module = sys.modules["docx4j_py.wml"]
            annotation = eval(annotation, vars(module))
        except Exception:  # noqa: BLE001 - an annotation we cannot resolve holds nothing
            return []
    args = typing.get_args(annotation)
    return [arg for arg in args if isinstance(arg, type)]


#: The classes of an ``mc:AlternateContent``'s branches.
_MC_BRANCHES = frozenset({"AlternateContentChoice", "AlternateContentFallback"})
#: ...and of the element and its branches: what stands between a branch's child and its container.
_MC_CLASSES = frozenset({"AlternateContent", *_MC_BRANCHES})
#: The generated classes of the block, row and cell containers (docx4j ``McSplice.isBlockContainer``,
#: ``isRowContainer`` and ``isCellContainer``), whose ``mc:AlternateContent`` a reader of structure
#: reads through: at block level since CR-012 Phase A, among a table's rows (a ``w:tbl``, a row-level
#: control's content, row-level custom XML) and a row's cells since Phase A2, as docx4j's CR-051
#: phase 2 and the text model read them.
_CONTAINERS = frozenset(
    {
        "Body1",
        "Hdr",
        "Ftr",
        "Tc",
        "CTFtnEdn",
        "CommentsComment",
        "CTTxbxContent",
        "SdtContentBlock",
        "CTCustomXmlBlock",
        "Tbl",
        "CTSdtContentRow",
        "CTCustomXmlRow",
        "Tr",
        "CTSdtContentCell",
        "CTCustomXmlCell",
    }
)


def _block_level(ac: Any) -> bool:
    """Whether an ``mc:AlternateContent`` stands in a block, row or cell container, through any ``mc`` wrappers.

    The name is Phase A's, when only a block container's element was read through.
    """
    node = getattr(ac, "parent", None)
    for _ in range(256):
        if node is None:
            return False
        cls = node.__class__
        if cls.__module__ != "docx4j_py.mce" or cls.__name__ not in _MC_CLASSES:
            return cls.__module__ == "docx4j_py.wml" and any(
                base.__name__ in _CONTAINERS for base in cls.__mro__ if base.__module__ == "docx4j_py.wml"
            )
        node = getattr(node, "parent", None)
    return False


def in_kept_branch(element: Any) -> bool:
    """Whether an element lies in a branch of an ``mc:AlternateContent`` the tree keeps whole.

    CR-012 Phase A. At block level in ``preserve`` mode (the default), and at
    run and paragraph level always, the element is kept with every branch; a
    reader of structure reads the branch Word reads, and an edit there would
    change that branch alone, so the content API refuses it
    (``address.read_only``). Through the parent pointers.
    """
    node = getattr(element, "parent", None)
    for _ in range(256):
        if node is None:
            return False
        cls = node.__class__
        if cls.__module__ == "docx4j_py.mce" and cls.__name__ in _MC_BRANCHES:
            return True
        node = getattr(node, "parent", None)
    return False


def kept_content_under(element: Any) -> Any | None:
    """The first element of Word's branch of a kept block-, row- or cell-level ``mc:AlternateContent`` under `element`.

    CR-012 Phase A2. What a tracked deletion of a row, a table or a body would
    mark deleted is walked through :func:`block_children_of`, which reads such
    an element by the branch Word reads, so the marks would land in that
    branch alone; the verb asks here first and refuses (``address.read_only``)
    on what this returns. Since CR-011 Phase F also the first text item a
    paragraph's run- or paragraph-level element gives its paragraph's text,
    which a tracked deletion of that text would mark in Word's branch alone.
    None when there is no such element, or Word's branch of each is empty.
    """
    stack = [element]
    while stack:
        node = stack.pop()
        if node is not element and is_alternate_content(node) and block_list_of(node) is not None:
            items = block_children_of(node) or []
            if items:
                return items[0]
        if element_name(node) == W_P:
            # a paragraph's run- and paragraph-level elements: a tracked deletion of
            # its text would mark Word's branch alone deleted (CR-011 Phase F)
            kept = next((s for s in segments_of(node) if s.branch is not None and s.text), None)
            if kept is not None:
                return kept.item
            continue
        children = block_children_of(node)
        if children is not None:
            stack.extend(reversed(children))
    return None


def block_list_of(value: Any) -> tuple[Any, str] | None:
    """The object and field name that hold a container's block-level list.

    The pair rather than the list itself, because a caller that inserts needs
    the owner for the parent pointers, and a caller that validates needs the
    class to ask the metadata about (:func:`docx4j_py.model.content.body`).

    An ``mc:AlternateContent`` kept whole at block level answers with the
    branch Word reads (:func:`docx4j_py.child.mce_branch`, CR-012 Phase A), so
    that a reader of structure --- :attr:`Body.paragraphs` and the rest ---
    reads what Word shows; a view of what is there is read-only
    (:func:`in_kept_branch`). So does one among a table's rows or a row's
    cells (CR-012 Phase A2), so that :func:`rows_of` and :func:`cells_of`
    read the rows and cells Word shows, and a path through the element names
    the branch's index (``body/1/1/0``, the text model's address).
    """
    if is_alternate_content(value):
        if not _block_level(value):
            return None
        branch = mce_branch(value)
        if branch is None or is_any_element(branch):
            return None
        return branch, "any_element"
    cls = value.__class__
    if cls.__module__ == "docx4j_py.mce" and cls.__name__ in _MC_BRANCHES:
        # the branch itself, so that a path through the element names the branch's index
        # (body/2/0) and the element's own step is transparent; any branch, Word's or not,
        # so that an address is never another branch's element
        ac = getattr(value, "parent", None)
        if ac is None or not _block_level(ac) or mce_branch(ac) is not value:
            return None
        return value, "any_element"
    inner = getattr(value, "sdt_content", None)
    if inner is not None:
        value = inner
    if isinstance(getattr(value, "content", None), list):
        return value, "content"
    name = _sole_list_field(value)
    if name is not None and isinstance(getattr(value, name, None), list):
        return value, name
    return None


def block_children_of(value: Any) -> list | None:
    """The block-level list a container keeps, or None when it keeps none.

    A content control of any form answers with its ``w:sdtContent``'s list, so
    a traversal never has to know the four ``w:sdt`` classes apart (CR-003
    section 4: "the four ``sdtContent`` levels are skipped in paths"). A table
    answers with its rows, a row with its cells and a cell with its blocks,
    because in this model they are all the same ``content`` field --- which is
    what makes descending into a table one line rather than three helpers.
    """
    found = block_list_of(value)
    return None if found is None else getattr(found[0], found[1])


def _wrapped(items: Any, wanted: str, out: list[tuple[Any, list]]) -> None:
    """Collect ``(element, the live list holding it)`` for one element name.

    A row- or cell-level ``w:sdt`` (an OpenDoPE repeat) or ``w:customXml`` is
    transparent and is descended into, so the container reported for a row
    inside a repeat is the **control's** list, which is where an insert beside
    it belongs.
    """
    for item in items:
        name = element_name(item)
        if name == wanted:
            out.append((item, items))
        elif name in _ROW_WRAPPERS or is_alternate_content(item):
            # a row- or cell-level mc:AlternateContent kept whole: the branch Word
            # reads, whose list holds what is found in it (CR-012 Phase A2)
            children = block_children_of(item)
            if children is not None:
                _wrapped(children, wanted, out)


def rows_of(table: Any) -> list[tuple[Any, list]]:
    """A table's ``w:tr``\\ s and the list holding each, controls descended into.

    CR-003 section 4: "rows and cells descend into row- and cell-level content
    controls (an OpenDoPE repeat wraps its ``w:tr`` in a ``w:sdt``), so ``rows``
    is what Word shows". A row-level ``mc:AlternateContent`` kept whole is read
    by the branch Word reads, and a row there is held by the branch's list, so
    its view is read-only (CR-012 Phase A2).
    """
    return _collected(table, W_TR)


def cells_of(row: Any) -> list[tuple[Any, list]]:
    """A row's ``w:tc``\\ s and the list holding each, controls and a kept cell-level element descended into."""
    return _collected(row, W_TC)


def _collected(container: Any, wanted: str) -> list[tuple[Any, list]]:
    out: list[tuple[Any, list]] = []
    items = block_children_of(container)
    if items is not None:
        _wrapped(items, wanted, out)
    return out


def child_qname(parent: Any, child: Any) -> str | None:
    """The element name `child` is written under inside `parent`."""
    for qname, item in iter_children(parent, mce="all"):
        if item is child:
            return qname
    return element_name(child)


# ---------------------------------------------------------------------------
# grapheme clusters (CR-003 section 3.12, decided question 10)
# ---------------------------------------------------------------------------

_ZWJ = "‍"
_MARKS = frozenset({"Mn", "Mc", "Me"})


def _is_variation_selector(ch: str) -> bool:
    code = ord(ch)
    return 0xFE00 <= code <= 0xFE0F or 0xE0100 <= code <= 0xE01EF


def _is_regional_indicator(ch: str) -> bool:
    return 0x1F1E6 <= ord(ch) <= 0x1F1FF


def _is_emoji_modifier(ch: str) -> bool:
    return 0x1F3FB <= ord(ch) <= 0x1F3FF


def _is_extend(ch: str) -> bool:
    return unicodedata.category(ch) in _MARKS or ch == _ZWJ or _is_variation_selector(ch)


def is_grapheme_boundary(text: str, index: int) -> bool:
    """True when `index` may be split without cutting a grapheme cluster in two.

    CR-003 section 3.12 and decided question 10: ``unicodedata``'s combining
    classes and general categories, plus the cases they do not cover ---

    * a combining mark, a variation selector or an emoji skin-tone modifier
      never begins a cluster;
    * the zero-width joiner binds both ways, which is what holds an emoji
      sequence such as U+1F468 ZWJ U+1F469 ZWJ U+1F467 together;
    * two regional indicators make one flag, so only every second one is a
      boundary;
    * a virama (canonical combining class 9) binds the consonant after it, so
      a Devanagari conjunct such as ``स्ते`` is one cluster;
    * CR LF is one cluster.

    ``0`` and ``len(text)`` are always boundaries.
    """
    if index <= 0 or index >= len(text):
        return True
    previous = text[index - 1]
    current = text[index]
    if previous == "\r" and current == "\n":
        return False
    if current == _ZWJ or previous == _ZWJ:
        return False
    if unicodedata.category(current) in _MARKS or unicodedata.combining(current):
        return False
    if _is_variation_selector(current) or _is_emoji_modifier(current):
        return False
    if _is_regional_indicator(current) and _is_regional_indicator(previous):
        run = 0
        position = index
        while position > 0 and _is_regional_indicator(text[position - 1]):
            run += 1
            position -= 1
        return run % 2 == 0
    # Indic conjunct: consonant, extenders, virama, extenders, consonant.
    position = index - 1
    while position >= 0 and _is_extend(text[position]):
        if unicodedata.combining(text[position]) == 9:
            return False
        position -= 1
    return True


def snap_back(text: str, index: int) -> int:
    """The largest grapheme boundary at or before `index`."""
    index = max(0, min(index, len(text)))
    while not is_grapheme_boundary(text, index):
        index -= 1
    return index


def snap_forward(text: str, index: int) -> int:
    """The smallest grapheme boundary at or after `index`."""
    index = max(0, min(index, len(text)))
    while not is_grapheme_boundary(text, index):
        index += 1
    return index


def grapheme_clusters(text: str) -> list[str]:
    """`text` split into grapheme clusters.

    The ``regex`` module's ``\\X`` when that module is installed --- it carries
    the whole of UAX #29 --- and :func:`is_grapheme_boundary` otherwise, which
    is CR-003 decided question 10: no dependency is added for this.
    """
    try:
        import regex  # type: ignore[import-not-found]
    except ImportError:
        out: list[str] = []
        start = 0
        for index in range(1, len(text) + 1):
            if is_grapheme_boundary(text, index):
                out.append(text[start:index])
                start = index
        return out
    return regex.findall(r"\X", text)


# ---------------------------------------------------------------------------
# splitting a run at an offset
# ---------------------------------------------------------------------------


def split_at(
    paragraph_element: Any, offset: int, *, prefer: str = "back", made: list | None = None
) -> int:
    """Split the run at a text offset so that ``[offset, ...)`` begins a run.

    The offset is snapped to a grapheme boundary first (CR-003 section 3.12):
    ``prefer="back"`` to the start of the cluster it falls in, ``"forward"`` to
    the end of it. An offset that is already a run boundary splits nothing.

    Args:
        paragraph_element: the ``w:p``.
        offset: a code-point offset into the paragraph's accepted-view text.
        prefer: ``"back"`` or ``"forward"``, which way to snap.
        made: a list the ``(head, tail, item)`` triple is appended to, for a
            caller that has to know which two runs a split made and whether the
            split is what put ``xml:space`` on the head's ``w:t``
            (CR-003 section 16.12); `item` is None when the document had it.

    Returns:
        The offset actually used, which is the one a caller must go on with.
    """
    from docx4j_py.child import ChildList, deep_copy, link_parents
    from docx4j_py.wml import R, el

    text = text_of_view(paragraph_element)
    offset = snap_back(text, offset) if prefer == "back" else snap_forward(text, offset)

    segment = next(
        (s for s in segments_of(paragraph_element) if s.editable and s.start < offset < s.end),
        None,
    )
    if segment is None:
        return offset

    at = offset - segment.start
    space_before = getattr(segment.item, "space", None)
    set_text(segment.item, segment.text[:at])
    added_space = space_before is None and getattr(segment.item, "space", None) is not None
    tail = segment.text[at:]
    rest = list(segment.owner[segment.index + 1 :])
    del segment.owner[segment.index + 1 :]
    r_pr = getattr(segment.run, "r_pr", None)
    second = R(content=ChildList([el.t(tail), *rest]))
    if r_pr is not None:
        second.r_pr = deep_copy(r_pr, second)
    # the halves of a split run keep the original's own attributes, so that they
    # are identical but for their text (CR-003 section 16.12)
    for name in ("rsid_rpr", "rsid_del", "rsid_r"):
        value = getattr(segment.run, name, None)
        if value is not None:
            setattr(second, name, value)
    segment.run_owner.insert(segment.run_index + 1, second)
    link_parents(second)
    second.parent = getattr(segment.run, "parent", None) or paragraph_element
    if made is not None:
        made.append((segment.run, second, segment.item if added_space else None))
    return offset


def set_text(item: Any, value: str) -> None:
    """Write a ``w:t``'s characters, with ``xml:space="preserve"`` where needed.

    Word writes the attribute whenever the value begins or ends with a space or
    holds two in a row; ``el.t`` does the same for a new one, and this keeps an
    edited one in step. An attribute the document already had is **left alone**
    when the new value does not need it --- Word writes a redundant one often
    enough --- and only :func:`~docx4j_py.model.content.tracking.merge_runs`,
    putting two halves back together, recomputes it (CR-003 section 16.12).
    """
    item.value = value
    if value != value.strip() or "  " in value:
        item.space = "preserve"


# ---------------------------------------------------------------------------
# search (Office JS Word.SearchOptions)
# ---------------------------------------------------------------------------


def search_pattern(
    text: str,
    *,
    match_case: bool = False,
    match_whole_word: bool = False,
    match_wildcards: bool = False,
) -> re.Pattern[str]:
    """A compiled pattern for a search text under Office JS's search options.

    Args:
        text: what to look for; a literal unless `match_wildcards` is on.
        match_case: ``False`` (the default) searches case-insensitively, which
            is ``re.IGNORECASE`` (CR-003 section 3.12).
        match_whole_word: the match must begin and end on a Unicode word
            boundary.
        match_wildcards: Word's wildcards --- ``?`` one character, ``*`` any
            run, ``[a-z]`` a class, ``[!x]`` a negated one, ``<`` and ``>`` the
            start and end of a word, ``@`` one or more of what precedes,
            ``{n,m}`` a repeat, and ``\\`` to escape any of them.
    """
    if match_wildcards:
        out: list[str] = []
        index = 0
        while index < len(text):
            char = text[index]
            if char == "?":
                out.append(".")
            elif char == "*":
                out.append(".*?")
            elif char == "<":
                out.append(r"\b(?=\w)")
            elif char == ">":
                out.append(r"\b(?<=\w)")
            elif char == "@":
                out.append("+")
            elif char == "[":
                close = text.find("]", index)
                if close < 0:
                    out.append(r"\[")
                else:
                    group = text[index + 1 : close]
                    if group.startswith("!"):
                        group = "^" + group[1:]
                    out.append(f"[{group.replace(chr(92), chr(92) * 2)}]")
                    index = close
            elif char == "{":
                close = text.find("}", index)
                if close < 0:
                    out.append(r"\{")
                else:
                    out.append("{" + text[index + 1 : close] + "}")
                    index = close
            elif char == "\\" and index + 1 < len(text):
                index += 1
                out.append(re.escape(text[index]))
            else:
                out.append(re.escape(char))
            index += 1
        source = "".join(out)
    else:
        source = re.escape(text)
    if match_whole_word:
        source = rf"(?<!\w){source}(?!\w)"
    return re.compile(source, 0 if match_case else re.IGNORECASE)


def find_all(text: str, pattern: re.Pattern[str]) -> list[tuple[int, int]]:
    """Every ``[start, end)`` match of a pattern, empty matches skipped."""
    out: list[tuple[int, int]] = []
    position = 0
    while position <= len(text):
        match = pattern.search(text, position)
        if match is None:
            break
        if match.end() == match.start():
            position = match.start() + 1
            continue
        out.append((match.start(), match.end()))
        position = match.end()
    return out
