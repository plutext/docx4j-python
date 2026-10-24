"""A paragraph's text as segments, and the block structure under a container.

This is the piece every view rests on. A paragraph's text is not a string in the
tree: it is spread over runs, and over the hyperlinks, content controls, smart
tags and tracked insertions those runs sit in. :func:`segments_of` flattens it
into one list of :class:`Segment`, each mapping a slice of the accepted-view
text to the item that carries it and to the lists that hold that item and its
run --- which is everything an edit needs to change the text in place, split a
run at an offset, or apply formatting to exactly one span.

**Office JS's characters** (CR-013 Phase A, decided 2026-10-23, Q1; docx4j-core-ts
CR-012 phase 4 at ``9307bdb``, held to Word 365 by its checks 48, 48b and 51). A
segment's text is Word's own character stream, as Office JS's ``paragraph.text``
gives it in Word's default view (Simple Markup, hidden text not shown):

=============================================  ============================
item                                           Office JS's text
=============================================  ============================
``w:t``                                        its value
``w:tab``                                      ``\\t``
``w:br`` (a line break)                        U+000B
``w:br w:type="page"``, ``"column"``           U+000C, U+000E
``w:cr``                                       U+000D
``w:ptab``                                     ``0``
``w:noBreakHyphen``, ``w:softHyphen``          U+001E, U+001F
``w:sym`` (any font, any code)                 ``(``
a note reference, a note's own mark            U+0002
an equation (``m:oMath``, ``m:oMathPara``)     Word's text of it, one segment
                                               (:mod:`.math_text`)
a hidden run (its own ``w:vanish``, else its   nothing (:func:`hidden_runs_of`)
styles')
a run in a content control showing its         nothing (:func:`hidden_runs_of`;
placeholder (``w:showingPlcHdr``)              CR-016 Phase B)
a comment reference, a field code, a picture   nothing
=============================================  ============================

:data:`OFFICE_CHARACTERS` is that table. The 0.1.0 characters --- docx4j's
``TextUtils``, :func:`docx4j_py.traversal.text_of`'s, a break ``\\n``, the hyphens
U+2011 and U+00AD, a symbol its code point, a note reference nothing, hidden text
read --- are :func:`item_text`'s, which the python-docx facade reads (CR-011 Q5),
and ``characters="docx4j"``'s, which a bound content control writes back into its
custom XML (CR-013 section 4.1). ``w:delText``, ``w:instrText`` and
``w:delInstrText`` are not text --- except that a ``w:delText`` *is* the text of
the original view, inside the ``w:del`` or ``w:moveFrom`` that removed it, which
is the only place it is read.

**The writers** (Q2, with section 11's amendment). Six of those characters are
not XML 1.0 characters, so a text read and written back maps each one to the
element it reads from (:func:`text_items`): U+000B a ``w:br``, U+000C a page
break, U+000E a column break, U+001E a ``w:noBreakHyphen``, U+001F a
``w:softHyphen``, ``\\t`` a ``w:tab``. ``\\n``, ``\\r`` and ``\\r\\n``
are Word's own (docx4j-core-ts check 55, its CR-002 section 60.4): every writer
through ``Paragraph.splice`` makes a new paragraph of each (:data:`PARAGRAPH_BREAKS`,
``\\r\\n`` one), and the one-paragraph writers (``insert_paragraph``,
``Paragraph.text = ...``) write ``\\n`` and ``\\r\\n`` as one ``w:br`` and
``\\r`` as a ``w:cr``. U+0002 and the other characters XML 1.0 excludes are
dropped at the write, as Word drops them (:func:`drop_unwritable`; check 55):
no text makes a note, and lxml's refusal at save stays as the guard behind it.

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
from collections.abc import Callable
from typing import Any

from docx4j_py.child import (
    is_alternate_content,
    is_any_element,
    iter_children,
    mce_branch,
)
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.values import is_on
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of
from docx4j_py.wml.sdt import sdt_property

__all__ = [
    "FIND_CODES",
    "OBJECT_REPLACEMENT",
    "OFFICE_CHARACTERS",
    "PARAGRAPH_BREAKS",
    "RUN_HOLDER_NAMES",
    "TEXT_ITEM_NAMES",
    "VIEWS",
    "WRITTEN_AS",
    "MarkedMatch",
    "MarkedSegments",
    "Reading",
    "Segment",
    "block_children_of",
    "block_list_of",
    "branch_extents",
    "branch_text",
    "canonical_view",
    "cells_of",
    "drop_unwritable",
    "find_all",
    "for_writer",
    "grapheme_clusters",
    "hidden_runs_of",
    "in_showing_placeholder",
    "is_block_container",
    "is_branch",
    "is_grapheme_boundary",
    "item_text",
    "kept_content_under",
    "markup_text",
    "office_as_model",
    "office_item_text",
    "office_matches",
    "paragraph_mark_pieces",
    "placeholder_control_of",
    "reading_of",
    "refuse_branch_text",
    "refuse_equation",
    "refuse_mark_replacement",
    "refuse_tracked_equation",
    "refuse_wildcard_mark_replacement",
    "require_view",
    "rows_of",
    "runs_of",
    "search_pattern",
    "segments_of",
    "set_text",
    "snap_back",
    "snap_forward",
    "split_at",
    "text_items",
    "text_of_view",
    "text_run",
    "touches_branch_text",
    "unwritable",
    "wildcard_matches_across",
]


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


W_P = _w("p")
W_R = _w("r")
W_T = _w("t")
W_SYM = _w("sym")
W_BR = _w("br")
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

#: Office JS's character for each run child that adds one (CR-013 Phase A, Q1;
#: core-ts ``officeItemText`` at ``9307bdb``, Word 365 check 48): the run child's
#: element name to its character, a ``w:br`` by its ``w:type``. ``w:t`` gives its
#: value and an equation Word's text of it; everything else gives nothing.
OFFICE_CHARACTERS: dict[str, str] = {
    _w("tab"): "\t",
    W_BR: "\x0b",
    f"{W_BR}@page": "\x0c",
    f"{W_BR}@column": "\x0e",
    _w("cr"): "\r",
    _w("ptab"): "0",
    _w("noBreakHyphen"): "\x1e",
    _w("softHyphen"): "\x1f",
    W_SYM: "(",
    _w("footnoteReference"): "\x02",
    _w("endnoteReference"): "\x02",
    _w("footnoteRef"): "\x02",
    _w("endnoteRef"): "\x02",
}

#: The run children an Office JS reader lists as a place of no text (core-ts
#: ``OBJECT_ITEMS``): a picture's place, which ``search`` reads as U+FFFC.
_OBJECT_NAMES = frozenset({_w("drawing"), _w("pict"), _w("object")})

#: What the text model writes for each item whose Office JS character differs
#: (``model/text/extractor.py``'s ``_CONSTANT``, and a ``w:sym``'s code point):
#: Q11's character map, under which a text model run's ``office_offset`` agrees
#: with ``Paragraph.text`` (:func:`office_as_model`).
_MODEL_CHARACTERS: dict[str, str] = {
    _w("tab"): "\t",
    _w("ptab"): "\t",
    W_BR: "\n",
    _w("cr"): "\n",
    _w("noBreakHyphen"): "\u2011",
    _w("softHyphen"): "\u00ad",
}

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

#: An equation at run-holder level (CR-013 Phase A, :mod:`.math_text`).
_MATH_NAMES = frozenset(
    {
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath",
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara",
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


def _wildcard_text(qname: str | None, node: Any, office: bool = True) -> str | None:
    """The text a wildcard run child contributes (a kept branch's ``AnyElement``).

    In Office JS's characters (:func:`office_item_text`), or with `office`
    false in docx4j's (:func:`item_text`).
    """
    if qname in (W_T, W_DEL_TEXT):
        value = getattr(node, "text", None)
        return value if isinstance(value, str) else ""
    if office:
        return office_item_text(qname, node)
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
    """The text a run child contributes in docx4j's characters, or None when it contributes none.

    docx4j's ``TextUtils``, :func:`docx4j_py.traversal.text_of`'s: a ``w:br``
    and a ``w:cr`` ``\\n``, the hyphens U+2011 and U+00AD, a ``w:sym`` the
    character its ``w:char`` names. Not Office JS's (:func:`office_item_text`):
    the python-docx facade's ``run_text`` reads these (CR-011 Q5, CR-013 section
    2.5 item 4).
    """
    if qname == W_T:
        value = getattr(item, "value", None)
        return value if isinstance(value, str) else ""
    if qname == W_SYM:
        return _sym_text(item)
    if qname is None:
        return None
    return _CONSTANT_TEXT.get(qname)


def _break_type(item: Any) -> str | None:
    """A ``w:br``'s ``w:type``, typed or wildcard."""
    kind = getattr(item, "type_value", None)
    if kind is None:
        attributes = getattr(item, "attributes", None)
        if isinstance(attributes, dict):
            kind = attributes.get(_w("type")) or attributes.get("type")
    return getattr(kind, "value", kind)


def office_item_text(qname: str | None, item: Any) -> str | None:
    """The text a run child contributes in Office JS's characters, or None when it contributes none.

    :data:`OFFICE_CHARACTERS` (CR-013 Phase A): a ``w:t`` its value, a
    ``w:br`` U+000B (U+000C a page break, U+000E a column break), a ``w:cr``
    U+000D, a ``w:ptab`` ``0``, the hyphens U+001E and U+001F, any ``w:sym``
    ``(``, a note reference and a note's own mark U+0002.
    """
    if qname == W_T:
        value = getattr(item, "value", None)
        if value is None and is_any_element(item):
            value = getattr(item, "text", None)
        return value if isinstance(value, str) else ""
    if qname == W_BR:
        kind = _break_type(item)
        if kind in ("page", "column"):
            return OFFICE_CHARACTERS[f"{W_BR}@{kind}"]
        return OFFICE_CHARACTERS[W_BR]
    if qname is None:
        return None
    return OFFICE_CHARACTERS.get(qname)


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
    math: bool = False
    """An equation (``m:oMath``, ``m:oMathPara``): one segment of Word's text of
    it (:mod:`.math_text`). :attr:`item` and :attr:`run` are the equation and
    :attr:`owner` the list holding it; an edit may take it whole, untracked, and
    never part of it (:func:`refuse_equation`). CR-013 Phase A."""
    hidden: bool = False
    """An item of a hidden run --- its text, or a drawing in it --- listed only by
    ``segments_of(objects=True)``: no text and no offset of its own, so that a
    deletion of a span takes the hidden run strictly inside it. CR-013 Phase A."""
    object: bool = False
    """A place of no text listed only by ``segments_of(objects=True)``: a drawn
    object (``w:drawing``, ``w:pict``, ``w:object``), which ``search`` reads as
    U+FFFC, or a hidden run's item (:attr:`hidden`)."""

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


def segments_of(
    container: Any,
    *,
    view: str = "accepted",
    hidden: Callable[[Any], bool] | None = None,
    objects: bool = False,
    characters: str = "office",
    visit: Callable[..., None] | None = None,
) -> list[Segment]:
    """The text segments of a paragraph (or any run holder), in order, in Office JS's characters.

    A run- or paragraph-level ``mc:AlternateContent`` is read as the branch
    Word shows (:func:`docx4j_py.child.mce_branch`; CR-011 Phase F): its
    segments carry the element as :attr:`Segment.branch` and are not editable.
    An equation is one segment of Word's text of it (:attr:`Segment.math`).

    Args:
        container: a ``w:p``, or anything :func:`run_items_of` knows.
        view: ``"accepted"`` (the default: ``w:ins`` and ``w:moveTo`` count,
            ``w:del`` and ``w:moveFrom`` do not) or ``"original"``, the other
            way round.
        hidden: which runs are hidden (:func:`hidden_runs_of`): their text has
            no offset, as Office JS leaves hidden text out (core-ts check 48b).
            None: no run is hidden.
        objects: also list each drawn object and each hidden run's text item as
            a segment of no text where it stands (:attr:`Segment.object`,
            :attr:`Segment.hidden`), for a deletion of a span and for
            ``search``'s U+FFFC.
        characters: ``"office"`` (the default, :data:`OFFICE_CHARACTERS`) or
            ``"docx4j"``, the 0.1.0 characters (:func:`item_text`: a break
            ``\\n``, the hyphens U+2011 and U+00AD, a symbol its code point, a
            note reference and an equation nothing), which a bound content
            control writes back into its custom XML (CR-013 section 4.1).
        visit: called ``visit(item, offset, owner, run, run_owner, branch)``
            for every element the walk passes that adds no text --- each element
            of a run holder's list before it is entered (``run`` and
            ``run_owner`` None), and each run child that adds none (a comment
            reference, a field character, a drawing, a hidden run's text) ---
            with the offset it stands at: what gives a comment marker or a
            revision its offset in the same text (core-ts's
            ``SegmentOptions.visit``). `branch` is the kept
            ``mc:AlternateContent`` it is read from, or None.

    Returns:
        One :class:`Segment` per text-bearing run child, in document order,
        with ``start``/``end`` running over the whole text.
    """
    walker = _Walker(view, hidden=hidden, objects=objects, office=characters != "docx4j", callback=visit)
    walker.visit(run_items_of(container), None, None, None)
    return walker.out


class _Walker:
    """:func:`segments_of`'s walk: run-level items, runs, and Word's branch of a kept element."""

    __slots__ = ("callback", "hidden", "objects", "office", "out", "position", "view")

    def __init__(
        self,
        view: str,
        *,
        hidden: Callable[[Any], bool] | None = None,
        objects: bool = False,
        office: bool = True,
        callback: Callable[..., None] | None = None,
    ) -> None:
        self.view = view
        self.hidden = hidden
        self.objects = objects
        self.office = office
        self.callback = callback
        self.out: list[Segment] = []
        self.position = 0

    def visit(self, items: list | None, revision: str | None, kept: Any, outer: Any) -> None:
        """A run-level list: runs, revision holders, run holders, kept elements, equations."""
        if not items:
            return
        callback = self.callback
        for run_index, element in enumerate(items):
            if isinstance(element, str):
                continue
            if callback is not None:
                callback(element, self.position, items, None, None, kept)
            name = element_name(element) if kept is None else _qname(element)
            if name == W_R:
                content = getattr(element, "content", None)
                if not isinstance(content, list):
                    content = _wildcard_items(element) if kept is not None else None
                    if content is None:
                        continue
                hidden = self.hidden is not None and not is_any_element(element) and self.hidden(element)
                self.run(content, element, items, run_index, revision, kept, outer, hidden)
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
            if self.office and name in _MATH_NAMES:
                self.math(element, items, run_index, revision, kept, outer)
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

    def math(self, element: Any, items: list, index: int, revision: str | None, kept: Any, outer: Any) -> None:
        """An equation: one segment of Word's text of it (:mod:`.math_text`; CR-013 Phase A)."""
        from docx4j_py.model.content.math_text import math_text_of

        text = math_text_of(element, self.view)
        if not text:
            return
        position = self.position
        self.out.append(
            Segment(
                item=element,
                owner=items,
                index=index,
                run=element,
                run_owner=items,
                run_index=index,
                text=text,
                start=position,
                end=position + len(text),
                editable=False,
                revision=revision,
                branch=kept,
                outer=None if outer is None else outer[0],
                outer_owner=None if outer is None else outer[1],
                outer_index=-1 if outer is None else outer[2],
                math=True,
            )
        )
        self.position = position + len(text)

    def run(
        self,
        content: list,
        run: Any,
        items: list,
        run_index: int,
        revision: str | None,
        kept: Any,
        outer: Any,
        hidden: bool = False,
    ) -> None:
        """A run's children, or a run-level element's branch read as them."""
        callback = self.callback
        office = self.office
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
                        hidden,
                    )
                continue
            if kept is None:
                qname = element_name(item)
                if qname == W_DEL_TEXT:
                    value = getattr(item, "value", None)
                    text: str | None = value if isinstance(value, str) else ""
                else:
                    text = office_item_text(qname, item) if office else item_text(qname, item)
            else:
                qname = _qname(item)
                text = _wildcard_text(qname, item, office) if is_any_element(item) else (
                    _typed_text(qname, item, office)
                )
            if qname == W_DEL_TEXT and revision not in ("del", "moveFrom"):
                # deleted text is text in the original view only, and only inside
                # the revision that deleted it
                if callback is not None:
                    callback(item, self.position, content, run, items, kept)
                continue
            if text is None or hidden:
                # no text of its own, or a hidden run's: no offset (CR-013 Phase A)
                if callback is not None:
                    callback(item, self.position, content, run, items, kept)
                if self.objects and ((hidden and bool(text)) or qname in _OBJECT_NAMES):
                    position = self.position
                    self.out.append(
                        Segment(
                            item=item,
                            owner=content,
                            index=index,
                            run=run,
                            run_owner=items,
                            run_index=run_index,
                            text="",
                            start=position,
                            end=position,
                            editable=False,
                            revision=revision,
                            branch=kept,
                            outer=None if outer is None else outer[0],
                            outer_owner=None if outer is None else outer[1],
                            outer_index=-1 if outer is None else outer[2],
                            # a hidden run's item, a drawing in it too: Word's default
                            # view does not show it, so search's ^g passes it by
                            hidden=hidden,
                            object=True,
                        )
                    )
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


def _typed_text(qname: str | None, item: Any, office: bool = True) -> str | None:
    """A typed run child's text, ``w:delText`` included (its view is the caller's)."""
    if qname == W_DEL_TEXT:
        value = getattr(item, "value", None)
        return value if isinstance(value, str) else ""
    return office_item_text(qname, item) if office else item_text(qname, item)


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


def refuse_branch_text(
    container: Any,
    start: int,
    end: int,
    operation: str = "this edit",
    *,
    segments: list[Segment] | None = None,
    hidden: Callable[[Any], bool] | None = None,
) -> None:
    """Refuse an edit of ``[start, end)`` that would change text read from a kept branch.

    CR-011 Phase F (section 15 item 3), CR-012 Phase A's rule: text a run- or
    paragraph-level ``mc:AlternateContent`` gives its paragraph is read as the
    branch Word shows, and an edit there would change that branch alone. Called
    by every verb before it changes anything.

    Raises:
        AddressError: ``address.read_only``.
    """
    found = segments if segments is not None else segments_of(container, hidden=hidden)
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


def text_of_view(
    container: Any,
    *,
    view: str = "accepted",
    hidden: Callable[[Any], bool] | None = None,
    characters: str = "office",
) -> str:
    """The text of one run holder in one of the two views, in Office JS's characters.

    `hidden` and `characters` are :func:`segments_of`'s.
    """
    return "".join(
        segment.text
        for segment in segments_of(container, view=view, hidden=hidden, characters=characters)
    )


def _on(flag: Any) -> bool:
    """An on/off property as Word reads it (CR-018): present, and not exactly ``false``, ``off`` or ``0``."""
    return is_on(flag)


def _shows_placeholder(node: Any) -> bool:
    """Whether `node` is a content control (any level) whose ``w:showingPlcHdr`` is on."""
    sdt_pr = getattr(node, "sdt_pr", None)
    if sdt_pr is None or getattr(node, "sdt_content", None) is None:
        return False
    return is_on(sdt_property(sdt_pr, "showingPlcHdr", WML_NS))


def in_showing_placeholder(node: Any) -> bool:
    """Whether a run (or a paragraph) is inside a content control showing its placeholder. CR-016 Phase B.

    docx4j-core-ts ``inShowingPlaceholder`` (``tree.mts`` at ``222cc20``; its
    CR-012 section 18): up the parent pointers, a run-, block-, row- or
    cell-level ``w:sdt`` whose ``w:showingPlcHdr`` is on, read as Word reads an
    on/off value (``w:val="0"`` is off; CR-018). Office JS reads a placeholder
    as no text and finds no text there (core-ts check 53 cases 34 and 35, check
    58 cases 11 and 12), so :func:`hidden_runs_of` takes its runs as it takes
    hidden ones.
    """
    current = getattr(node, "parent", None)
    for _ in range(256):
        if current is None:
            return False
        if _shows_placeholder(current):
            return True
        current = getattr(current, "parent", None)
    return False


def placeholder_control_of(run: Any, paragraph: Any) -> Any | None:
    """The run-level content control showing its placeholder that holds `run` inside `paragraph`, or None.

    Up the run's parent pointers to `paragraph` (CR-016 Phase B): a control
    above the paragraph, block-level, is not this paragraph's.
    """
    current = getattr(run, "parent", None)
    for _ in range(256):
        if current is None or current is paragraph:
            return None
        if _shows_placeholder(current):
            return current
        current = getattr(current, "parent", None)
    return None


def hidden_runs_of(resolver: Any, paragraph: Any) -> Callable[[Any], bool]:
    """Which of a paragraph's runs are hidden, as the text model's rule has it (CR-013 Phase A).

    WD REQ-507, ``model/text/extractor.py``'s ``_run``: a run's own
    ``w:vanish``, else its paragraph style's and character style's
    (:meth:`~docx4j_py.model.property_resolver.PropertyResolver.is_vanished_by_styles`).
    Office JS's text leaves such a run out in every view (docx4j-core-ts check
    48b, ``hiddenRunsOf`` at ``9307bdb``). Without a resolver only a run's own
    ``w:vanish`` counts.

    **A run in a content control showing its placeholder** is hidden too, first
    (CR-016 Phase B; core-ts ``hiddenRunsOf`` at ``222cc20``,
    :func:`in_showing_placeholder`): Office JS reads a placeholder as no text,
    so every reader through this rule --- ``Paragraph.text``, the story layout,
    a range, ``search``, a comment's and a change's offsets, the text model's
    ``office_offset`` --- reads it as nothing. The text model's own text keeps
    it (docx4j's semantics). A paragraph inside such a control is found once,
    and a run's own climb stops at the paragraph.
    """
    p_pr = getattr(paragraph, "p_pr", None)
    style = getattr(p_pr, "p_style", None) if p_pr is not None else None
    value = getattr(style, "val", None) if style is not None else None
    p_style = getattr(value, "value", value)
    above: list[bool] = []

    def placeholder(run: Any) -> bool:
        current = getattr(run, "parent", None)
        for _ in range(256):
            if current is None or current is paragraph:
                break
            if _shows_placeholder(current):
                return True
            current = getattr(current, "parent", None)
        # the run's climb reached its paragraph (or its parent pointers stop short of it, as a run
        # read from its paragraph's lists is the paragraph's): the controls above the paragraph
        if not above:
            above.append(in_showing_placeholder(paragraph))
        return above[0]

    def hidden(run: Any) -> bool:
        if placeholder(run):
            return True
        r_pr = getattr(run, "r_pr", None)
        own = getattr(r_pr, "vanish", None) if r_pr is not None else None
        if own is not None:
            return _on(own)
        if resolver is None:
            return False
        r_style = getattr(getattr(r_pr, "r_style", None), "val", None) if r_pr is not None else None
        return resolver.is_vanished_by_styles(p_style, getattr(r_style, "value", r_style))

    return hidden


def _model_character(segment: Segment) -> str:
    """The text model's character for a segment's item, where its Office JS character differs (Q11)."""
    name = element_name(segment.item) if not is_any_element(segment.item) else segment.item.qname
    if name == W_SYM:
        code = None
        for attribute in ("char_value", "char", "char_attribute"):
            code = getattr(segment.item, attribute, None)
            if isinstance(code, str):
                break
        if code is None and is_any_element(segment.item):
            attributes = getattr(segment.item, "attributes", None) or {}
            code = attributes.get(_w("char")) or attributes.get("char")
        if isinstance(code, str) and re.fullmatch(r"[0-9A-Fa-f]{1,6}", code):
            point = int(code, 16)
            if point <= 0x10FFFF and not 0xD800 <= point <= 0xDFFF:
                return chr(point)
        return segment.text
    return _MODEL_CHARACTERS.get(name or "", segment.text)


def office_as_model(segments: list[Segment]) -> str:
    """``Paragraph.text`` with Office JS's characters read as the text model's: Q11's character map.

    CR-013 Q11, decided 2026-10-23: a text model run keeps the
    ``office_offset`` the extractor gives it (where the item's length agrees,
    as core-ts's ``Extractor.mts`` assigns it), and the invariant of CR-011
    section 12.6 is restated as agreement **under the character map**: each
    run with an offset equals this string at its offset, character for
    character, a break's U+000B read as the model's ``\\n``, a ``w:ptab``'s
    ``0`` as ``\\t``, the hyphens U+001E and U+001F as U+2011 and U+00AD and
    a ``w:sym``'s ``(`` as the character its ``w:char`` names. The same length
    as the segments' text, offset for offset.
    """
    out: list[str] = []
    for segment in segments:
        if segment.math or len(segment.text) != 1 or segment.text == "\t":
            out.append(segment.text)
            continue
        mapped = _model_character(segment)
        out.append(mapped if len(mapped) == 1 else segment.text)
    return "".join(out)


def refuse_equation(segments: list[Segment], start: int, end: int, *, tracked: bool | str = False) -> None:
    """Refuse an edit that would cut an equation, or replace one under change tracking (CR-013 Phase A).

    An equation's text is one segment, Word's drawing of its math runs
    (:mod:`.math_text`), so an edit may take it whole but not part of it, nor
    insert into it. Under change tracking (`tracked`) a deletion
    (``"delete"``) takes it whole in Word's form, a ``w:del`` in each math run
    and structure (CR-013 section 13.7, D2; docx4j-core-ts check 57), and a
    replacement (``"replace"``, or True) doesn't take it at all: where Word puts
    the new text beside a deleted equation isn't measured. docx4j-core-ts's
    ``assertPlaceable`` at ``906d5e3``.

    Raises:
        SpanError: ``range.cuts_equation``, an edit or an insertion strictly
            inside the equation's text.
        TrackedChangeError: ``tracking.equation``, a tracked edit taking one.
    """
    for segment in segments:
        if not segment.math:
            continue
        cuts = segment.start < start < segment.end or segment.start < end < segment.end
        takes = start < segment.end and segment.start < end
        if cuts:
            from docx4j_py.model.content.errors import SpanError

            raise SpanError(
                f"{start}:{end} cuts the equation at {segment.start}:{segment.end}: an edit takes an "
                "equation whole, or not at all (CR-013 Phase A)",
                code="range.cuts_equation",
                hint=f"widen the span to {min(start, segment.start)}:{max(end, segment.end)}, or edit beside the equation",
            )
        if tracked and tracked != "delete" and takes:
            from docx4j_py.model.content.errors import TrackedChangeError

            raise TrackedChangeError(
                f"{start}:{end} takes the equation at {segment.start}:{segment.end} under change tracking, "
                "with text in its place: a tracked replacement of an equation is not measured (CR-013 "
                "section 13.7, D2; a tracked deletion takes it whole in Word's form)",
                code="tracking.equation",
                hint="delete the equation first (a tracked deletion takes it), then insert the text; or "
                "replace a Range, whose insertion goes first",
            )


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


def is_block_container(value: Any) -> bool:
    """Whether `value` is one of the block, row or cell containers whose list :func:`block_children_of` reads.

    The classes of :data:`_CONTAINERS` (docx4j ``McSplice``'s block, row and
    cell containers): a body, a header or footer, a cell, a note, a comment, a
    text box's content, a block control's content, block custom XML, a table,
    a row-level control's content, a row and a cell-level control's content.
    Not a ``w:p`` or a ``w:r``, whose lists are run content. CR-013 Phase B's
    ``story_of`` climbs through these.
    """
    cls = value.__class__
    return cls.__module__ == "docx4j_py.wml" and any(
        base.__name__ in _CONTAINERS for base in cls.__mro__ if base.__module__ == "docx4j_py.wml"
    )


def is_branch(value: Any) -> bool:
    """Whether `value` is a typed ``mc:Choice`` or ``mc:Fallback``."""
    cls = value.__class__
    return cls.__module__ == "docx4j_py.mce" and cls.__name__ in _MC_BRANCHES


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


def refuse_tracked_equation(element: Any, operation: str) -> None:
    """Formerly: refuse a tracked deletion of `element` --- a body, a table, a row --- that would take an equation.

    CR-013 Phase A refused it, a tracked change inside an equation being its
    math runs' own; since CR-013 section 13.7 (D2; docx4j-core-ts check 57) a
    tracked deletion writes an equation deleted in Word's form
    (:func:`~.tracking.mark_equation_deleted`), so nothing is refused here and
    no verb calls this. Kept so that a caller of the Phase A name keeps working.
    """


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
    paragraph_element: Any,
    offset: int,
    *,
    prefer: str = "back",
    made: list | None = None,
    hidden: Callable[[Any], bool] | None = None,
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
        hidden: which runs are hidden (:func:`segments_of`'s), so that the
            offset is the paragraph's ``text``'s.

    Returns:
        The offset actually used, which is the one a caller must go on with.

    Raises:
        SpanError: ``range.cuts_equation``, the offset is inside an equation's
            text (:func:`refuse_equation`).
    """
    from docx4j_py.child import ChildList, deep_copy, link_parents
    from docx4j_py.wml import R, el

    segments = segments_of(paragraph_element, hidden=hidden)
    text = "".join(segment.text for segment in segments)
    offset = snap_back(text, offset) if prefer == "back" else snap_forward(text, offset)
    refuse_equation(segments, offset, offset)

    segment = next(
        (s for s in segments if s.editable and s.start < offset < s.end),
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
# writing Office JS's characters back (CR-013 Phase A, Q2)
# ---------------------------------------------------------------------------

#: What a text writer makes of each character that is not written as text
#: (CR-013 Q2, decided 2026-10-23): the element it reads from, so that a text
#: read and written back reads as the items it read. ``\\t`` and ``\\n`` keep
#: 0.1.0's ``w:tab`` and ``w:br``. A ``w:sym``'s ``(`` and a ``w:ptab``'s ``0``
#: are ordinary text, so writing their read text back is lossy, as it is in Word.
#: ``\\n`` and ``\\r`` are this map's only in the one-paragraph writers; every
#: writer through ``Paragraph.splice`` makes a paragraph of them first (CR-013
#: section 11, E1; :data:`PARAGRAPH_BREAKS`).
WRITTEN_AS: dict[str, tuple[str, str | None]] = {
    "\t": ("tab", None),
    "\n": ("br", None),
    "\x0b": ("br", None),
    "\x0c": ("br", "page"),
    "\x0e": ("br", "column"),
    "\r": ("cr", None),
    "\x1e": ("noBreakHyphen", None),
    "\x1f": ("softHyphen", None),
}

_MAPPED = re.compile("[\t\n\x0b\x0c\x0e\r\x1e\x1f]")

#: A paragraph break in a text written through ``Paragraph.splice``: ``\\r\\n``
#: as one, then ``\\r`` and ``\\n`` each, as Word's ``insertText`` makes them
#: (docx4j-core-ts check 55, cases 2 to 4 and 48 to 49, at ``a923aeb``; CR-013
#: section 11, E1).
PARAGRAPH_BREAKS = re.compile("\r\n|\r|\n")

#: What makes a line in a plain-text content control's text: the paragraph breaks
#: and the three ``w:br`` characters, U+000B, U+000C and U+000E (CR-013 section 11,
#: F3; core-ts's ``forWriter`` at ``6285dfa``).
LINE_CHARACTERS = re.compile("[\r\n\x0b\x0c\x0e]")


def for_writer(text: str, lines: str) -> str:
    """`text` as a writer of one kind writes its line feeds and carriage returns (core-ts's ``forWriter``, ``6285dfa``).

    * ``"one-paragraph"``: ``\\r\\n`` and ``\\n`` a ``w:br``, ``\\r`` a ``w:cr``
      (CR-013 section 11, E2: ``insert_paragraph``, ``Paragraph.text = ...``);
    * ``"multi-line"``: ``\\r\\n``, ``\\r`` and ``\\n`` each a ``w:br`` (a multi-line
      ``w:text`` control, section 11's item 2);
    * ``"single-line"``: none of them, nor a line, page or column break (U+000B,
      U+000C, U+000E): a single-line ``w:text`` control, or a bound block control
      Word reads as one (F3; core-ts check 55b case 8, check 56 cases 3, 11, 15).

    A ``w:br`` is written from ``\\n`` or U+000B by :data:`WRITTEN_AS`.
    """
    if lines == "one-paragraph":
        return text.replace("\r\n", "\n")
    if lines == "multi-line":
        return PARAGRAPH_BREAKS.sub("\n", text)
    if lines == "single-line":
        return LINE_CHARACTERS.sub("", text)
    return text


#: The characters a text writer drops: U+0002 (a note's character: no text
#: makes a note), and every other character XML 1.0 excludes that the map does
#: not take --- U+0000 to U+0008, U+000F to U+001D, a lone surrogate, U+FFFE and
#: U+FFFF (CR-013 section 4.1; dropped since section 11, E3).
_UNWRITABLE = re.compile("[\x00-\x08\x0f-\x1d\ud800-\udfff\ufffe\uffff]")


def unwritable(text: str) -> tuple[int, str] | None:
    """The first character a text writer drops, with its offset, or None (:func:`drop_unwritable`)."""
    if not text:
        return None
    found = _UNWRITABLE.search(text)
    return None if found is None else (found.start(), found.group())


def drop_unwritable(text: str | None) -> str:
    """`text` without the characters no element can carry, as Word's writers drop them.

    CR-013 section 11, E3 (decided 2026-10-23, "Drop, as Word does"; docx4j-core-ts
    CR-002 section 60.4): U+0002, a note reference's character in Office JS's
    text, and the other characters XML 1.0 excludes are dropped at the write,
    with no error. Word 365 dropped U+0002, U+0001 and U+0007 through
    ``insertText``, ``insertParagraph``, ``insertComment`` and ``comment.content``
    and refused none (core-ts check 55, cases 10 to 12, 24 to 26 and 38 to 40, at
    ``a923aeb``). It supersedes Phase A's refusal, ``text.not_writable``.
    lxml's refusal at save stays as the guard behind this one, for a value set
    on the tree directly. None is ``""``.
    """
    if not text:
        return ""
    return _UNWRITABLE.sub("", text)


def _written_item(char: str) -> Any:
    from docx4j_py.wml import el

    name, kind = WRITTEN_AS[char]
    if name == "br":
        return el.br() if kind is None else el.br(type_value=kind)
    return getattr(el, name)()


def text_items(text: str) -> list[Any]:
    """The run items a text writer makes of `text`: ``w:t`` pieces and each mapped character's element.

    :data:`WRITTEN_AS` (CR-013 Q2): U+000B and ``\\n`` a ``w:br``, U+000C and
    U+000E a page and a column break, U+000D a ``w:cr``, U+001E and U+001F the
    hyphens, ``\\t`` a ``w:tab``; ``\\r\\n`` one ``w:br`` (CR-013 section 11,
    E2). A text with none of them is one ``w:t``, as before. The caller has
    dropped the characters no element carries (:func:`drop_unwritable`).
    """
    from docx4j_py.wml import el

    if not text:
        return []
    text = text.replace("\r\n", "\n")
    if _MAPPED.search(text) is None:
        return [el.t(text)]
    out: list[Any] = []
    buffer: list[str] = []
    for char in text:
        if char in WRITTEN_AS:
            if buffer:
                out.append(el.t("".join(buffer)))
                buffer = []
            out.append(_written_item(char))
        else:
            buffer.append(char)
    if buffer:
        out.append(el.t("".join(buffer)))
    return out


def text_run(text: str, r_pr: Any = None) -> Any:
    """A ``w:r`` of `text` through the writers' map (:func:`text_items`), with `r_pr` if given."""
    from docx4j_py.child import ChildList
    from docx4j_py.wml import R

    run = R(content=ChildList(text_items(text)))
    if r_pr is not None:
        run.r_pr = r_pr
        r_pr.parent = run
    return run


def explode_mapped(segments: list[Segment], start: int, end: int) -> int:
    """Turn the mapped characters a writer put into ``w:t``\\ s at ``[start, end)`` into their elements.

    A write puts its text into a ``w:t`` first, where the edit primitives place
    and format it as Word does; this then replaces each character of
    :data:`WRITTEN_AS` in that span with the element it reads from, splitting
    the ``w:t`` around it inside the same run. Each element reads as one
    character, so no offset moves. Text outside the span is left alone, a
    document's own raw tab in a ``w:t`` among it. Returns how many were made.
    """
    from docx4j_py.child import link_parents
    from docx4j_py.wml import el

    made = 0
    targets = [
        segment
        for segment in segments
        if segment.editable and segment.end > start and segment.start < end and _MAPPED.search(segment.text)
    ]
    # last first, so that the indices of what is still to do hold
    for segment in reversed(targets):
        value = segment.text
        low = max(0, start - segment.start)
        high = min(len(value), end - segment.start)
        if _MAPPED.search(value, low, high) is None:
            continue
        pieces: list[Any] = []
        buffer = value[:low]
        for char in value[low:high]:
            if char in WRITTEN_AS:
                pieces.append(buffer)
                pieces.append(_written_item(char))
                buffer = ""
                made += 1
            else:
                buffer += char
        pieces.append(buffer + value[high:])
        # the first piece of text stays in the w:t the document had, its attributes kept
        items: list[Any] = []
        head = pieces[0]
        for piece in pieces:
            if isinstance(piece, str):
                if not piece:
                    continue
                if piece is head and not items:
                    set_text(segment.item, piece)
                    items.append(segment.item)
                else:
                    items.append(el.t(piece))
            else:
                items.append(piece)
        position = next((i for i, item in enumerate(segment.owner) if item is segment.item), -1)
        if position < 0:
            continue
        segment.owner[position : position + 1] = items
        owner = getattr(segment.item, "parent", None) or segment.run
        for item in items:
            link_parents(item)
            item.parent = owner
    return made


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
            ``{n,m}`` a repeat, ``^13`` a paragraph's mark, and ``\\`` to
            escape any of them.

    **Wildcards across a mark** (CR-016 Phase C; core-ts ``searchPattern`` at
    ``3ba1e68``, its CR-002 section 32's K7, held to Word 365 by check 59):
    ``^13`` is ``\\r``, Word's wildcard code for a paragraph's mark, and the
    pattern is compiled with ``re.DOTALL``, so that ``?``, ``*`` and a class
    cross a mark where a search reads one
    (:func:`wildcard_matches_across`), as core-ts's ``[\\s\\S]`` does. ``.``
    crosses ``\\r`` in Python without it, but not ``\\n``, so ``DOTALL`` is what
    lets ``?`` and ``*`` match a line feed held in a ``w:t``, which
    :attr:`Paragraph.text` keeps (CR-016 section 12.5 item 5; core-ts parity);
    no wildcard search reads a row's ``\\r\\n``, since no hit crosses a cell.
    A test pins it. ``?`` is a code point, so it takes an emoji whole, as
    Word's does (core-ts's K4; check 58 case 13).
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
            elif char == "^" and text.startswith("13", index + 1):
                # ^13, Word's wildcard code for a paragraph mark (core-ts check 59)
                out.append(r"\r")
                index += 2
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
    flags = 0 if match_case else re.IGNORECASE
    if match_wildcards:
        flags |= re.DOTALL
    return re.compile(source, flags)


#: A code point at or above U+0300: below it every code point is a grapheme
#: cluster of its own, so a text without one is never filtered (core-ts's guard).
_CLUSTERING = re.compile(r"[^\x00-\u02ff]")


def find_all(text: str, pattern: re.Pattern[str]) -> list[tuple[int, int]]:
    """Every ``[start, end)`` match of a pattern, empty matches skipped, and none that cuts a grapheme cluster.

    **The grapheme filter** (CR-016 Phase C, Q8; docx4j-core-ts's
    ``matchesOf`` since ``ebb218e``, its CR-002 section 32 item 1): a match
    whose start or end is not a grapheme boundary (:func:`is_grapheme_boundary`)
    is dropped --- an Arabic letter without the fatha on it, an ``e`` without
    the U+0301 after it. Such a range is no text a person can see, and
    replacing it would split a character. The matches are found first and
    filtered after, so a dropped one moves no other. The boundaries are read
    only for a text holding a code point at or above U+0300, below which every
    code point is a cluster of its own, so Latin text pays nothing: an ASCII
    text is passed by ``str.isascii()``, which CPython answers without a scan,
    and any other by one scan for such a code point (core-ts's guard). Every
    search, ``find`` and ``replace_text`` goes through here.
    """
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
    if out and not text.isascii() and _CLUSTERING.search(text) is not None:
        out = [(a, b) for a, b in out if is_grapheme_boundary(text, a) and is_grapheme_boundary(text, b)]
    return out


# ---------------------------------------------------------------------------
# Office JS's search over a paragraph (CR-013 Phase A, Q1; core-ts officeMatches)
# ---------------------------------------------------------------------------

#: The object replacement character: a picture's place in the text ``search``
#: reads, which ``^g`` matches and a plain query does not match across.
OBJECT_REPLACEMENT = "￼"

#: Word's Find codes a plain query may hold (core-ts check 48: each matched its
#: character in Office JS's text): ``^f`` a footnote reference, ``^e`` an endnote
#: reference, ``^l`` a line break, ``^m`` a page break, ``^t`` a tab, ``^~`` a
#: non-breaking hyphen, ``^-`` an optional hyphen, ``^g`` a picture, ``^^`` a caret.
FIND_CODES: dict[str, str] = {
    "f": "\x02",
    "e": "\x02",
    "l": "\x0b",
    "m": "\x0c",
    "t": "\t",
    "~": "\x1e",
    "-": "\x1f",
    "g": OBJECT_REPLACEMENT,
    "^": "^",
}

_NOTE_KINDS = {
    _w("footnoteReference"): "footnote",
    _w("footnoteRef"): "footnote",
    _w("endnoteReference"): "endnote",
    _w("endnoteRef"): "endnote",
}


def _plain_query(query: str) -> tuple[str, list[str | None]]:
    """A plain query as an expression over :func:`office_matches`' text, and which of its characters are ``^f`` or ``^e``."""
    source: list[str] = []
    tokens: list[str | None] = []
    index = 0
    while index < len(query):
        char = query[index]
        code = query[index + 1] if char == "^" and index + 1 < len(query) else None
        if code is not None and code in FIND_CODES:
            source.append(re.escape(FIND_CODES[code]))
            tokens.append("footnote" if code == "f" else "endnote" if code == "e" else None)
            index += 2
            continue
        if char == "-":
            source.append("[-\\x1e]")  # a non-breaking hyphen is a hyphen to a plain query
        else:
            source.append(re.escape(char))
        tokens.append(None)
        index += 1
    return "".join(source), tokens


def paragraph_mark_pieces(query: str) -> list[str] | None:
    """A plain query split at Word's paragraph mark code, ``^p``: its pieces, one more than the marks; None where it holds none.

    docx4j-core-ts's ``paragraphMarkPieces`` (``search.mts`` at ``87bbb38``;
    its CR-014 phase 3, CR-013 Phase C here): each piece a plain query of its
    own, ``^^`` staying an escaped caret. What a ``Range``'s search across
    paragraphs matches piece by piece (Word's check 52 case 2).
    """
    pieces: list[str] = []
    piece: list[str] = []
    index = 0
    while index < len(query):
        char = query[index]
        following = query[index + 1] if index + 1 < len(query) else None
        if char == "^" and following == "p":
            pieces.append("".join(piece))
            piece = []
            index += 2
            continue
        if char == "^" and following is not None:
            piece.append(char + following)
            index += 2
            continue
        piece.append(char)
        index += 1
    if not pieces:
        return None
    pieces.append("".join(piece))
    return pieces


def refuse_mark_replacement(owner: str, find: str, options: dict[str, Any] | None = None) -> None:
    """Refuse a replacement whose query holds ``^p``, Word's paragraph mark (core-ts's ``refuseMarkReplacement``).

    ``search`` finds such a hit, but what Office JS's replacement of a mark
    leaves isn't measured, and ``insert_text(..., "Replace")`` keeps a last
    mark (core-ts check 52; its CR-014 section 11.3). Wildcards and regular
    expressions read ``^p`` as the two characters, and are not refused.

    Raises:
        SpanError: ``range.replaces_mark``.
    """
    if (options or {}).get("match_wildcards") or paragraph_mark_pieces(find) is None:
        return
    from docx4j_py.model.content.errors import SpanError

    raise SpanError(
        f"{owner}.replace_text refused: the query holds ^p, Word's paragraph mark, and what replacing "
        "one does isn't measured (CR-013 Phase C; docx4j-core-ts CR-014 section 11.3)",
        code="range.replaces_mark",
        hint="search for the hits (search('...^p...')) and edit their ranges: delete() a hit to join "
        "its paragraphs, or insert_text on each",
    )


def refuse_wildcard_mark_replacement(owner: str, hits: list[Any]) -> None:
    """Refuse a wildcard replacement where a hit holds a paragraph mark (core-ts's ``refuseWildcardMarkReplacement``).

    CR-016 Phase C (Q9; core-ts ``3ba1e68``, its CR-002 section 32's K7):
    ``search`` with wildcards finds a hit holding a mark, as Word's does
    (check 59), but what replacing one leaves isn't measured, as for a plain
    ``^p`` query (:func:`refuse_mark_replacement`). `hits` are the search's,
    every one found before anything changes (CR-013 section 4.9 item 3's rule);
    a hit holds a mark where it spans paragraphs or holds its end mark.

    Raises:
        SpanError: ``range.replaces_mark``.
    """
    if not any(hit.spans_paragraphs or hit.holds_end_mark for hit in hits):
        return
    from docx4j_py.model.content.errors import SpanError

    raise SpanError(
        f"{owner}.replace_text refused: a hit holds a paragraph mark, which Word's wildcards match "
        "(core-ts check 59), and what replacing one does isn't measured (CR-016 Phase C; "
        "docx4j-core-ts CR-002 section 32)",
        code="range.replaces_mark",
        hint="search for the hits (search(..., match_wildcards=True)) and edit their ranges",
    )


def office_matches(
    segments: list[Segment],
    query: str,
    *,
    match_case: bool = False,
    match_whole_word: bool = False,
    match_wildcards: bool = False,
    start: int = 0,
    end: int | None = None,
) -> list[tuple[int, int]]:
    """Every ``[start, end)`` match of a query in a paragraph's text by Office JS's rules.

    CR-013 Phase A (Q1), the port of docx4j-core-ts's ``officeMatches`` at
    ``9307bdb``, held to Word 365 by checks 48 and 48b (Word's default view).
    The text is ``Paragraph.text`` --- hidden text left out, so a query matches
    across it and never in it --- where a picture is a place of no text that a
    plain query does not match across and ``^g`` matches (U+FFFC), and an
    optional hyphen is passed over unless the query names it (``^-``). A plain
    query takes Word's Find codes (:data:`FIND_CODES`; ``^f`` and ``^e`` each
    match only their own kind of reference), and its ``-`` matches a
    non-breaking hyphen too. Wildcards read the same text, the codes not
    translated. Only text in ``[start, end)`` is read: a range's search.

    Args:
        segments: the paragraph's, with its objects and hidden items
            (``segments_of(objects=True)``).
        query: what to look for.
        match_case, match_whole_word, match_wildcards: Office JS's
            ``SearchOptions``, as :func:`search_pattern`.
        start, end: the span searched; None for the paragraph's end.

    Returns:
        The matches, as offsets in ``Paragraph.text``. An empty match is skipped.
    """
    plain = not match_wildcards
    keep_soft_hyphens = not plain or "^-" in query.replace("^^", "")
    tokens: list[str | None] = []
    if plain:
        source, tokens = _plain_query(query)
        if match_whole_word:
            source = rf"(?<!\w){source}(?!\w)"
        pattern = re.compile(source, 0 if match_case else re.IGNORECASE)
    else:
        pattern = search_pattern(
            query, match_case=match_case, match_whole_word=match_whole_word, match_wildcards=True
        )
    limit = float("inf") if end is None else end
    # the usual paragraph: nothing but text, no optional hyphen to pass over, no
    # reference a code must tell apart: the text searched is Paragraph.text's own
    if all(not s.object for s in segments) and all(kind is None for kind in tokens):
        whole = "".join(s.text for s in segments)
        low = max(0, start)
        high = len(whole) if end is None else min(len(whole), end)
        span = whole[low:high]
        if keep_soft_hyphens or "\x1f" not in span:
            return [(low + a, low + b) for a, b in find_all(span, pattern)]
    text: list[str] = []
    starts: list[int] = []
    ends: list[int] = []
    notes: list[str | None] = []
    for segment in segments:
        if segment.hidden:
            continue
        if segment.object:
            if start < segment.start < limit:
                text.append(OBJECT_REPLACEMENT)
                starts.append(segment.start)
                ends.append(segment.start)
                notes.append(None)
            continue
        kind = _NOTE_KINDS.get(element_name(segment.item) or "") if not segment.math else None
        for offset, char in enumerate(segment.text):
            at = segment.start + offset
            if at < start or at >= limit:
                continue
            if char == "\x1f" and not keep_soft_hyphens:
                continue
            text.append(char)
            starts.append(at)
            ends.append(at + 1)
            notes.append(kind if char == "\x02" else None)
    joined = "".join(text)
    out: list[tuple[int, int]] = []
    for a, b in find_all(joined, pattern):
        # ^f and ^e both read U+0002: each matches only its own kind of reference
        if any(
            kind is not None and (a + i >= len(notes) or notes[a + i] != kind)
            for i, kind in enumerate(tokens)
        ):
            continue
        out.append((starts[a], ends[b - 1]))
    return out


# ---------------------------------------------------------------------------
# Word's wildcards across a paragraph's mark (CR-016 Phase C; core-ts K7)
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class MarkedSegments:
    """A paragraph's part in a wildcard search across its marks (core-ts's ``MarkedSegments`` at ``3ba1e68``)."""

    segments: list[Segment]
    """The paragraph's, with its objects and hidden items (``segments_of(objects=True)``)."""
    start: int
    """Where the text read begins, in ``Paragraph.text``."""
    end: int
    """Where it ends."""
    length: int
    """The length of ``Paragraph.text``: where the mark stands."""
    mark: bool
    """The paragraph's mark follows its text, as a hit may hold it."""


@dataclasses.dataclass(frozen=True, slots=True)
class MarkedMatch:
    """A hit of a wildcard search across marks (core-ts's ``MarkedMatch``): from `start` in piece `first` to `end` in piece `last`."""

    first: int
    start: int
    last: int
    end: int
    mark: bool
    """The hit ends with piece `last`'s mark, which it holds."""


@dataclasses.dataclass(slots=True)
class _Placed:
    """Where a piece's characters stand in the text a wildcard search reads."""

    base: int
    piece: int
    low: int
    count: int
    length: int
    starts: list[int] | None
    ends: list[int] | None


def wildcard_matches_across(
    pieces: list[MarkedSegments],
    query: str,
    *,
    match_case: bool = False,
    match_whole_word: bool = False,
    match_wildcards: bool = True,
) -> list[MarkedMatch]:
    """A wildcard query's hits over consecutive paragraphs, each followed by its mark where it says so.

    CR-016 Phase C, the port of docx4j-core-ts's ``wildcardMatchesAcross``
    (``search.mts`` at ``3ba1e68``, its CR-002 section 32's K7), held to Word
    365 by core-ts check 59: Word's wildcards match a paragraph's mark exactly
    where ``^p`` does, so ``?``, ``*`` and a class cross one and ``^13`` is
    one. The text read in each paragraph is :func:`office_matches`' for
    wildcards (hidden text left out, a picture U+FFFC, only ``[start, end)``),
    and each mark that follows is ``\\r``. The caller gives only paragraphs a
    hit may span (one cell's, or none's) and only the marks it may hold (a
    paragraph's own, never a cell's or a row's end, never a deleted one).
    Every hit is filtered by :func:`find_all`'s grapheme rule over that text.

    Returns:
        The hits, each from offset `start` in piece `first` to `end` in piece
        `last`, holding that piece's mark where `mark`.
    """
    from bisect import bisect_right

    pattern = search_pattern(query, match_case=match_case, match_whole_word=match_whole_word, match_wildcards=True)
    chunks: list[str] = []
    placed: list[_Placed] = []
    position = 0
    for k, piece in enumerate(pieces):
        starts: list[int] | None = None
        ends: list[int] | None = None
        if all(not segment.object for segment in piece.segments):
            # the usual paragraph: nothing but text, each offset itself
            whole = "".join(segment.text for segment in piece.segments)
            low = max(0, piece.start)
            chars = whole[low : max(low, min(len(whole), piece.end))]
        else:
            low = 0
            out: list[str] = []
            starts, ends = [], []
            for segment in piece.segments:
                if segment.hidden:
                    continue
                if segment.object:
                    if piece.start < segment.start < piece.end:
                        out.append(OBJECT_REPLACEMENT)
                        starts.append(segment.start)
                        ends.append(segment.start)
                    continue
                for offset, char in enumerate(segment.text):
                    at = segment.start + offset
                    if piece.start <= at < piece.end:
                        out.append(char)
                        starts.append(at)
                        ends.append(at + 1)
            chars = "".join(out)
        placed.append(_Placed(position, k, low, len(chars), piece.length, starts, ends))
        chunks.append(chars)
        position += len(chars)
        if piece.mark:
            chunks.append("\r")
            position += 1
    text = "".join(chunks)
    bases = [entry.base for entry in placed]

    def place(index: int) -> tuple[int, int, int, bool]:
        """The piece, the start and end offsets, and whether it is the mark, of the character at `index`."""
        entry = placed[bisect_right(bases, index) - 1]
        j = index - entry.base
        if j == entry.count:
            return entry.piece, entry.length, entry.length, True
        if entry.starts is None or entry.ends is None:
            return entry.piece, entry.low + j, entry.low + j + 1, False
        return entry.piece, entry.starts[j], entry.ends[j], False

    hits: list[MarkedMatch] = []
    for a, b in find_all(text, pattern):
        first, start, _end, _mark = place(a)
        last, _start, end, mark = place(b - 1)
        hits.append(MarkedMatch(first, start, last, end, mark))
    return hits
