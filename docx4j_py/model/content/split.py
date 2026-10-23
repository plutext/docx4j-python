"""A paragraph split in two, as Word splits one: the port of docx4j-core-ts's ``splitParagraph``.

CR-013 section 11, E1 (decided 2026-10-23, "let's fix core-ts's E1 and E3 now"):
every writer through :meth:`~docx4j_py.model.content.paragraph.Paragraph.splice`
makes a new paragraph of each ``\\n`` and each ``\\r`` in its text, ``\\r\\n`` one,
as Word's ``insertText`` does (core-ts check 55, cases 2 to 4 and 48 to 49, at
``a923aeb``). The split is core-ts's ``splitParagraph`` (``spanEdits.mts`` at
``1e61bb1``, its CR-014 section 10.1, measured by its check 52's cases 57 and 63):

* the text before the split becomes a **new paragraph in front**, in a copy of
  the paragraph's properties less its mark's run properties, its section break
  and any ``w:pPrChange``, with a mark of its own in the formatting of the run
  before the split;
* the original ``w:p`` keeps the rest, **its own mark, its section break and its
  identity** (its ``w14:paraId``), as Word keeps them; so a
  :class:`~docx4j_py.model.content.paragraph.Paragraph` view goes on viewing the
  last of the paragraphs a write makes;
* a hyperlink, a revision (``w:ins``, ``w:del``, ``w:moveFrom``, ``w:moveTo``), a
  smart tag, a ``w:customXml``, a ``w:dir`` or a ``w:bdo`` the split falls inside
  is cut into two of the same, a revision's second half under a new ``w:id``;
* a run-level content control (``range.splits_control``), a field --- inside a
  ``w:fldSimple``, or between a ``w:fldChar`` begin and its end in the paragraph
  (``range.splits_field``) --- an equation (``range.cuts_equation``) or a kept
  ``mc:AlternateContent`` branch (``address.read_only``) there is refused,
  before anything changes, since each half would be one of it;
* under change tracking the new mark is marked inserted (``w:pPr/w:rPr/w:ins``),
  as Word marks the first half's (check 55 case 48).

A field that begins in an earlier paragraph and ends in a later one is not seen
from a paragraph between them: such a paragraph holds no ``w:fldChar``, and a
split there leaves a field spanning one more paragraph, which is well formed.
"""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, deep_copy, deep_copy_as, link_parents
from docx4j_py.model.content.errors import AddressError, SpanError
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = [
    "SPLITTABLE",
    "field_depths",
    "holders_above",
    "refuse_cut",
    "split_paragraph",
]


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


_W_P = _w("p")
_W_R = _w("r")
_W_SDT = _w("sdt")
_W_SDT_CONTENT = _w("sdtContent")
_W_FLD_SIMPLE = _w("fldSimple")
_W_FLD_CHAR = _w("fldChar")

#: The run holders a split may cut in two: each half is a whole one of the same
#: (core-ts's ``SPLITTABLE`` at ``1e61bb1``).
SPLITTABLE = frozenset(
    {
        _w("hyperlink"),
        _w("ins"),
        _w("del"),
        _w("moveFrom"),
        _w("moveTo"),
        _w("smartTag"),
        _w("customXml"),
        _w("dir"),
        _w("bdo"),
    }
)

_REVISIONS = frozenset({_w("ins"), _w("del"), _w("moveFrom"), _w("moveTo")})


def holders_above(item: Any, paragraph_element: Any) -> list[Any]:
    """The elements between `item` and the ``w:p``, innermost first, `item` itself left out."""
    out: list[Any] = []
    current = getattr(item, "parent", None)
    for _guard in range(64):
        if current is None or current is paragraph_element:
            return out
        out.append(current)
        current = getattr(current, "parent", None)
    return out  # pragma: no cover - 64 holders deep


def refuse_cut(holders: list[Any], operation: str = "this write") -> None:
    """Refuse a split that would cut one of `holders` and cannot (core-ts's rule): all but :data:`SPLITTABLE` and runs.

    Raises:
        SpanError: ``range.splits_control``, a run-level content control;
            ``range.splits_field``, a ``w:fldSimple``; ``range.crosses_holder``,
            any other holder.
        AddressError: ``address.read_only``, a kept ``mc:AlternateContent``.
    """
    from docx4j_py.child import is_alternate_content

    for holder in holders:
        name = element_name(holder)
        if name in SPLITTABLE or name in (_W_R, _w("t")):
            continue
        if name in (_W_SDT, _W_SDT_CONTENT):
            raise SpanError(
                f"{operation} refused: a paragraph break there would cut a run-level content control in two "
                "(CR-013 section 11, E1; docx4j-core-ts F4)",
                code="range.splits_control",
                hint="write the break outside the control, or write a line break (U+000B) instead",
            )
        if name == _W_FLD_SIMPLE:
            raise SpanError(
                f"{operation} refused: a paragraph break there would cut a field (w:fldSimple) in two "
                "(CR-013 section 11, E1)",
                code="range.splits_field",
                hint="write the break before or after the field, or write a line break (U+000B) instead",
            )
        if is_alternate_content(holder):
            raise AddressError(
                f"{operation} refused: a paragraph break there would cut a kept mc:AlternateContent "
                "(CR-011 Phase F; CR-013 section 11, E1)",
                code="address.read_only",
                hint="write the break beside the element, where it goes into no branch",
            )
        raise SpanError(  # pragma: no cover - every run holder of WordprocessingML is named above
            f"{operation} refused: a paragraph break there would cut a {name} in two (CR-013 section 11, E1)",
            code="range.crosses_holder",
            hint="write the break outside it",
        )


def _refuse_field(operation: str) -> None:
    raise SpanError(
        f"{operation} refused: a paragraph break there would fall inside a field, between its "
        "w:fldChar begin and its end (CR-013 section 11, E1)",
        code="range.splits_field",
        hint="write the break before or after the field, or write a line break (U+000B) instead",
    )


@dataclasses.dataclass(slots=True)
class FieldDepths:
    """How deep in a complex field each element of a paragraph stands, on entry and on exit.

    :attr:`floor` is the lowest depth the walk reached: below zero where the
    paragraph ends a field an earlier paragraph began, so that ``depth -
    floor`` is how many fields hold a place.
    """

    entry: dict[int, int]
    exit: dict[int, int]
    floor: int
    closing: int

    def inside(self, depth: int) -> bool:
        """Whether a place at this raw depth is inside a field."""
        return depth - self.floor > 0


def field_depths(paragraph_element: Any) -> FieldDepths:
    """Every element's field depth in a ``w:p``, by ``id()``: what a ``w:fldChar`` begin opens and its end closes."""
    entry: dict[int, int] = {}
    exit_: dict[int, int] = {}
    state = {"depth": 0, "floor": 0}

    def walk(items: list) -> None:
        for item in items:
            entry[id(item)] = state["depth"]
            if element_name(item) == _W_R:
                for child in getattr(item, "content", None) or ():
                    entry[id(child)] = state["depth"]
                    if element_name(child) == _W_FLD_CHAR:
                        kind = getattr(getattr(child, "fld_char_type", None), "value", None)
                        if kind == "begin":
                            state["depth"] += 1
                        elif kind == "end":
                            state["depth"] -= 1
                            state["floor"] = min(state["floor"], state["depth"])
                    exit_[id(child)] = state["depth"]
            else:
                nested = run_items_of(item)
                if nested is not None:
                    walk(nested)
            exit_[id(item)] = state["depth"]

    walk(getattr(paragraph_element, "content", None) or [])
    return FieldDepths(entry, exit_, state["floor"], state["depth"])


def refuse_place(paragraph_element: Any, anchor: Any, how: str, operation: str = "this write") -> None:
    """Refuse a paragraph break in text written at a place: into `anchor` (``"into"``), or ``"after"`` or ``"before"`` it.

    The place is where :meth:`~docx4j_py.model.content.paragraph.Paragraph.splice`
    puts a text (its landing); `anchor` None is the paragraph's own list, which
    nothing holds. A text that lands inside a run-level content control or a
    field is refused whatever its breaks are, since the text is written into it
    (core-ts's F4 for a control).

    Raises:
        SpanError: ``range.splits_control`` or ``range.splits_field``.
    """
    if anchor is None:
        return
    refuse_cut(holders_above(anchor, paragraph_element), operation)
    depths = field_depths(paragraph_element)
    depth = depths.exit.get(id(anchor), 0) if how == "after" else depths.entry.get(id(anchor), 0)
    if depths.inside(depth):
        _refuse_field(operation)


@dataclasses.dataclass(slots=True)
class _Place:
    """Where a split falls in the tree: the list, and the index its items go to the tail from."""

    items: list
    index: int
    run_cut: Any = None
    """The run the split falls inside, which is cut first; None at a run boundary."""
    depth: int = 0
    """The raw field depth at the split (:class:`FieldDepths`)."""
    segment: Any = None
    """The text segment the split stands beside, for the run split."""
    run_before: Any = None
    """The ``w:rPr`` of the run holding the character before the split, for the new mark."""


def _site_of(segment: Any) -> tuple[Any, list, int]:
    if segment.math:
        return segment.run, segment.run_owner, segment.run_index
    return segment.site()


def _place(paragraph: Paragraph, offset: int, keep: str, depths: FieldDepths) -> _Place:
    """Where a split at `offset` falls, read from the tree as it stands: nothing changes.

    `keep` says which side the places of no text at `offset` go to: ``"head"``
    (core-ts's rule: the tail begins at the character at `offset`) or
    ``"tail"`` (the head ends with the character before it).
    """
    content = paragraph.element.content
    segments = paragraph.segments()
    before = next((s for s in reversed(segments) if s.start < offset and s.text), None)
    after = next((s for s in segments if s.end > offset and s.text), None)
    run_before = getattr(before.run, "r_pr", None) if before is not None and not before.math else None
    if after is not None and after.start < offset:
        # strictly inside an item: its run is cut there
        return _Place(after.run_owner, after.run_index + 1, after.run, depths.entry.get(id(after.item), 0), after, run_before)
    if keep == "head":
        if after is None:
            return _Place(content, len(content), depth=depths.closing, run_before=run_before)
        if after.branch is None and not after.math and after.index > 0:
            return _Place(after.run_owner, after.run_index + 1, after.run, depths.entry.get(id(after.item), 0), after, run_before)
        item, owner, _index = _site_of(after)
        return _Place(owner, _index_of(owner, item), depth=depths.entry.get(id(item), 0), segment=after, run_before=run_before)
    if before is None:
        return _Place(content, 0, run_before=None)
    if before.branch is None and not before.math and before.index < len(before.owner) - 1:
        return _Place(before.run_owner, before.run_index + 1, before.run, depths.exit.get(id(before.item), 0), before, run_before)
    item, owner, _index = _site_of(before)
    return _Place(owner, _index_of(owner, item) + 1, depth=depths.exit.get(id(item), 0), segment=before, run_before=run_before)


def _holder_of(items: list, paragraph_element: Any) -> Any:
    """The element whose list `items` is."""
    if items is getattr(paragraph_element, "content", None):
        return paragraph_element
    owner = getattr(items, "owner", None)
    if owner is not None:
        return owner
    sample = next((item for item in items if getattr(item, "parent", None) is not None), None)
    return getattr(sample, "parent", None) if sample is not None else None


def _climb(paragraph_element: Any, place: _Place, cut: Any = None) -> tuple[list[Any], int, int]:
    """Walk from the split up to the ``w:p``: the holders it cuts, and its index in the paragraph's list.

    Each holder holding items on both sides of the split is cut; with `cut`
    given, ``cut(holder, items, index, outer_items, at, outer)`` does it. A run
    split first (:attr:`_Place.run_cut`) counts as one more item in its list.
    The same walk answers the check and does the split, so the two cannot
    disagree. Returns ``(holders cut, index, extra)``: the paragraph's items
    from `index` on go to the tail, `extra` being 1 where the list will have
    gained an item there.
    """
    items = place.items
    index = place.index
    extra = 1 if place.run_cut is not None else 0
    holder = _holder_of(items, paragraph_element)
    holders: list[Any] = []
    for _guard in range(64):
        if holder is None or holder is paragraph_element:
            break
        if element_name(holder) == _W_SDT_CONTENT:
            holder = getattr(holder, "parent", None)
            continue
        outer = getattr(holder, "parent", None)
        outer_items = _items_holding(outer, paragraph_element) if outer is not None else None
        if outer_items is None:
            break
        at = _index_of(outer_items, holder)
        if 0 < index < len(items) + extra:
            holders.append(holder)
            if cut is not None:
                cut(holder, items, index, outer_items, at, outer)
            index, extra = at + 1, 1
        elif index == 0:
            index, extra = at, 0
        else:
            index, extra = at + 1, 0
        items = outer_items
        holder = outer
    return holders, index, extra


def refuse_split(paragraph: Paragraph, offset: int, keep: str = "head", operation: str = "this write") -> None:
    """Refuse a split at `offset` that would cut what can't be cut, before anything changes.

    The split's own walk, dry (:func:`_climb`): a run-level content control or
    a ``w:fldSimple`` it would cut, an equation or a kept branch's text it
    falls inside, or a place between a ``w:fldChar`` begin and its end in the
    paragraph with something of the paragraph on either side.

    Raises:
        SpanError: ``range.splits_control``, ``range.splits_field`` or
            ``range.cuts_equation``.
        AddressError: ``address.read_only``, the split is strictly inside text
            read from a kept ``mc:AlternateContent`` branch.
    """
    from docx4j_py.model.content.text_model import refuse_equation

    segments = paragraph.segments()
    refuse_equation(segments, offset, offset)
    inside = next((s for s in segments if s.start < offset < s.end), None)
    if inside is not None and inside.branch is not None:
        raise AddressError(
            f"{operation} refused: a paragraph break there would cut text read from a kept "
            "mc:AlternateContent branch (CR-011 Phase F; CR-013 section 11, E1)",
            code="address.read_only",
            hint="write the break beside the element, where it goes into no branch",
        )
    before = next((s for s in reversed(segments) if s.start < offset and s.text), None)
    after = next((s for s in segments if s.end > offset and s.text), None)
    if (
        before is not None
        and after is not None
        and (before.branch is not None or after.branch is not None)
        and before.site()[0] is after.site()[0]
    ):
        # the run holding a kept branch holds text on the other side of the split too
        raise AddressError(
            f"{operation} refused: a paragraph break there would cut the run holding a kept "
            "mc:AlternateContent (CR-011 Phase F; CR-013 section 11, E1)",
            code="address.read_only",
            hint="write the break beside the run, where it goes into no branch",
        )
    depths = field_depths(paragraph.element)
    place = _place(paragraph, offset, keep, depths)
    holders, index, extra = _climb(paragraph.element, place)
    refuse_cut(holders, operation)
    content = paragraph.element.content
    edge = index == 0 or index >= len(content) + extra
    if not edge and depths.inside(place.depth):
        _refuse_field(operation)


def _mark_r_pr(r_pr: Any) -> Any:
    """A mark's run properties made from a run's: ``w:rPrChange`` left out (core-ts's ``copyRPr``)."""
    from docx4j_py.wml import ParaRPr

    if r_pr is None:
        return None
    copy = deep_copy_as(r_pr, ParaRPr)
    copy.r_pr_change = None
    return copy


def _properties_for_head(element: Any, r_pr: Any) -> Any:
    """A copy of the paragraph's properties for the new half: no mark, section break or change; `r_pr`'s mark.

    core-ts's ``propertiesOf`` at ``1e61bb1``.
    """
    from docx4j_py.wml import el

    p_pr = deep_copy(element.p_pr) if getattr(element, "p_pr", None) is not None else None
    if p_pr is not None:
        p_pr.r_pr = None
        p_pr.sect_pr = None
        p_pr.p_pr_change = None
    mark = _mark_r_pr(r_pr)
    if mark is None:
        return p_pr
    out = p_pr if p_pr is not None else el.pPr()
    out.r_pr = mark
    link_parents(out)
    return out


def _holder_like(holder: Any, items: list, ids: Any) -> Any:
    """A holder of the same kind and attributes as `holder`, holding `items`; a revision's under a new ``w:id``."""
    values = {
        field.name: deep_copy(getattr(holder, field.name))
        for field in dataclasses.fields(holder)
        if field.name != "content"
    }
    clone = type(holder)(**values, content=ChildList())
    clone.content.extend(items)
    if element_name(holder) in _REVISIONS and ids is not None:
        clone.id = ids.next_id()
    link_parents(clone)
    return clone


def _index_of(items: list, element: Any) -> int:
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def _items_holding(holder: Any, paragraph_element: Any) -> list | None:
    """The list `holder` keeps its children in: a run holder's, or the ``w:p``'s."""
    if holder is paragraph_element:
        return getattr(paragraph_element, "content", None)
    return run_items_of(holder)


def split_paragraph(
    paragraph: Paragraph, offset: int, *, keep: str = "head", operation: str = "this write"
) -> Paragraph:
    """Split `paragraph` at `offset` of its text; the text before it becomes a new paragraph in front, which is returned.

    The module's rules (core-ts's ``splitParagraph`` at ``1e61bb1``). The
    original ``w:p`` keeps the text from `offset` on, its mark, its section
    break and its ``w14:paraId``; the new one, before it in the same list, takes
    a copy of the properties with a mark in the formatting of the run before
    the split, and a ``w14:paraId`` of its own where the package assigns them.
    `offset` may be 0 (an empty paragraph in front) or the text's length (the
    original left with its mark alone). `keep` is :func:`_place`'s. Tracked,
    the new mark is marked inserted.

    Raises:
        SpanError: ``range.splits_control``, ``range.splits_field``,
            ``range.cuts_equation``; nothing has changed.
        AddressError: ``address.read_only``; nothing has changed.
    """
    from docx4j_py.model.content.addresses import assign_para_id
    from docx4j_py.model.content.hyperlinks import _id_source
    from docx4j_py.model.content.reports import current_recorder, moved_by_insert
    from docx4j_py.wml import P

    refuse_split(paragraph, offset, keep, operation)
    element = paragraph.element
    body = paragraph.parent_body
    container = paragraph.container
    ids = _id_source(paragraph)

    # the first change: a run the split falls inside is cut there
    place = _place(paragraph, offset, keep, field_depths(element))
    if place.run_cut is not None:
        paragraph.split_at(offset)
        segment = place.segment
        segments = paragraph.segments()
        if keep == "head" or segment.end > offset:
            at = next(s for s in segments if s.start >= offset and s.text)
            if at.index > 0:
                paragraph._split_run_before(at)
                at = next(s for s in paragraph.segments() if s.start >= offset and s.text)
            place = _Place(at.run_owner, _index_of(at.run_owner, at.run), run_before=place.run_before)
        else:
            before = next(s for s in reversed(segments) if s.start < offset and s.text)
            paragraph._split_run_after(before)
            place = _Place(before.run_owner, _index_of(before.run_owner, before.run) + 1, run_before=place.run_before)

    def cut(holder: Any, items: list, index: int, outer_items: list, at: int, outer: Any) -> None:
        refuse_cut([holder], operation)  # refuse_split has refused it already
        tail = list(items[index:])
        del items[index:]
        clone = _holder_like(holder, tail, ids)
        outer_items.insert(at + 1, clone)
        clone.parent = outer

    _holders, index, _extra = _climb(element, place, cut)
    run_before = place.run_before
    head_items = list(element.content[:index])
    del element.content[:index]
    head = P(content=ChildList(head_items))
    p_pr = _properties_for_head(element, run_before)
    if p_pr is not None:
        head.p_pr = p_pr
    link_parents(head)
    position = _index_of(container, element)
    container.insert(position, head)
    head.parent = getattr(element, "parent", None)

    change = current_recorder(getattr(body, "package", None))
    change.created(assign_para_id(body, head))
    if getattr(change, "active", False):
        change.shifted(moved_by_insert(body, container, position, 1))
    tracker = paragraph.change_tracker
    if tracker is not None:
        tracker.mark_paragraph_inserted(head)
    from docx4j_py.model.content.paragraph import Paragraph as ParagraphView

    view = ParagraphView(head, container, body)
    change.touched(view)
    return view
