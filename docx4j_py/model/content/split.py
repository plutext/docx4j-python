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
* **a field's result is split** (CR-013 section 13.7, D1; docx4j-core-ts check 57
  cases 1 to 4, its ``BREAKABLE`` and ``unfoldSimpleFieldsAround`` at
  ``906d5e3``): a complex field's begin, instruction and separator stay in the
  first paragraph and its end goes with the rest of its result to the next, so
  the field spans the two; a ``w:fldSimple`` the split falls in, or holds the
  character at the split, is first written as the complex field Word writes for
  it (:func:`unfold_simple_field`), the innermost first. A split inside a
  field's instruction, between its ``w:fldChar`` begin and its separator, is
  refused (``range.splits_field``): no text stands there, so only a place of no
  text could ask for one;
* a run-level content control (``range.splits_control``), an equation
  (``range.cuts_equation``) or a kept ``mc:AlternateContent`` branch
  (``address.read_only``) there is refused, before anything changes, since each
  half would be one of it;
* under change tracking the new mark is marked inserted (``w:pPr/w:rPr/w:ins``),
  as Word marks the first half's (check 55 case 48).

A field that begins in an earlier paragraph and ends in a later one is not seen
from a paragraph between them: such a paragraph holds no ``w:fldChar``, and a
split there leaves a field spanning one more paragraph, which is well formed (and
is what Word itself writes for a break in a result, check 57).
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
    "BREAKABLE",
    "SPLITTABLE",
    "field_depths",
    "holders_above",
    "refuse_cut",
    "split_paragraph",
    "unfold_simple_field",
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

#: The run holders a paragraph break may be made inside: those it cuts in two, and a
#: ``w:fldSimple``, made a complex field first (core-ts's ``BREAKABLE`` at ``906d5e3``;
#: CR-013 section 13.7, D1).
BREAKABLE = SPLITTABLE | {_W_FLD_SIMPLE}


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

    A ``w:fldSimple`` is not refused: a break there makes it a complex field first
    (CR-013 section 13.7, D1).

    Raises:
        SpanError: ``range.splits_control``, a run-level content control;
            ``range.crosses_holder``, any other holder.
        AddressError: ``address.read_only``, a kept ``mc:AlternateContent``.
    """
    from docx4j_py.child import is_alternate_content

    for holder in holders:
        name = element_name(holder)
        if name in BREAKABLE or name in (_W_R, _w("t")):
            continue
        if name in (_W_SDT, _W_SDT_CONTENT):
            raise SpanError(
                f"{operation} refused: a paragraph break there would cut a run-level content control in two "
                "(CR-013 section 11, E1; docx4j-core-ts F4)",
                code="range.splits_control",
                hint="write the break outside the control, or write a line break (U+000B) instead",
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
        f"{operation} refused: a paragraph break there would fall inside a field's instruction, between "
        "its w:fldChar begin and its separator (CR-013 section 13.7, D1; a field's result is split)",
        code="range.splits_field",
        hint="write the break in the field's result, before or after the field, or write a line break "
        "(U+000B) instead",
    )


@dataclasses.dataclass(slots=True)
class FieldDepths:
    """Where each element of a paragraph stands in its complex fields, on entry and on exit, by ``id()``.

    A place is **in a field's instruction** where a field this paragraph begins
    has had its ``w:fldChar`` begin and not yet its separator (or its end): the
    values of :attr:`entry` and :attr:`exit` are those flags. A field's result,
    after its separator, is split (CR-013 section 13.7, D1); a field an earlier
    paragraph began is read as in its result, its instruction holding no text
    of this paragraph's.
    """

    entry: dict[int, bool]
    exit: dict[int, bool]
    closing: bool

    def inside(self, in_code: bool) -> bool:
        """Whether a place with this flag is inside a field's instruction."""
        return bool(in_code)


def field_depths(paragraph_element: Any) -> FieldDepths:
    """Every element's place in a ``w:p``'s complex fields, by ``id()``: in an instruction, or not (:class:`FieldDepths`)."""
    entry: dict[int, bool] = {}
    exit_: dict[int, bool] = {}
    # one entry per field this paragraph begins and hasn't ended: True until its separator
    stack: list[bool] = []

    def in_code() -> bool:
        return any(stack)

    def walk(items: list) -> None:
        for item in items:
            entry[id(item)] = in_code()
            if element_name(item) == _W_R:
                for child in getattr(item, "content", None) or ():
                    entry[id(child)] = in_code()
                    if element_name(child) == _W_FLD_CHAR:
                        kind = getattr(getattr(child, "fld_char_type", None), "value", None)
                        if kind == "begin":
                            stack.append(True)
                        elif kind == "separate" and stack:
                            stack[-1] = False
                        elif kind == "end" and stack:
                            stack.pop()
                    exit_[id(child)] = in_code()
            else:
                nested = run_items_of(item)
                if nested is not None:
                    walk(nested)
            exit_[id(item)] = in_code()

    walk(getattr(paragraph_element, "content", None) or [])
    return FieldDepths(entry, exit_, in_code())


def unfold_simple_field(field: Any) -> None:
    """A ``w:fldSimple`` written in its place as the complex field Word writes for it (docx4j-core-ts's ``unfoldSimpleField``).

    Word's form in check 57 (cases 2 and 4; CR-013 section 13.7, D1): a run
    holding the begin ``w:fldChar``, the field's ``w:fldLock``, ``w:dirty`` and
    ``w:fldData`` on it; a run holding its instruction as a ``w:instrText``; a
    run holding the separator; the field's content; a run holding the end. The
    field characters' runs take a copy of the result's first run's properties,
    core-ts's choice: Word's result run had none, and one with some isn't
    measured.
    """
    from docx4j_py.model.content.tracking import copy_r_pr
    from docx4j_py.wml import R, el

    parent = getattr(field, "parent", None)
    items = run_items_of(parent) if parent is not None else None
    index = _index_of(items, field) if items is not None else -1
    if items is None or index < 0:
        raise SpanError(
            "the w:fldSimple could not be found in the list that holds it",
            code="range.stale",
            hint="take the range again",
        )
    content = list(field.content or ())
    first_run = next((item for item in content if isinstance(item, R)), None)
    r_pr = getattr(first_run, "r_pr", None)

    def run_for(item: Any) -> Any:
        run = R(content=ChildList([item]))
        copy = copy_r_pr(r_pr)
        if copy is not None:
            run.r_pr = copy
        link_parents(run)
        return run

    begin = el.fldChar(fld_char_type="begin", fld_lock=field.fld_lock, dirty=field.dirty)
    if field.fld_data is not None:
        begin.fld_data_or_ff_data_or_numbering_change = deep_copy(field.fld_data)
    instr = field.instr or ""
    instruction = el.instrText(value=instr, space="preserve" if instr != instr.strip() else None)
    del field.content[:]
    pieces = [
        run_for(begin),
        run_for(instruction),
        run_for(el.fldChar(fld_char_type="separate")),
        *content,
        run_for(el.fldChar(fld_char_type="end")),
    ]
    items[index : index + 1] = pieces
    for piece in pieces:
        link_parents(piece)
        piece.parent = parent


def _unfold_simple_fields(paragraph: Paragraph, offset: int, keep: str) -> None:
    """Every ``w:fldSimple`` a split at `offset` falls in, below the paragraph, made a complex field, the innermost first.

    Those the split would cut, and those holding the character at `offset`
    (core-ts's ``unfoldSimpleFieldsAround(seg.run)``, so that a split at the
    start of a result unfolds the field too, its unmeasured choice).
    """
    for _guard in range(64):
        segments = paragraph.segments()
        at = next((s for s in segments if s.start <= offset < s.end), None)
        found = None
        if at is not None:
            found = next((h for h in holders_above(at.run, paragraph.element) if element_name(h) == _W_FLD_SIMPLE), None)
        if found is None:
            place = _place(paragraph, offset, keep, field_depths(paragraph.element))
            holders, _index, _extra = _climb(paragraph.element, place)
            found = next((h for h in holders if element_name(h) == _W_FLD_SIMPLE), None)
        if found is None:
            return
        unfold_simple_field(found)


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
    in_code = depths.exit.get(id(anchor), False) if how == "after" else depths.entry.get(id(anchor), False)
    if depths.inside(in_code):
        _refuse_field(operation)


@dataclasses.dataclass(slots=True)
class _Place:
    """Where a split falls in the tree: the list, and the index its items go to the tail from."""

    items: list
    index: int
    run_cut: Any = None
    """The run the split falls inside, which is cut first; None at a run boundary."""
    depth: bool = False
    """Whether the split is in a field's instruction (:class:`FieldDepths`)."""
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
        return _Place(after.run_owner, after.run_index + 1, after.run, depths.entry.get(id(after.item), False), after, run_before)
    if keep == "head":
        if after is None:
            return _Place(content, len(content), depth=depths.closing, run_before=run_before)
        if after.branch is None and not after.math and after.index > 0:
            return _Place(after.run_owner, after.run_index + 1, after.run, depths.entry.get(id(after.item), False), after, run_before)
        item, owner, _index = _site_of(after)
        return _Place(owner, _index_of(owner, item), depth=depths.entry.get(id(item), False), segment=after, run_before=run_before)
    if before is None:
        return _Place(content, 0, run_before=None)
    if before.branch is None and not before.math and before.index < len(before.owner) - 1:
        return _Place(before.run_owner, before.run_index + 1, before.run, depths.exit.get(id(before.item), False), before, run_before)
    item, owner, _index = _site_of(before)
    return _Place(owner, _index_of(owner, item) + 1, depth=depths.exit.get(id(item), False), segment=before, run_before=run_before)


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

    The module's rules (core-ts's ``splitParagraph`` at ``1e61bb1``, and its
    ``unfoldSimpleFieldsAround`` at ``906d5e3``). The
    original ``w:p`` keeps the text from `offset` on, its mark, its section
    break and its ``w14:paraId``; the new one, before it in the same list, takes
    a copy of the properties with a mark in the formatting of the run before
    the split, and a ``w14:paraId`` of its own where the package assigns them.
    `offset` may be 0 (an empty paragraph in front) or the text's length (the
    original left with its mark alone). `keep` is :func:`_place`'s. Tracked,
    the new mark is marked inserted.

    Raises:
        SpanError: ``range.splits_control``, ``range.splits_field`` (a
            field's instruction), ``range.cuts_equation``; nothing has changed.
        AddressError: ``address.read_only``; nothing has changed.
    """
    from docx4j_py.model.content.addresses import assign_para_id
    from docx4j_py.model.content.hyperlinks import _id_source
    from docx4j_py.model.content.reports import current_recorder, moved_by_insert
    from docx4j_py.wml import P

    refuse_split(paragraph, offset, keep, operation)
    # a w:fldSimple there is made a complex field first, whose result is then split (CR-013 section 13.7, D1)
    _unfold_simple_fields(paragraph, offset, keep)
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
