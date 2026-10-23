"""The writes of a ``Range`` across paragraphs, and of one holding its last paragraph's mark: CR-013 Phase C.

The port of docx4j-core-ts's ``spanEdits.mts`` at its committed ``HEAD``
(``f544b8f``: the file as its CR-014 phases 2 and 3 left it, ``53db5bc`` and
``87bbb38``), as Word 365 makes these writes through Office JS (core-ts checks
52 and 54). :class:`~docx4j_py.model.content.range.Range` calls in here for
every write whose answer is not one paragraph's splice; what it gets back is
where the text closed up (:class:`Landing`).

**Where a range stands** (:func:`extent_of`): the innermost list of blocks
holding both ends, and in it the block holding each end --- the end's
paragraph, or a table or block content control the range holds whole. A range
starting or ending inside a block it doesn't hold whole is refused before
anything changes: inside a table as Word refuses it (``range.partial_table``,
check 52's INTO and CELL), inside a content control as unmeasured
(``range.partial_control``), but for a deletion one control level down
(``cross_controls``, check 54 cases 18 to 20).

**A deletion** (:func:`delete_span`, check 52 cases 3 to 28, check 54 cases 5
to 20): untracked, the start's paragraph loses its text from the start, every
block between goes whole, the end's paragraph loses its text to the end, and
the two join through the resolution's own
:func:`~docx4j_py.model.revisions.resolver.join_with_next`, in the end
paragraph's properties and ``w14:paraId``, each run keeping its formatting. A
mark the range holds goes, whatever ends the range: its paragraph joins the
next, or, with none after it in its list, goes itself where its text went with
the range (before a table, a control's last: check 54 cases 8 and 12), a
cell's last paragraph and a story's last kept (case 10). Across a control's
edge nothing joins: the start's paragraph takes the end's properties (cases 18
to 20). Tracked, nothing joins: the text is a ``w:del``, every mark the range
holds marked deleted but a cell's last or a story's last, with no
``w:pPrChange``, a table's rows and their cells' marks deleted, and a whole
control deleted between Word's ``w:customXmlDelRange*`` markers (case 15).

**What a deletion takes with it** (:func:`queue_removed`,
:func:`queue_comment_removal`, :func:`settle_removals`): a note whose
reference goes, a comment whose reference goes, and the relationships the
removed content held (a section break's headers and footers, a table's
pictures) are recorded on the package and settled when the write ends,
through the resolution's own :class:`~docx4j_py.model.revisions.removals.Removals`,
against what still stands in every story --- docx4j's ``Removals`` rule.
core-ts queues them for the save; the saved document is the same.

**A replacement** (:func:`replace_span`): untracked, the deletion keeping a
last mark, then the text where the paragraphs joined in the first replaced
character's look (:func:`look_of`, check 52 cases 29 and 30, check 54 case
26); tracked, the ``w:ins`` first, before the deletion, as Word writes it
(check 52 cases 39 to 48, check 54 case 4), within one paragraph too.

**Formatting** (:func:`format_span`): every run in the range, a table's cells'
too, and each mark it holds, a property written only where the run or mark
does not read it already (check 52 cases 49 to 56, check 54 cases 21 to 25);
tracked, a mark's ``w:pPr/w:rPr/w:rPrChange``.

**A paragraph** (:func:`insert_paragraph_at`): at the range's start or end,
splitting the paragraph there (check 52 cases 57 to 70; within one paragraph
too, check 54 cases 1 to 3).
"""

from __future__ import annotations

import dataclasses
from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, deep_copy, link_parents
from docx4j_py.model.content.errors import ContentError, SpanError
from docx4j_py.model.content.story import first_and_last_paragraph, story_of
from docx4j_py.model.content.text_model import (
    PARAGRAPH_BREAKS,
    W_P,
    W_TBL,
    block_children_of,
    drop_unwritable,
    text_run,
)
from docx4j_py.traversal import element_name

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "Extent",
    "FirstCharacter",
    "Landing",
    "assert_trackable",
    "blocks_holding",
    "delete_span",
    "extent_of",
    "first_character_of",
    "format_span",
    "insert_paragraph_at",
    "is_range_marker",
    "look_of",
    "place_replacement",
    "queue_comment_removal",
    "queue_removed",
    "refuse_partial",
    "settle_removals",
]


# ---------------------------------------------------------------------------
# where a range stands
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class Extent:
    """Where a range's two ends stand in the innermost list of blocks holding both (core-ts's ``Extent``)."""

    items: list
    """The list."""
    first: int
    """The index of the block holding the start."""
    last: int
    """The index of the block holding the end."""
    head: Paragraph | None = None
    """The start's paragraph, where it is the first block itself (or in :attr:`head_inside`); else that block goes whole."""
    tail: Paragraph | None = None
    """The end's paragraph, where it is the last block itself (or in :attr:`tail_inside`); else that block goes whole."""
    head_inside: list | None = None
    """Where the range starts inside a block content control at :attr:`first`: the control's list holding :attr:`head` (check 54 case 20)."""
    tail_inside: list | None = None
    """Where the range ends inside a block content control at :attr:`last`: the control's list holding :attr:`tail` (check 54 case 18)."""


@dataclasses.dataclass(frozen=True, slots=True)
class Landing:
    """Where a deletion closed up: the paragraph, and the offset in it a replacement's text goes at."""

    paragraph: Paragraph
    offset: int


#: Not block-level, though a block container's list holds them: a table's rows, a row's cells and their wrappers.
_ROW_AND_CELL = frozenset({"CTSdtRow", "CTSdtCell", "CTCustomXmlRow", "CTCustomXmlCell", "Tr", "Tc"})

#: The block-level elements a range may end inside, one level down (:func:`extent_of`'s ``cross_controls``).
_CONTROLS = frozenset({"SdtBlock", "CTCustomXmlBlock"})


def _block_level(node: Any) -> bool:
    cls = node.__class__
    return not any(base.__name__ in _ROW_AND_CELL for base in cls.__mro__ if base.__module__ == "docx4j_py.wml")


def is_control(block: Any) -> bool:
    """A block-level content control or custom XML element (core-ts's ``CONTROLS``)."""
    cls = block.__class__
    return cls.__module__ == "docx4j_py.wml" and cls.__name__ in _CONTROLS


def blocks_holding(element: Any, story: list) -> list[tuple[list, Any]]:
    """The blocks holding a paragraph, innermost first: the paragraph in its list, each block above in its, up to the story's.

    core-ts's ``blocksHolding``.
    """
    out: list[tuple[list, Any]] = []
    node = element
    for _ in range(256):
        parent = getattr(node, "parent", None)
        if parent is None:
            break
        items = block_children_of(parent)
        if isinstance(items, list) and any(item is node for item in items) and _block_level(node):
            out.append((items, node))
            if items is story:
                break
        node = parent
    return out


def _index_of(items: list, element: Any) -> int:
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def extent_of(span: Range, member: str, cross_controls: bool = False) -> Extent:
    """The extent of a range, refusing one that starts or ends inside a block it doesn't hold whole.

    core-ts's ``extentOf``. Inside a table that is Word's own refusal (check
    52: ``delete()`` and ``Replace`` on a range from a cell to a paragraph after
    the table, or from a paragraph into a cell, threw ``GeneralException`` and
    changed nothing, tracked or not; D3). Inside a block content control,
    `cross_controls` lets a range end in the control's own paragraphs, one
    level down, as ``delete()`` takes one (check 54 cases 18 to 20); anything
    else there is refused as unmeasured.

    Raises:
        SpanError: ``range.partial_table`` or ``range.partial_control``.
        ContentError: ``range.stale``, the range's paragraphs are not in their
            story any more.
    """
    story, _owner = story_of(span.paragraph.element)
    up = blocks_holding(span.paragraph.element, story)
    down = blocks_holding(span.end_paragraph.element, story)
    if not up or not down or up[-1][0] is not story or down[-1][0] is not story:
        raise ContentError(
            "this range's paragraphs are not in the document any more: an edit through another view "
            "removed or joined them",
            code="range.stale",
            hint="take the range again",
        )
    for i, (items, a) in enumerate(up):
        j = next((k for k, (other, _b) in enumerate(down) if other is items), -1)
        if j < 0:
            continue
        b = down[j][1]
        extent = Extent(items, _index_of(items, a), _index_of(items, b))

        def inside(k: int, block: Any) -> bool:
            # a range starting or ending in a block control's own paragraph, the other end outside the control
            return cross_controls and k == 1 and a is not b and is_control(block)

        if i > 0 and not _starts_whole(span, items, extent.first):
            if not inside(i, a):
                refuse_partial(a, member, "starts")
            extent.head_inside = up[0][0]
        if j > 0 and not _ends_whole(span, items, extent.last):
            if not inside(j, b) or span.holds_end_mark:
                refuse_partial(b, member, "ends")
            extent.tail_inside = down[0][0]
        if i == 0 or extent.head_inside is not None:
            extent.head = span.paragraph
        if j == 0 or extent.tail_inside is not None:
            extent.tail = span.end_paragraph
        return extent
    raise SpanError(
        "this range's two ends are not in one story",
        code="range.other_story",
        hint="take both ends from the same body, note, comment or text box",
    )


def _starts_whole(span: Range, items: list, index: int) -> bool:
    """The range starts at the start of the block: its first paragraph, as a reader reads it, at offset 0."""
    first, _last = first_and_last_paragraph(items, index)
    return span.start == 0 and first is not None and first.element is span.paragraph.element


def _ends_whole(span: Range, items: list, index: int) -> bool:
    """The range ends at the end of the block: its last paragraph's, as a reader reads it, the mark (or a row end) included."""
    _first, last = first_and_last_paragraph(items, index)
    return span.holds_end_mark and last is not None and last.element is span.end_paragraph.element


def refuse_partial(block: Any, member: str, side: str) -> None:
    """Refuse a range starting or ending inside `block`, which it doesn't hold whole.

    Raises:
        SpanError: ``range.partial_table`` (Word's own refusal) or ``range.partial_control``.
    """
    if element_name(block) == W_TBL:
        raise SpanError(
            f"Range.{member}: this range {side} inside a table, which Word refuses too (core-ts check 52)",
            code="range.partial_table",
            hint="take the whole table, or edit inside one cell; or delete the table, then the paragraphs",
        )
    raise SpanError(
        f"Range.{member}: this range {side} inside a content control (or another block) it doesn't hold "
        "whole, and what Word does with that isn't measured",
        code="range.partial_control",
        hint="take the whole content control, or edit inside it",
    )


def is_range_marker(block: Any) -> bool:
    """A marker that stands between two paragraphs without keeping them apart (core-ts's ``isRangeMarker``)."""
    local = (element_name(block) or "").rsplit("}", 1)[-1]
    return local.endswith(("RangeStart", "RangeEnd")) or local in (
        "bookmarkStart",
        "bookmarkEnd",
        "permStart",
        "permEnd",
        "proofErr",
    )


def assert_trackable(extent: Extent, member: str) -> None:
    """Refuse, before anything changes, a tracked deletion of a range holding a whole ``w:altChunk`` or kept ``mc:AlternateContent``.

    core-ts's ``assertTrackable``: neither has a deleted form, and Accept All
    would leave it where the untracked deletion takes it, the paragraphs
    either side unjoined. A content control or custom XML element held whole
    is deleted as Word deletes one (check 54 cases 15 and 17).

    Raises:
        SpanError: ``range.not_trackable``.
    """
    for index in range(extent.first, extent.last + 1):
        if (index == extent.first and extent.head is not None) or (index == extent.last and extent.tail is not None):
            continue
        block = extent.items[index]
        name = element_name(block)
        if name in (W_P, W_TBL) or is_control(block) or is_range_marker(block):
            continue
        local = (name or type(block).__name__).rsplit("}", 1)[-1]
        raise SpanError(
            f"Range.{member}: this range holds a whole {local}, and there is no way to delete one "
            "under change tracking (core-ts CR-014 phase 2)",
            code="range.not_trackable",
            hint="turn change tracking off, or write the paragraphs either side of it",
        )


def refuse_tracked_content(extent: Extent, member: str, body: Any) -> None:
    """Refuse, before anything changes, a tracked deletion whose blocks between hold what can't be marked deleted.

    Word's branch of a kept block-, row- or cell-level ``mc:AlternateContent``
    (``address.read_only``: a mark there would change one branch alone, CR-012
    Phase A2, as a tracked ``Body.clear()`` refuses it). An equation is deleted
    in Word's form (CR-013 section 13.7, D2).

    Raises:
        AddressError: ``address.read_only``.
    """
    from docx4j_py.model.content.reports import refuse_read_only
    from docx4j_py.model.content.text_model import kept_content_under

    for index in range(extent.first, extent.last + 1):
        if (index == extent.first and extent.head is not None) or (index == extent.last and extent.tail is not None):
            continue
        block = extent.items[index]
        if is_range_marker(block):
            continue
        found = kept_content_under(block)
        if found is not None:
            refuse_read_only(body.body_for(found), member)


# ---------------------------------------------------------------------------
# what a deletion takes with it
# ---------------------------------------------------------------------------


def _real_package(package: Any) -> bool:
    if package is None:
        return False
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

    return isinstance(package, OpcPackage)


def _pending(package: Any) -> Any:
    """The package's pending :class:`~docx4j_py.model.revisions.removals.Removals`, made on first use; None for a trial's.

    A dry run's package is not a real one, and a resolution there touches
    nothing outside its own part (CR-010 section 10.3), so nothing is queued.
    """
    if not _real_package(package):
        return None
    found = package._pending_removals
    if found is None:
        from docx4j_py.model.revisions.removals import Removals

        found = Removals(package)
        package._pending_removals = found
    return found


def _settle_unless_recording(package: Any) -> None:
    """Settle at once when no write is being recorded; else :class:`~.reports.recording` settles when it ends."""
    if getattr(package, "_current_change", None) is None:
        settle_removals(package)


def queue_removed(body: Any, subtree: Any) -> None:
    """Record what removed content took with it: its notes, comments and relationships, settled when the write ends.

    core-ts's ``queueNoteRemovals`` and ``releaseRelationships``: a note whose
    reference went goes (check 52 case 12), a comment whose reference went
    goes, a comment left with half its range loses that half, and a
    relationship nothing references any more goes, its part with it --- the
    resolution's own :class:`~docx4j_py.model.revisions.removals.Removals`,
    docx4j's rule, judged against what still stands in every story when the
    write ends. A bookmark is not touched: one a deletion breaks keeps what is
    left of it, as Word's check 54 case 30 left one at the join.
    """
    if getattr(body, "copy", False) or subtree is None:
        return  # a detached copy's trim takes nothing from the document (CopyBody)
    _name_parts(body, subtree)
    removals = _pending(getattr(body, "package", None))
    if removals is None:
        return
    removals.discard_removed(getattr(body, "part", None), subtree)
    _settle_unless_recording(body.package)


def queue_comment_removal(body: Any, identifier: int) -> None:
    """A comment whose reference a deletion took, to go when the write ends, its markers and replies with it.

    core-ts's ``queueCommentRemoval`` (its CR-014 section 12.6; Word's check 54
    cases 27 and 28): Word keeps no comment without a reference (core-ts check
    32), and docx4j's ``Removals`` takes such a comment away whole.
    """
    if getattr(body, "copy", False):
        return  # a detached copy's trim takes nothing from the document (CopyBody)
    _name_comment_parts(body)
    removals = _pending(getattr(body, "package", None))
    if removals is None:
        return
    removals.comments.add(identifier)
    _settle_unless_recording(body.package)


def _main_of(body: Any) -> Any:
    package = getattr(body, "package", None)
    try:
        return getattr(package, "main_document_part", None)
    except Exception:  # noqa: BLE001 - a package without a main document part names no part
        return None


def _name_parts(body: Any, subtree: Any) -> None:
    """The write's report names the notes and comment parts the removed content's references reach.

    Named when the removal is queued, so that a dry run's report, whose
    package settles nothing (a trial touches no part outside its own), says
    what the write would touch as the real one's does (CR-003 section 3.4).
    """
    from docx4j_py.model.content.reports import current_recorder
    from docx4j_py.traversal import iter_nodes

    change = current_recorder(getattr(body, "package", None))
    if not getattr(change, "active", False):
        return
    names = set()
    for node in iter_nodes(subtree, mce="all"):
        local = (element_name(node) or "").rsplit("}", 1)[-1]
        if local in ("footnoteReference", "endnoteReference", "commentReference"):
            names.add(local)
    main = _main_of(body)
    if main is None:
        return
    if "footnoteReference" in names:
        change.part(main.footnotes_part)
    if "endnoteReference" in names:
        change.part(main.endnotes_part)
    if "commentReference" in names:
        _name_comment_parts(body)


def _name_comment_parts(body: Any) -> None:
    from docx4j_py.model.content.reports import current_recorder

    change = current_recorder(getattr(body, "package", None))
    main = _main_of(body)
    if not getattr(change, "active", False) or main is None:
        return
    change.part(
        main.comments_part,
        main.comments_extended_part,
        main.comments_ids_part,
        main.comments_extensible_part,
    )


def settle_removals(package: Any, change: Any = None) -> None:
    """Settle what the writes queued (:func:`queue_removed`): what :class:`~.reports.recording` calls as a write ends.

    `change`, the write's recorder, is told the parts the settling may touch:
    the notes part a note leaves, the comment parts a comment leaves.
    """
    if not _real_package(package):
        return
    removals = package._pending_removals
    if removals is None:
        return
    package._pending_removals = None
    if change is not None:
        main = getattr(package, "main_document_part", None)
        if main is not None:
            if removals.footnotes:
                change.part(main.footnotes_part)
            if removals.endnotes:
                change.part(main.endnotes_part)
            if removals.comments or removals.comments_whole:
                change.part(
                    main.comments_part,
                    main.comments_extended_part,
                    main.comments_ids_part,
                    main.comments_extensible_part,
                )
    removals.flush()


def _story_for(body: Any, paragraph: Paragraph) -> Any:
    """The resolution's :class:`~docx4j_py.model.revisions.story.Story` over a paragraph's story."""
    from docx4j_py.model.revisions.story import Story

    items, owner = story_of(paragraph.element)
    return Story(getattr(body, "package", None), getattr(body, "part", None), items, owner)


# ---------------------------------------------------------------------------
# delete
# ---------------------------------------------------------------------------


def delete_span(span: Range, extent: Extent, keep_last_mark: bool) -> Landing | None:
    """Deletes a range across paragraphs, or one holding its last paragraph's mark, as Word does.

    core-ts's ``deleteSpan`` (check 52 cases 3 to 28; check 54 cases 5 to 20):
    the module's rules. `keep_last_mark` is a replacement's, which keeps a last
    mark the range holds. Returns where the text closed up, or None where
    nothing of a paragraph is left at either end (a range of whole blocks).
    """
    tracker = span.paragraph.change_tracker
    if tracker is not None:
        return _delete_tracked(span, extent, tracker, keep_last_mark)
    return _delete_untracked(span, extent, keep_last_mark)


def _delete_untracked(span: Range, extent: Extent, keep_last_mark: bool) -> Landing | None:
    items, first, last, head, tail = extent.items, extent.first, extent.last, extent.head, extent.tail
    takes_mark = span.holds_end_mark and not keep_last_mark
    body = span.paragraph.parent_body
    if head is not None and tail is not None and head.element is tail.element:
        _cut(head, span.start, span.end, (span.whole, span.whole or span.holds_end_mark))
        if takes_mark:
            # what is left of the paragraph's text went to the start of the one it joined, if any
            survivor = _drop_mark(head)
            return Landing(survivor, span.start) if survivor is not None else None
        return Landing(head, span.start)
    # the end first, so that the indexes before it hold
    if tail is not None:
        _cut(tail, 0, span.end, (True, span.whole or span.holds_end_mark))
        if extent.tail_inside is not None:
            _remove_between(body, extent.tail_inside, -1, _index_of(extent.tail_inside, tail.element))
        elif takes_mark:
            survivor = _drop_mark(tail)
            if survivor is not None:
                tail = survivor
    else:
        _remove_block(body, items, last)
    # the blocks between; a range marker between blocks stays, for the join to bring into the joined paragraph
    _remove_between(body, items, first, last)
    if head is not None:
        if extent.head_inside is not None:
            _remove_between(
                body, extent.head_inside, _index_of(extent.head_inside, head.element), len(extent.head_inside)
            )
        _cut(head, span.start, len(head.text), (span.whole, True))
        if extent.head_inside is not None or extent.tail_inside is not None:
            # across a control's edge nothing joins: the start's paragraph takes the end's properties (check 54)
            if tail is not None:
                _take_properties(head, tail)
            return Landing(head, span.start)
        # the head's mark is in the range: it joins what follows, the tail or the paragraph after a block taken whole
        survivor = _join_forward(head)
        return Landing(survivor, span.start)
    if tail is not None or first < last:
        # a range of one block taken whole (a table, a control) removed it as its last above; core-ts's
        # deleteUntracked removes the block after it too, found here (CR-013 section 14.9)
        _remove_block(body, items, first)
    return Landing(tail, 0) if tail is not None else None


def _remove_between(body: Any, items: list, low: int, high: int) -> None:
    """Removes the blocks of a list strictly between two indexes, last first; a range marker stays."""
    for index in range(high - 1, low, -1):
        if 0 <= index < len(items) and not is_range_marker(items[index]):
            _remove_block(body, items, index)


def _remove_block(body: Any, items: list, index: int) -> None:
    """Removes a block, untracked: its notes, comments and relationships settled when the write ends."""
    from docx4j_py.model.content.reports import current_recorder, moved_by_delete

    if not 0 <= index < len(items):
        return
    removed = items[index]
    change = current_recorder(getattr(body, "package", None))
    if getattr(change, "active", False):
        change.touched(removed)
        change.shifted(moved_by_delete(body, items, index))
    del items[index]
    queue_removed(body, removed)


def _cut(paragraph: Paragraph, start: int, end: int, edges: tuple[bool, bool]) -> None:
    """Removes ``[start, end)`` of a paragraph's text, untracked; what it took with it queued (core-ts's ``cut``).

    A note whose reference goes is queued to go too (check 52 case 12); a
    comment whose reference goes goes with it, which the paragraph's splice
    does for every deletion (CR-014 section 12.6).
    """
    from docx4j_py.model.content.paragraph import _in_span

    removed = [
        segment.item
        for segment in paragraph.segments(objects=True)
        if (_in_span(segment.start, start, end, edges) if segment.object else segment.start < end and segment.end > start)
    ]
    if end <= start and not removed:
        return
    for item in removed:
        queue_removed(paragraph.parent_body, item)
    paragraph.splice(start, end, "", whole=edges)


def _list_of(paragraph: Paragraph) -> list:
    container = paragraph.container
    if _index_of(container, paragraph.element) >= 0:
        return container
    parent = getattr(paragraph.element, "parent", None)
    found = block_children_of(parent) if parent is not None else None
    return found if isinstance(found, list) else container


def _next_paragraph(paragraph: Paragraph) -> bool:
    """Whether the block right after `paragraph` in its list, past range markers, is a paragraph."""
    items = _list_of(paragraph)
    j = _index_of(items, paragraph.element) + 1
    while j < len(items) and is_range_marker(items[j]):
        j += 1
    return j < len(items) and element_name(items[j]) == W_P


def _removable(paragraph: Paragraph) -> bool:
    """A paragraph whose mark can go: not a cell's last, not its story's last, not its list's only block."""
    if paragraph.mark_text == "\t":
        return False
    items = _list_of(paragraph)
    blocks = [item for item in items if not is_range_marker(item)]
    if len(blocks) <= 1:
        return False
    story, _owner = story_of(paragraph.element)
    return not (items is story and blocks[-1] is paragraph.element)


def _drop_mark(paragraph: Paragraph) -> Paragraph | None:
    """A paragraph's mark the range took: joined to the next paragraph, else gone where its text went.

    core-ts's ``dropMark``: the paragraph joins the one after it where there is
    one (W1, W2); else, its text gone with the range, it goes (check 54: before
    a table, case 8; a control's last paragraph, case 12), unless it is a
    cell's last paragraph (case 10, CR-002 section 55), the last in its story
    or its list's only block, where the mark stays. Returns the paragraph
    left holding the place, or None where the paragraph went.
    """
    if _next_paragraph(paragraph):
        return _join(paragraph)
    if paragraph.text != "" or paragraph.holds_objects() or not _removable(paragraph):
        return paragraph
    items = _list_of(paragraph)
    _remove_block(paragraph.parent_body, items, _index_of(items, paragraph.element))
    return None


def _join_forward(paragraph: Paragraph) -> Paragraph:
    """The paragraph after `paragraph` in its list, past range markers, joined into it where there is one."""
    if _next_paragraph(paragraph):
        return _join(paragraph)
    return paragraph


def _join(paragraph: Paragraph) -> Paragraph:
    """`paragraph` joined with the next one in its list, as Accept All joins (the resolution's ``join_with_next``).

    The joined paragraph keeps the next one's properties and ``w14:paraId``; a
    section break in `paragraph`'s mark goes with the mark, the relationships
    its headers and footers held released when the write ends. Returns the
    view of the joined paragraph.
    """
    from docx4j_py.model.content.paragraph import Paragraph as ParagraphView
    from docx4j_py.model.content.reports import current_recorder, moved_by_delete
    from docx4j_py.model.revisions.change import Para
    from docx4j_py.model.revisions.removals import Removals
    from docx4j_py.model.revisions.resolver import join_with_next

    body = paragraph.parent_body
    items = _list_of(paragraph)
    removals = _pending(getattr(body, "package", None))
    if removals is None:
        removals = Removals(getattr(body, "package", None), local=True)
    change = current_recorder(getattr(body, "package", None))
    if getattr(change, "active", False):
        change.shifted(moved_by_delete(body, items, _index_of(items, paragraph.element)))
    joined = join_with_next(Para(paragraph.element, items), removals, _story_for(body, paragraph))
    if joined is None or joined is paragraph.element:
        return paragraph
    return ParagraphView(joined, items, body)


def _take_properties(head: Paragraph, tail: Paragraph) -> None:
    """The start's paragraph across a control's edge from the end's: it takes the end paragraph's properties.

    core-ts's ``takeProperties`` (check 54 cases 18 and 20: ``HEAD18A`` centred
    as BODY18 was; ``BODY20A`` in PLAIN20's properties, its mark italic as its
    run): its own mark kept in its last run's formatting, its section break
    kept.
    """
    from docx4j_py.model.content.tracking import copy_r_pr
    from docx4j_py.wml import ParaRPr, el, rpr_from_elements, rpr_to_elements

    sect_pr = getattr(getattr(head.element, "p_pr", None), "sect_pr", None)
    p_pr = deep_copy(tail.element.p_pr) if tail.element.p_pr is not None else el.pPr()
    p_pr.r_pr = None
    p_pr.sect_pr = None
    p_pr.p_pr_change = None
    if sect_pr is not None:
        p_pr.sect_pr = sect_pr
    runs = [segment for segment in head.segments() if segment.text and not segment.math]
    r_pr = copy_r_pr(getattr(runs[-1].run, "r_pr", None)) if runs else None
    if r_pr is not None:
        p_pr.r_pr = rpr_from_elements(rpr_to_elements(r_pr), cls=ParaRPr)
    head.element.p_pr = p_pr
    link_parents(p_pr)
    p_pr.parent = head.element


# ---------------------------------------------------------------------------
# delete, tracked
# ---------------------------------------------------------------------------


def _delete_tracked(span: Range, extent: Extent, tracker: Any, keep_last_mark: bool) -> Landing | None:
    items, first, last, head, tail = extent.items, extent.first, extent.last, extent.head, extent.tail
    takes_mark = span.holds_end_mark and not keep_last_mark
    body = span.paragraph.parent_body
    if head is not None and tail is not None and head.element is tail.element:
        head.splice(span.start, span.end, "", whole=(span.whole, span.whole or span.holds_end_mark))
        if takes_mark:
            _delete_mark(head, tracker)
        return Landing(head, span.start)
    # the blocks as they stand now: a control's deletion puts markers into the list, which moves the indexes
    first_block = items[first]
    last_block = items[last]
    between = list(items[first + 1 : last])
    if head is not None:
        head.splice(span.start, len(head.text), "", whole=(span.whole, True))
        _delete_mark(head, tracker)
        if extent.head_inside is not None:
            _mark_between(
                body,
                extent.head_inside,
                _index_of(extent.head_inside, head.element),
                len(extent.head_inside),
                tracker,
            )
    else:
        _delete_block(body, items, first_block, tracker)
    for block in between:
        if not is_range_marker(block):
            _delete_block(body, items, block, tracker)
    if tail is not None:
        if extent.tail_inside is not None:
            _mark_between(body, extent.tail_inside, -1, _index_of(extent.tail_inside, tail.element), tracker)
        tail.splice(0, span.end, "", whole=(True, span.whole or span.holds_end_mark))
        if takes_mark:
            _delete_mark(tail, tracker)
    elif last_block is not first_block or head is not None:
        # a range of one block taken whole was marked once, as its first
        _delete_block(body, items, last_block, tracker)
    if head is not None:
        return Landing(head, span.start)
    return Landing(tail, 0) if tail is not None else None


def _mark_between(body: Any, items: list, low: int, high: int, tracker: Any) -> None:
    """Marks the blocks of a list strictly between two indexes deleted; a range marker stays as it is."""
    for block in list(items[low + 1 : high]):
        if not is_range_marker(block):
            _delete_block(body, items, block, tracker)


def _delete_block(body: Any, items: list, block: Any, tracker: Any) -> None:
    """A block the range holds whole, deleted under tracking: a control with Word's markers around it, its content within."""
    from docx4j_py.model.content.tracking import _delete_blocks, mark_holder_deleted

    if is_control(block):
        mark_holder_deleted(tracker, block, items, lambda: _delete_blocks(tracker, [block], body))
    else:
        _delete_blocks(tracker, [block], body)


def _is_story_last(paragraph: Paragraph) -> bool:
    items = _list_of(paragraph)
    story, _owner = story_of(paragraph.element)
    if items is not story:
        return False
    blocks = [item for item in items if not is_range_marker(item)]
    return bool(blocks) and blocks[-1] is paragraph.element


def _delete_mark(paragraph: Paragraph, tracker: Any) -> None:
    """A paragraph's mark marked deleted, unless it is already, or can't go: a cell's last, a story's last.

    core-ts's ``deleteMark`` (check 54 case 11, as CR-002 section 55's cells):
    the skipping form of :meth:`~.tracking.ChangeTracker.mark_paragraph_deleted`.
    """
    from docx4j_py.model.content.tracking import mark_deleted

    if mark_deleted(paragraph.element) or paragraph.mark_text == "\t" or _is_story_last(paragraph):
        return
    tracker.mark_paragraph_deleted(paragraph.element)


# ---------------------------------------------------------------------------
# Replace
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class FirstCharacter:
    """The first replaced character's run properties and where it stood, read before anything is deleted."""

    r_pr: Any
    p_pr: Any
    element: Any


def first_character_of(span: Range) -> FirstCharacter:
    """The character at a range's start: its run's properties, or its paragraph's mark's where the range starts at the mark."""
    paragraph = span.paragraph
    segment = next((s for s in paragraph.segments() if s.start <= span.start < s.end and not s.math), None)
    r_pr = getattr(segment.run, "r_pr", None) if segment is not None else _mark_r_pr_of(paragraph.element)
    return FirstCharacter(r_pr, paragraph.element.p_pr, paragraph.element)


def _mark_r_pr_of(element: Any) -> Any:
    """A paragraph mark's formatting as run properties: the mark's own revisions left out."""
    from docx4j_py.wml import RPr, rpr_from_elements, rpr_to_elements

    r_pr = getattr(getattr(element, "p_pr", None), "r_pr", None)
    if r_pr is None:
        return None
    elements = rpr_to_elements(r_pr)
    return rpr_from_elements(elements, cls=RPr) if elements else None


#: Not compared by :func:`look_of`: a record, and the character style, which a run keeps of its own.
_LOOK_SKIPPED = frozenset({"r_pr_change", "r_style"})


def _neutral(name: str) -> Any:
    """What a property reads as where nothing sets it, for the properties that have one (core-ts's ``NEUTRAL``)."""
    from docx4j_py.wml import el

    toggles = (
        "b",
        "b_cs",
        "i",
        "i_cs",
        "caps",
        "small_caps",
        "strike",
        "dstrike",
        "outline",
        "shadow",
        "emboss",
        "imprint",
        "vanish",
        "web_hidden",
    )
    if name in toggles:
        factory = {"b_cs": "bCs", "i_cs": "iCs", "small_caps": "smallCaps", "web_hidden": "webHidden"}.get(name, name)
        return getattr(el, factory)(val=False)
    if name == "color":
        return el.color(val="auto")
    if name == "u":
        return el.u(val="none")
    if name == "highlight":
        return el.highlight(val="none")
    if name == "vert_align":
        return el.vertAlign(val="baseline")
    if name == "sz":
        return el.sz(val=20)
    if name == "sz_cs":
        return el.szCs(val=20)
    return None


def look_of(first: FirstCharacter, into: Paragraph) -> Any:
    """The run properties that give text in `into` the look the first replaced character had.

    core-ts's ``lookOf`` (check 52 cases 29 and 30; check 54 case 26): its
    run's own properties, and each property its paragraph's style gave it
    that `into`'s doesn't, written direct --- a heading's font, size, colour
    and bold in a body paragraph (case 29), and a body paragraph's
    ``w:b w:val="0"`` and automatic colour in a heading (case 30). A property
    no style sets where it was and `into`'s does is written as its neutral
    value where it has one (``w:caps w:val="0"``, automatic colour, 10 points,
    ``w:u w:val="none"``: check 54 case 26).
    """
    from docx4j_py.model.content.tracking import copy_r_pr
    from docx4j_py.model.property_resolver import resolver_of
    from docx4j_py.wml import RPr

    direct = copy_r_pr(first.r_pr)
    resolver = resolver_of(into.parent_body.package)
    # the same paragraph properties (a range within one paragraph): its own properties give it its look
    if resolver is None or first.p_pr is into.element.p_pr:
        return direct
    was = resolver.get_effective_rpr(first.r_pr, first.p_pr)
    now = resolver.get_effective_rpr(first.r_pr, into.element.p_pr)
    out = direct if direct is not None else RPr()
    for field in dataclasses.fields(RPr):
        name = field.name
        if name in _LOOK_SKIPPED:
            continue
        before, after = getattr(was, name, None), getattr(now, name, None)
        if before == after:
            continue
        value = deep_copy(before) if before is not None else _neutral(name)
        if value is not None:
            setattr(out, name, value)
    link_parents(out)
    return out if any(getattr(out, f.name, None) is not None for f in dataclasses.fields(RPr)) else None


def place_replacement(landing: Landing, first: FirstCharacter, text: str) -> tuple[Paragraph, Paragraph, int, int]:
    """The text of an untracked ``Replace``, put where the deletion closed up, in the first replaced character's look.

    core-ts's ``placeReplacement``: Office JS's characters written back as the
    elements they read as (CR-013 Q2), the excluded ones dropped (section 11,
    E3), a line feed or carriage return a paragraph break (E1). Returns the
    first and last paragraph the text is in, its start in the first and its
    end in the last.
    """
    from docx4j_py.model.content.split import split_paragraph

    text = drop_unwritable(text)
    lines = PARAGRAPH_BREAKS.split(text)
    joined = "".join(lines)
    paragraph = landing.paragraph
    if joined:
        paragraph.insert_items_at(landing.offset, [text_run(joined, look_of(first, paragraph))])
    if len(lines) == 1:
        return paragraph, paragraph, landing.offset, landing.offset + len(joined)
    position = landing.offset
    edge = landing.offset + len(joined)
    consumed = 0
    first_paragraph: Paragraph | None = None
    for line in lines[:-1]:
        position += len(line)
        keep = "tail" if joined and position == edge else "head"
        head = split_paragraph(paragraph, position - consumed, keep=keep, operation="insert_text")
        if first_paragraph is None:
            first_paragraph = head
        consumed = position
    return first_paragraph or paragraph, paragraph, landing.offset, len(lines[-1])


# ---------------------------------------------------------------------------
# insert_paragraph
# ---------------------------------------------------------------------------


def insert_paragraph_at(paragraph: Paragraph, offset: int, text: str) -> tuple[Paragraph, Paragraph | None]:
    """A paragraph of `text` put at `offset` in `paragraph`, as Word puts one where a range starts or ends.

    core-ts's ``insertParagraphAt`` (check 52 cases 57 to 70; check 54 cases 1
    to 3): at the paragraph's start it goes before it and at its end after it,
    in its properties (:meth:`Paragraph.insert_paragraph`); inside it the
    paragraph is split there (:func:`~.split.split_paragraph`) and the new
    paragraph goes between the halves, its mark and its text in the formatting
    of the run before the split (case 63's italic). Tracked, the new paragraph
    and the first half's mark are inserted (cases 69 and 70). Returns the new
    paragraph, and the first half where `paragraph` was split.
    """
    from docx4j_py.model.content.split import _properties_for_head, refuse_split, split_paragraph
    from docx4j_py.model.content.text_model import for_writer
    from docx4j_py.model.content.tracking import copy_r_pr
    from docx4j_py.wml import P

    text = drop_unwritable(text)
    if offset <= 0:
        return paragraph.insert_paragraph(text, location="Before"), None
    if offset >= len(paragraph.text):
        return paragraph.insert_paragraph(text, location="After"), None
    refuse_split(paragraph, offset, "head", "insert_paragraph")
    r_pr = _r_pr_before(paragraph, offset)
    first_half = split_paragraph(paragraph, offset, keep="head", operation="insert_paragraph")
    p_pr = _properties_for_head(paragraph.element, r_pr)
    content = ChildList()
    if text:
        content.append(text_run(for_writer(text, "one-paragraph"), copy_r_pr(r_pr)))
    new = P(content=content)
    if p_pr is not None:
        new.p_pr = p_pr
    link_parents(new)
    inserted = paragraph.parent_body.insert_element(new, location="Before", target=paragraph)
    return inserted, first_half


def _r_pr_before(paragraph: Paragraph, offset: int) -> Any:
    """The run properties of the text just before `offset`, which a split there gives its new marks (check 52 case 63)."""
    before = [s for s in paragraph.segments() if s.start < offset and s.text and not s.math]
    return getattr(before[-1].run, "r_pr", None) if before else None


# ---------------------------------------------------------------------------
# font
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class FontPiece:
    """A piece of a range in one paragraph: its runs, and whether the range holds the paragraph's mark."""

    paragraph: Paragraph
    start: int
    end: int
    runs: Callable[[], list[Any]]
    holds_mark: bool


def _reads(r_pr: Any, options: dict[str, Any]) -> bool:
    """True where run properties already read every option as given, so that nothing is written.

    core-ts's ``reads``: Word wrote ``w:b`` on no heading run (check 52),
    ``w:sz`` on none of a 14-point heading's (check 54 case 23), ``w:i`` on no
    run italic through its style (case 21). A colour or a name compares
    case-insensitively, a colour with or without its ``#``.
    """
    from docx4j_py.wml import read_run_options

    read = read_run_options(r_pr)

    def same(a: Any, b: Any) -> bool:
        if isinstance(a, str) and isinstance(b, str):
            return a.lstrip("#").lower() == b.lstrip("#").lower()
        if isinstance(b, bool) or isinstance(a, bool):
            return bool(a) == bool(b)
        return a == b

    for key, value in options.items():
        name = "name" if key == "font" else key
        if name == "underline" and value is True:
            value = "Single"
        elif name == "underline" and value is False:
            value = "None"
        if not same(read.get(name), value):
            return False
    return True


def format_span(pieces: list[FontPiece], options: dict[str, Any]) -> None:
    """A ``font`` write over a range across paragraphs, as Word formats one.

    core-ts's ``formatSpan`` (check 52 cases 49 to 56; check 54 cases 21 to
    25): every run in the range, a table's cells' too, split at the range's
    ends so that nothing outside it is formatted, and each mark the range
    holds. A property goes only where a run or mark doesn't read the value
    already (:func:`_reads`, effective): no ``w:b`` on a heading's runs bold
    through its style, and no tracked change there. Tracked, each run and
    mark changed has its ``w:rPrChange`` (case 56: the body paragraph's
    mark's).
    """
    from docx4j_py.model.content.font import Font
    from docx4j_py.model.property_resolver import resolver_of

    for piece in pieces:
        paragraph = piece.paragraph
        resolver = resolver_of(paragraph.parent_body.package)
        p_pr = paragraph.element.p_pr
        if piece.start < piece.end:
            paragraph.split_runs_around(piece.start, piece.end)

        def effective(r_pr: Any, resolver: Any = resolver, p_pr: Any = p_pr) -> Any:
            return resolver.get_effective_rpr(r_pr, p_pr) if resolver is not None else r_pr

        runs = [run for run in piece.runs() if not _reads(effective(getattr(run, "r_pr", None)), options)]
        if runs:
            Font(lambda runs=runs: runs, tracking=paragraph.font_tracking)._write(**options)
        if piece.holds_mark:
            mark = (
                resolver.get_effective_paragraph_mark_rpr(p_pr)
                if resolver is not None
                else getattr(p_pr, "r_pr", None)
            )
            if not _reads(mark, options):
                _format_mark(paragraph, options)


def _format_mark(paragraph: Paragraph, options: dict[str, Any]) -> None:
    """A paragraph mark's direct formatting written; tracked, its ``w:rPrChange`` first, unless the mark is this author's insertion."""
    from docx4j_py.model.content.tracking import para_r_pr_of
    from docx4j_py.wml import apply_run_options

    r_pr = para_r_pr_of(paragraph.element)
    tracker = paragraph.change_tracker
    if tracker is not None:
        ins = getattr(r_pr, "ins", None)
        if ins is None or str(getattr(ins, "author", "") or "") != tracker.author:
            tracker.record_mark_r_pr_change(r_pr)
    apply_run_options(r_pr, **options)
    link_parents(r_pr)
