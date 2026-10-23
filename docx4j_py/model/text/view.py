"""What Accept All or Reject All would leave of a document, worked out without resolving anything.

CR-011 section 3.2 item 2, docx4j's ``org.docx4j.model.revisions.ResolvedView``
(its CR-046 D2), as core-ts CR-012 ports it over its own resolution: here over
CR-010's engine, :mod:`docx4j_py.model.revisions`. It lists the stories'
changes with that engine's ``Lister`` (every branch of a text box, its
islands shared with the walk so that both see one typed tree) and records,
change by change, what the engine's ``Resolver`` would do to the text: a
revision taken away; a move's halves, and the text typed into a rejected
move's destination; an inserted field rejected whole; a row or a cell taken
away, and a table left with no row; an equation's run-level revision or
structure; a paragraph mark taken away; an inserted content control rejected,
its content left in its place, and one deleted whole accepted, gone with it;
and a comment a resolution takes whole (docx4j ``a8bc6cda9``, CR-010 section
12.9). What a resolution does to range markers and parts is not recorded: it
changes no text. Nor is what it does to properties, but for Word 365's Accept
All rules where a mark taken away has no paragraph to join (docx4j CR-043
section 17, ``29dbf5964`` and ``e55a83b8e``; CR-010 sections 13 and 13.3a),
which :meth:`ResolvedView.removes`, :meth:`ResolvedView.joins_into_table` and
:meth:`ResolvedView.paragraph_properties_of` answer as the resolver's
``_no_paragraph_next`` decides them.

The walk (:mod:`.extractor`) asks this view what is there, so the text it
reads in a view is the text of the resolved document (REQ-402) and the
package is untouched (REQ-204). The agreement test (REQ-406,
``tests/text/test_agreement.py``) holds the walk equal to the text of a copy
the engine resolved: where the two differ, this module is wrong.

A snapshot over the trees it was made from.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.child import is_alternate_content
from docx4j_py.model.revisions import trees
from docx4j_py.model.revisions.change import (
    CellT,
    FieldT,
    GroupT,
    HolderT,
    MarkT,
    MathRunT,
    MathStructT,
    RowT,
    RunT,
    TrackedChange,
)
from docx4j_py.model.revisions.trees import Revision

__all__ = ["ResolvedView"]


def _is_marker(item: Any) -> bool:
    local = trees.name(item)
    return local.endswith(("RangeStart", "RangeEnd")) or local in (
        "bookmarkStart",
        "bookmarkEnd",
        "permStart",
        "permEnd",
        "proofErr",
    )


def _marker_at(items: list, start: int, stop: int, local: str, comment_id: int | None) -> int:
    for index in range(start, stop):
        entry = items[index]
        if not isinstance(entry, str) and trees.name(entry) == local and trees.marker_id(entry) == comment_id:
            return index
    return -1


class ResolvedView:
    """The document after Accept All (``accepted``) or Reject All, as a plan over the unresolved tree."""

    __slots__ = (
        "_candidates",
        "_comments_gone",
        "_decided",
        "_marks_gone",
        "_moved_marks_gone",
        "_removed",
        "_spans",
        "_unwrapped",
        "accepted",
    )

    def __init__(self, changes: list[TrackedChange], accepted: bool) -> None:
        """The plan over a listing of every story (``Lister(package, False, stories).list()``)."""
        self.accepted = accepted
        #: List members gone, by ``id`` of the member.
        self._removed: set[int] = set()
        #: Content controls and custom XML elements gone, their content standing in their place.
        self._unwrapped: set[int] = set()
        #: Paragraphs whose mark is gone.
        self._marks_gone: set[int] = set()
        #: Moved paragraphs whose mark is gone: gone themselves when nothing is left in them, else joined.
        self._moved_marks_gone: set[int] = set()
        self._candidates: list[_CommentCandidate] = []
        self._spans: list[_SpanCandidate] = []
        self._comments_gone: set[int] = set()
        #: What :meth:`removes` decided for a table or a paragraph with no paragraph next, by ``id``: a snapshot's.
        self._decided: dict[int, bool] = {}
        # moves first, story by story, in the resolution's order: their outcome depends on it
        moves: dict[int, tuple[Any, list[TrackedChange]]] = {}
        for change in changes:
            self._collect_moves(change, moves)
        unreplayed: set[int] = set()
        for story, halves in moves.values():
            for missing in _MoveReplay(self, story, halves).run():
                unreplayed.add(id(missing))
        for change in changes:
            self._plan(change, unreplayed)
        for candidate in self._candidates:
            if candidate.range_empty():
                self._comments_gone.add(candidate.comment_id)
        for span in self._spans:
            if span.range_empty():
                self._comments_gone.add(span.comment_id)

    # -- what the walk asks ------------------------------------------------------

    def removes(self, item: Any) -> bool:
        """Whether a list member is gone in this view.

        A run-level revision taken away with its content, a run, a row, a cell,
        a paragraph, an equation's run-level revision or structure, a content
        control deleted whole. A revision that is not gone is kept: its content
        is read as if it were not there.

        A table whose every row the view removes goes with them, as the
        resolver removes a table left with no row (check 54 case 6: a deletion
        from a heading across a whole table joins the heading with the paragraph
        after it). And an emptied paragraph whose mark is gone goes when a table
        or a block control is next in its list (check 54 case 9, the text-left
        probe's case 3; docx4j ``ResolvedView.removes`` at ``e55a83b8e``, CR-010
        section 13.3a). One with text left and a table next is not removed but
        moved: :meth:`joins_into_table`, which the walk applies.
        """
        if id(item) in self._removed:
            return True
        tbl_class, p_class = trees.cls("Tbl"), trees.cls("P")
        if isinstance(item, tbl_class):
            return self._decide(item, self._emptied_table)
        if not isinstance(item, p_class):
            return False
        if id(item) in self._moved_marks_gone and not self._remains_in(item):
            return True
        return self.removes_mark(item) and self._decide(item, self._goes_before_a_block)

    def joins_into_table(self, p: Any) -> tuple[Any, Any] | None:
        """The table a paragraph's content joins, and the paragraph in it that takes it; None where it joins none.

        Where its mark is gone, text is left and a table is the next block in
        its list: Word 365's Accept All put the content at the front of the
        first cell's first paragraph, which kept its own properties (the
        text-left probe's case 1; docx4j ``ResolvedView.joinsIntoTable``,
        ``e55a83b8e``; CR-010 section 13.3a). That paragraph is the one the
        resolver finds once the table is resolved, last first: past the rows,
        cells and paragraphs this view removes. A kept ``mc:AlternateContent``
        next takes no content, as the resolver's doesn't.
        """
        if not self.removes_mark(p) or self._emptied(p):
            return None
        container = trees.list_of(getattr(p, "parent", None))
        k = trees.index_of(container, p)
        if k < 0:
            return None
        tbl = self._next_block(container, k)
        if not isinstance(tbl, trees.cls("Tbl")):
            return None
        target = self._first_cell_paragraph(tbl)
        return (tbl, target) if target is not None else None

    def paragraph_properties_of(self, p: Any) -> Any:
        """A paragraph's properties in this view: its own, but where its mark is gone with no paragraph next.

        Word 365's Accept All (check 54; docx4j
        ``ResolvedView.paragraphPropertiesOf`` at ``29dbf5964``, CR-010 section
        13), as the resolver's ``_no_paragraph_next`` decides it: with a block
        content control next (a kept block-level ``mc:AlternateContent`` read
        as the branch Word takes), the properties of the control's first
        paragraph, as a join would have given them (check 54 case 19, the
        text-left probe's case 5); the last paragraph of a control's content,
        text left or not, none: a bare mark (check 54 case 13, the probe's
        case 2; ``e55a83b8e``). A paragraph that joins the next, or a table's
        first cell, keeps its own: the walk takes the last of the paragraphs
        joined.
        """
        own = getattr(p, "p_pr", None)
        if not self.removes_mark(p):
            return own
        container = trees.list_of(getattr(p, "parent", None))
        k = trees.index_of(container, p)
        if k < 0:
            return own
        following = self._next_block(container, k)
        if isinstance(following, trees.cls("P")):
            return own  # a join: the walk takes the chain's last
        block = self._as_word_reads(following)
        if block is not None and trees.is_holder(block):
            first = self._first_paragraph(trees.children(block))
            if first is not None:
                return self.paragraph_properties_of(first)
        if following is None and isinstance(getattr(p, "parent", None), trees.cls("SdtContentBlock")):
            return None  # a bare mark, text or not
        return own

    def unwraps(self, item: Any) -> bool:
        """Whether a content control or custom XML element is gone with its content left in its place (one inserted, rejected)."""
        return id(item) in self._unwrapped

    def _decide(self, item: Any, rule: Any) -> bool:
        """`rule` over `item`, decided once: the view is a snapshot, and a paragraph's answer reads the blocks after it."""
        found = self._decided.get(id(item))
        if found is None:
            found = self._decided[id(item)] = rule(item)
        return found

    def _emptied_table(self, tbl: Any) -> bool:
        """Whether the view removes every row of a table, which has some, and nothing else is in it but range markers."""
        rows = False
        tr_class = trees.cls("Tr")
        for item in tbl.content:
            if isinstance(item, tr_class):
                rows = True
                if id(item) not in self._removed:
                    return False
            elif not _is_marker(item):
                return False
        return rows

    def _goes_before_a_block(self, p: Any) -> bool:
        """An emptied paragraph, its mark gone, with a table or a block control next (a kept branch read as Word reads it).

        Emptied paragraphs whose marks are gone, one after another, are decided
        together, as the resolver, last first, removes each in turn: all go
        when the block after the last of them is a table or a control. A loop,
        not a recursion through :meth:`removes`, so that a long deletion is no
        deeper a call than a short one.
        """
        container = trees.list_of(getattr(p, "parent", None))
        k = trees.index_of(container, p)
        if k < 0 or not self._emptied(p):
            return False
        p_class, tbl_class = trees.cls("P"), trees.cls("Tbl")
        run = [p]
        following = None
        for item in container[k + 1 :]:
            if _is_marker(item) or id(item) in self._removed:
                continue
            if isinstance(item, p_class):
                if id(item) in self._moved_marks_gone and not self._remains_in(item):
                    continue
                decided = self._decided.get(id(item))
                if decided is None and self.removes_mark(item) and self._emptied(item):
                    run.append(item)
                    continue
                if decided:
                    continue
            elif isinstance(item, tbl_class) and self._decide(item, self._emptied_table):
                continue
            following = item
            break
        block = self._as_word_reads(following)
        goes = isinstance(block, tbl_class) or (block is not None and trees.is_holder(block))
        for q in run:
            self._decided[id(q)] = goes
        return goes

    def _first_paragraph(self, items: list | None) -> Any:
        """A control's first paragraph as the resolved copy has it when the mark before the control is resolved.

        The resolver's ``_first_paragraph``, before any table, over the control
        as resolved already, last first: past what this view removes, and
        through the paragraphs it joins to the last of them, whose properties
        the joined paragraph keeps. docx4j's view reads the first ``w:p`` as it
        stands; the two differ only where the control's own content changes.
        """
        p_class, tbl_class = trees.cls("P"), trees.cls("Tbl")
        items = items or []
        for index, item in enumerate(items):
            if isinstance(item, tbl_class) and not self.removes(item):
                return None
            if not isinstance(item, p_class) or self.removes(item):
                continue
            while self.removes_mark(item):
                following = self._next_block(items, index)
                if not isinstance(following, p_class):
                    break
                item = following
                index = trees.index_of(items, item)
            return item
        return None

    def _first_cell_paragraph(self, tbl: Any) -> Any:
        """The resolver's ``_first_cell_paragraph`` over the table as resolved already: past what this view removes."""
        tr_class, tc_class, p_class, tbl_class = trees.cls("Tr"), trees.cls("Tc"), trees.cls("P"), trees.cls("Tbl")
        for row in tbl.content:
            if not isinstance(row, tr_class) or self.removes(row):
                continue
            for cell in row.content:
                if not isinstance(cell, tc_class) or self.removes(cell):
                    continue
                for item in cell.content:
                    if isinstance(item, tbl_class) and not self.removes(item):
                        return None
                    if isinstance(item, p_class) and not self.removes(item):
                        return item
                return None
            return None
        return None

    def _emptied(self, p: Any) -> bool:
        """Whether nothing but range markers is left in a paragraph once the view's removals are made.

        The resolver's test (``_no_paragraph_next``, docx4j's
        ``hasTextOrContent``), where docx4j's view asks ``remainsIn``, which
        counts a range marker as content: a comment's markers kept from a
        deletion, a bookmark. So a paragraph left holding only those goes here
        as it goes from the resolved copy (CR-010 section 14.4).
        """
        for item in p.content:
            if _is_marker(item) or id(item) in self._removed:
                continue
            return False
        return True

    def _next_block(self, container: list, k: int) -> Any:
        """The block after `container[k]`, past range markers and what this view removes; None at the end."""
        for index in range(k + 1, len(container)):
            item = container[index]
            if _is_marker(item) or self.removes(item):
                continue
            return item
        return None

    def _as_word_reads(self, item: Any) -> Any:
        """A kept block-level ``mc:AlternateContent`` as the first block of the branch Word takes; any other block itself.

        The resolver's ``_as_word_reads``, past what this view removes too.
        """
        if item is None or not is_alternate_content(item):
            return item
        items = trees.branch_items(item) or []
        return self._next_block(items, -1)

    def removes_mark(self, p: Any) -> bool:
        """Whether a paragraph's mark is gone in this view: it joins the next paragraph in its list, if there is one."""
        return p is not None and (id(p) in self._marks_gone or id(p) in self._moved_marks_gone)

    def removes_comment(self, comment_id: Any) -> bool:
        """Whether a comment goes in this view, entry and reference (CR-010 section 12.9)."""
        try:
            return comment_id is not None and int(comment_id) in self._comments_gone
        except (TypeError, ValueError):
            return False

    # -- the plan, change by change, as the resolver resolves each -----------------

    def _collect_moves(self, change: TrackedChange, moves: dict[int, tuple[Any, list[TrackedChange]]]) -> None:
        t = change.target
        if isinstance(t, GroupT):
            for piece in t.pieces:
                self._collect_moves(piece, moves)
        elif isinstance(t, RowT):
            for inner in t.inner:
                self._collect_moves(inner, moves)
        elif isinstance(t, RunT) and t.revision in (Revision.MOVE_FROM, Revision.MOVE_TO):
            story = change.story
            entry = moves.setdefault(id(story.root), (story, []))
            entry[1].append(change)

    def _plan(self, change: TrackedChange, unreplayed: set[int]) -> None:
        t = change.target
        if isinstance(t, RunT):
            if t.revision in (Revision.MOVE_FROM, Revision.MOVE_TO) and id(change) not in unreplayed:
                return  # replayed
            kept = (t.revision in (Revision.INS, Revision.MOVE_TO)) == self.accepted
            if not kept:
                self.remove_revision(t.item, t.owner)
        elif isinstance(t, MarkT):
            if change.para is not None and t.ins != self.accepted:
                self._marks_gone.add(id(change.para.p))
        elif isinstance(t, RowT):
            if t.ins == self.accepted:
                for inner in t.inner:
                    self._plan(inner, unreplayed)
            else:
                self._removed.add(id(t.tr))
        elif isinstance(t, GroupT):
            for piece in t.pieces:
                self._plan(piece, unreplayed)
        elif isinstance(t, FieldT):
            # Word's inserted field: rejected, the whole field goes, its separator and result too
            r_class = trees.cls("R")
            for element in t.span:
                if trees.revision_of(element) is Revision.INS:
                    if not self.accepted:
                        self.remove_revision(element, t.owner)
                elif not self.accepted and isinstance(element, r_class):
                    self._removed.add(id(element))
        elif isinstance(t, MathRunT):
            if t.ins != self.accepted:
                self._removed.add(id(t.item))
        elif isinstance(t, MathStructT):
            if t.ins != self.accepted and t.struct_item is not None:
                self._removed.add(id(t.struct_item))
        elif isinstance(t, CellT):
            if t.ins != self.accepted:
                self._removed.add(id(t.tc))
        elif isinstance(t, HolderT):
            if t.ins:
                # an inserted content control keeps its content either way; rejected, the control itself goes
                if not self.accepted:
                    self._unwrapped.add(id(t.holder))
            elif self.accepted:
                # one deleted whole goes with everything in it (check 54 cases 15 and 17; docx4j 29dbf5964)
                self._removed.add(id(t.holder))
        # property records change no text

    def remove_revision(self, item: Any, owner: list | None) -> None:
        """``Resolver.remove_revision``: the revision goes, and the comments it may take with it are candidates.

        A comment whose range started inside it, or whose markers lie either
        side of it in its list; and one with a marker in its list on the
        revision's side and the other in another paragraph (CR-043 U23).
        """
        self._removed.add(id(item))
        inner = trees.run_items(item)
        if inner is None or owner is None:
            return
        at = trees.index_of(owner, item)
        if at < 0:
            return
        for entry in inner:
            if trees.name(entry) != "commentRangeStart":
                continue
            comment_id = trees.marker_id(entry)
            if _marker_at(owner, at + 1, len(owner), "commentRangeEnd", comment_id) >= 0:
                self._candidates.append(_CommentCandidate(self, comment_id, owner, at))
            else:
                self._spans.append(_SpanCandidate(self, comment_id, item, owner, at, True))
        for s in range(at):
            if trees.name(owner[s]) != "commentRangeStart":
                continue
            comment_id = trees.marker_id(owner[s])
            if _marker_at(owner, s + 1, len(owner), "commentRangeEnd", comment_id) < 0:
                self._spans.append(_SpanCandidate(self, comment_id, item, owner, s + 1, True))
            elif _marker_at(owner, at + 1, len(owner), "commentRangeEnd", comment_id) >= 0:
                self._candidates.append(_CommentCandidate(self, comment_id, owner, s))
        for e in range(at + 1, len(owner)):
            if trees.name(owner[e]) != "commentRangeEnd":
                continue
            comment_id = trees.marker_id(owner[e])
            if _marker_at(owner, 0, e, "commentRangeStart", comment_id) < 0:
                self._spans.append(_SpanCandidate(self, comment_id, item, owner, e, False))

    def remove(self, item: Any) -> None:
        """Records a list member gone (a move's half, the text typed into a rejected move's destination)."""
        self._removed.add(id(item))

    def mark_gone(self, p: Any) -> None:
        """Records a paragraph mark gone."""
        self._marks_gone.add(id(p))

    def moved_mark_gone(self, p: Any) -> None:
        """Records a moved paragraph's mark gone."""
        self._moved_marks_gone.add(id(p))

    def has_text(self, item: Any) -> bool:
        """Whether run-level content has text in this view, once its removals are made."""
        if item is None or isinstance(item, str) or id(item) in self._removed:
            return False
        if isinstance(item, trees.cls("R")):
            return trees.has_text(item)
        items = trees.run_items(item)
        return items is not None and any(self.has_text(entry) for entry in items)

    def _remains_in(self, p: Any) -> bool:
        """Whether anything is left in a moved paragraph once the removals are made; a comment's markers stay behind."""
        for item in p.content:
            if id(item) not in self._removed:
                return True
            inner = trees.run_items(item)
            if (
                inner is not None
                and trees.revision_of(item) is not None
                and any(trees.name(entry) in ("commentRangeStart", "commentRangeEnd") for entry in inner)
            ):
                return True
        return False


class _CommentCandidate:
    """A comment whose range held a revision taken away, and where its range's text is looked for."""

    __slots__ = ("at", "comment_id", "owner", "view")

    def __init__(self, view: ResolvedView, comment_id: int | None, owner: list, at: int) -> None:
        self.view = view
        self.comment_id = comment_id
        self.owner = owner
        self.at = at

    def range_empty(self) -> bool:
        """``range_is_empty`` on what the view leaves: nothing with text before the range's end, in the same list."""
        for index in range(self.at + 1, len(self.owner)):
            item = self.owner[index]
            if trees.name(item) == "commentRangeEnd" and trees.marker_id(item) == self.comment_id:
                return True
            if self.view.has_text(item):
                return False
        return False  # the end is elsewhere: text may be


class _SpanCandidate:
    """A comment whose range held a revision taken away, its start and end in different paragraphs of one list.

    Gone, as the resolution's flush judges it once the joins are made
    (``Removals.remove_comment_if_emptied``), when every paragraph from the
    start's up to the end's loses its mark, with nothing between them but
    range markers and what the view removes, and nothing with text in the
    range.
    """

    __slots__ = ("at", "comment_id", "owner", "revision", "start_here", "view")

    def __init__(
        self, view: ResolvedView, comment_id: int | None, revision: Any, owner: list, at: int, start_here: bool
    ) -> None:
        self.view = view
        self.comment_id = comment_id
        self.revision = revision
        self.owner = owner
        self.at = at
        self.start_here = start_here

    def range_empty(self) -> bool:
        p_class = trees.cls("P")
        here = getattr(self.revision, "parent", None)
        if not isinstance(here, p_class) or here.content is not self.owner:
            return False
        container = trees.list_of(getattr(here, "parent", None))
        if container is None:
            return False
        k = trees.index_of(container, here)
        if k < 0:
            return False
        view = self.view
        if self.start_here:
            for index in range(self.at, len(self.owner)):
                if view.has_text(self.owner[index]):
                    return False
            p = here
            while True:
                if not view.removes_mark(p):
                    return False  # a mark left between start and end (U24)
                k += 1
                while k < len(container) and self._skipped(container[k]):
                    if self._is(container[k], "commentRangeEnd"):
                        return True
                    k += 1
                if k >= len(container) or not isinstance(container[k], p_class):
                    return False
                p = container[k]
                items = p.content
                end = _marker_at(items, 0, len(items), "commentRangeEnd", self.comment_id)
                stop = len(items) if end < 0 else end
                for index in range(stop):
                    if view.has_text(items[index]):
                        return False
                if end >= 0:
                    return True
        for index in range(self.at):
            if view.has_text(self.owner[index]):
                return False
        while True:
            k -= 1
            while k >= 0 and self._skipped(container[k]):
                if self._is(container[k], "commentRangeStart"):
                    return True
                k -= 1
            if k < 0 or not isinstance(container[k], p_class):
                return False
            p = container[k]
            if not view.removes_mark(p):
                return False  # a mark left between start and end (U24)
            items = p.content
            start = _marker_at(items, 0, len(items), "commentRangeStart", self.comment_id)
            for index in range(start + 1, len(items)):
                if view.has_text(items[index]):
                    return False
            if start >= 0:
                return True

    def _skipped(self, item: Any) -> bool:
        """A range marker between paragraphs, which a join carries along, or a member the view removes."""
        return _is_marker(item) or (not isinstance(item, trees.cls("P")) and self.view.removes(item))

    def _is(self, item: Any, local: str) -> bool:
        return trees.name(item) == local and trees.marker_id(item) == self.comment_id


class _MoveReplay:
    """The moves of one story resolved as the engine resolves them (``_Moves.resolve``, ``resolve_rangeless``).

    Last half first; a half in a named range resolves its whole move, a half
    in none is an insertion or a deletion, and after a broken move what is
    left unpaired becomes the plain ``w:ins`` or ``w:del`` it stands for,
    resolved as such. docx4j ``ResolvedView.MoveReplay``.
    """

    __slots__ = ("done", "halves", "pieces", "plain", "view")

    def __init__(self, view: ResolvedView, story: Any, halves: list[TrackedChange]) -> None:
        from docx4j_py.model.revisions.resolver import _Moves

        self.view = view
        self.halves = halves
        self.pieces = _Moves.pieces_of(story.root)
        #: Pieces resolved, or turned plain: no longer move markup (by id).
        self.done: set[int] = set()
        #: Halves and marks turned into plain revisions.
        self.plain: list[Any] = []

    def run(self) -> list[TrackedChange]:
        """Replays the moves; returns the halves the story has no piece for, which resolve as plain revisions."""
        missing: list[TrackedChange] = []
        for change in reversed(self.halves):
            t = change.target
            half = next((p for p in self.pieces if p.kind == "HALF" and p.item is t.item), None)
            if half is None:
                missing.append(change)
                continue
            if id(half) in self.done:
                continue
            if half.names:
                self._named(half)
            else:
                self._rangeless(half, change)
        # what became plain is resolved as plain markup, in the next round
        for piece in self.plain:
            ins = not piece.from_
            if ins == self.view.accepted:
                continue
            if piece.kind == "HALF":
                self.view.remove_revision(piece.item, piece.owner)
            elif piece.kind == "MARK":
                self.view.mark_gone(piece.para.p)
        return missing

    def _live(self, piece: Any) -> bool:
        return id(piece) not in self.done

    def _named(self, half: Any) -> None:
        name = half.names[0]
        ids = {p.id for p in self.pieces if self._live(p) and p.kind == "START" and p.name == name}
        mine = [
            p
            for p in self.pieces
            if self._live(p)
            and (
                (p.kind == "START" and p.name == name)
                or (p.kind == "END" and p.id in ids)
                or (p.kind in ("HALF", "MARK", "INNER") and name in p.names)
            )
        ]
        orphan = not any(p.from_ for p in mine)
        kept_from = not (self.view.accepted or orphan)
        broken = orphan or not any(not p.from_ for p in mine)
        for p in mine:
            self.done.add(id(p))
            if p.kind == "INNER" and not p.from_ and kept_from:
                self.view.remove(p.item)
            if p.kind == "HALF" and p.from_ != kept_from:
                self.view.remove(p.item)
            if p.kind == "MARK" and p.from_ != kept_from:
                self.view.moved_mark_gone(p.para.p)
        if broken:
            self._dissolve_broken()

    def _rangeless(self, half: Any, change: TrackedChange) -> None:
        keep = (not half.from_) == self.view.accepted
        self.done.add(id(half))
        if not keep:
            self.view.remove(half.item)
        para = change.para
        if para is not None:
            for p in self.pieces:
                if (
                    not self._live(p)
                    or p.kind != "MARK"
                    or p.from_ != half.from_
                    or p.para.p is not para.p
                    or p.names
                ):
                    continue
                self.done.add(id(p))
                if not keep:
                    self.view.moved_mark_gone(p.para.p)
                break
        self._dissolve_broken()

    def _dissolve_broken(self) -> None:
        named_from = {p.name for p in self.pieces if self._live(p) and p.kind == "START" and p.from_}
        named_to = {p.name for p in self.pieces if self._live(p) and p.kind == "START" and not p.from_}
        for p in self.pieces:
            if not self._live(p) or p.kind not in ("HALF", "MARK"):
                continue
            other = named_to if p.from_ else named_from
            if any(n in other for n in p.names):
                continue
            self.done.add(id(p))
            self.plain.append(p)
