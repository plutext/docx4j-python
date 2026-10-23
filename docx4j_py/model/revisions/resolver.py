"""Accepting and rejecting one change. docx4j ``org.docx4j.model.revisions.Resolver``.

CR-043 section 3.2, row by row; CR-010 section 4. A port of docx4j's
``Resolver`` (itself core-ts's ``TrackedChange.accept`` / ``reject``) onto this
model, with Word's own rules where its saves differ from core-ts's: fields and
content controls go whole, a note goes with its reference, a bookmark touching
removed content goes whole; and the forms core-ts does not reach (math, text
boxes).

Lists here are ``ChildList``\\ s, which set a member's parent when it is put in
and clear it when it is taken out --- so content moving between lists is always
taken out of the one first (CLAUDE.md, CR-010 risk 3).
"""

from __future__ import annotations

from typing import Any

from docx4j_py.child import deep_copy, is_alternate_content, is_any_element
from docx4j_py.model.revisions import props, tables, trees
from docx4j_py.model.revisions.change import (
    CellT,
    FieldT,
    GroupT,
    HolderT,
    MarkPropsT,
    MarkT,
    MathRunT,
    MathStructT,
    Para,
    ParaPropsT,
    RowT,
    RunPropsT,
    RunT,
    SectionT,
    TableT,
    TrackedChange,
)
from docx4j_py.model.revisions.removals import Removals
from docx4j_py.model.revisions.trees import Revision

__all__ = ["join_with_next", "resolve", "resolve_pieces"]


def resolve(change: TrackedChange, accept: bool, removals: Removals) -> None:
    """Accept or reject one change (a group's pieces with it)."""
    _Op(change, removals).run(accept)


def resolve_pieces(pieces: list[TrackedChange], accept: bool, removals: Removals) -> None:
    """The changes of a group, or the cells of a row: runs last first, then marks last first.

    So that a join never strands a piece still to do.
    """
    marks = [piece for piece in pieces if isinstance(piece.target, MarkT)]
    rest = [piece for piece in pieces if not isinstance(piece.target, MarkT)]
    for piece in reversed(rest):
        resolve(piece, accept, removals)
    for piece in reversed(marks):
        resolve(piece, accept, removals)


class _Op:
    """One resolution: the change, and where what it removes is recorded."""

    __slots__ = ("change", "removals", "story")

    def __init__(self, change: TrackedChange, removals: Removals) -> None:
        self.change = change
        self.removals = removals
        self.story = change.story

    def run(self, accept: bool) -> None:
        t = self.change.target
        if not isinstance(t, GroupT):
            self.removals.touched(self.story)
        if isinstance(t, RunT):
            self.run_revision(t, accept)
        elif isinstance(t, MarkT):
            self.mark(t, accept)
        elif isinstance(t, RunPropsT):
            self.run_properties(t, accept)
        elif isinstance(t, ParaPropsT):
            self.paragraph_properties(t, accept)
        elif isinstance(t, MarkPropsT):
            self.mark_properties(t, accept)
        elif isinstance(t, RowT):
            self.row(t, accept)
        elif isinstance(t, GroupT):
            resolve_pieces(t.pieces, accept, self.removals)
        elif isinstance(t, FieldT):
            self.field(t, accept)
        elif isinstance(t, HolderT):
            self.holder(t, accept)
        elif isinstance(t, MathRunT):
            self.math_run(t, accept)
        elif isinstance(t, MathStructT):
            self.math_structure(t, accept)
        elif isinstance(t, TableT):
            if accept:
                tables.accept(t.tbl)
            else:
                tables.reject(t.tbl)
        elif isinstance(t, SectionT):
            if accept:
                t.sect_pr.sect_pr_change = None
            else:
                tables.reject_section(t.sect_pr)
        elif isinstance(t, CellT):
            self.cell(t, accept)

    def cell(self, t: CellT, accept: bool) -> None:
        """An inserted cell stays when accepted, a deleted one goes; rejected the other way round (check 27)."""
        tc_pr = t.tc.tc_pr
        if tc_pr is None:
            return
        if t.ins == accept:
            if tc_pr.cell_ins_or_cell_del_or_cell_merge is t.marker:
                tc_pr.cell_ins_or_cell_del_or_cell_merge = None
        else:
            tables.remove_cell(t.tc, t.tr, self.removals, self.story)

    # -- runs ------------------------------------------------------------------

    def run_revision(self, t: RunT, accept: bool) -> None:
        move = t.revision in (Revision.MOVE_FROM, Revision.MOVE_TO)
        if move and (_Moves.resolve(self, t, accept) or _Moves.resolve_rangeless(self, t, accept)):
            return
        kept = (t.revision in (Revision.INS, Revision.MOVE_TO)) == accept
        if kept:
            if t.revision in (Revision.DEL, Revision.MOVE_FROM):
                self.restore_deleted(t.item, t.owner)
            else:
                self.unwrap(t.item, t.owner)
        else:
            self.remove_revision(t.item, t.owner)

    def mark(self, t: MarkT, accept: bool) -> None:
        para = self.change.para
        if para is None:
            return
        self.drop_mark(para, t.ins)
        if t.ins != accept:
            self.join_with_next(para)

    def run_properties(self, t: RunPropsT, accept: bool) -> None:
        run = t.run
        if accept:
            if run.r_pr is not None:
                run.r_pr.r_pr_change = None
            props.prune_run(run)
        else:
            props.restore_r_pr(run, t.record)

    def paragraph_properties(self, t: ParaPropsT, accept: bool) -> None:
        para = self.change.para
        if para is None:
            return
        p = para.p
        if accept:
            if p.p_pr is not None:
                p.p_pr.p_pr_change = None
            props.prune_paragraph(p)
        else:
            props.restore_p_pr(p, t.record.p_pr, self.story.package)

    def mark_properties(self, t: MarkPropsT, accept: bool) -> None:
        para = self.change.para
        if para is None:
            return
        p = para.p
        if accept:
            if p.p_pr is not None and p.p_pr.r_pr is not None:
                p.p_pr.r_pr.r_pr_change = None
            props.prune_paragraph(p)
        else:
            props.restore_mark_properties(p, t.record)

    # -- rows ------------------------------------------------------------------

    def row(self, t: RowT, accept: bool) -> None:
        """An accepted insertion keeps the row, its content's own revisions going with the marker.

        An accepted deletion takes the row and everything in it away; a table
        left with no rows goes (check 30, C13, C25), but only when no branch of
        a row-level ``mc:AlternateContent`` has a row left (docx4j
        ``Resolver.row`` over ``Lister.hasRows``, its CR-051 phase 2).
        """
        if t.ins == accept:
            tr_pr = t.tr.tr_pr
            if tr_pr is not None:
                if t.ins:
                    tr_pr.ins = None
                else:
                    tr_pr.del_value = None
                if trees.is_empty(tr_pr):
                    t.tr.tr_pr = None
            for inner in reversed(t.inner):
                resolve(inner, accept, self.removals)
        else:
            self.remove(t.owner, t.tr)
            if not tables.has_rows(t.tbl):  # no branch left with a row
                self.remove(t.tbl_owner, t.tbl)

    # -- fields and controls ---------------------------------------------------

    def field(self, t: FieldT, accept: bool) -> None:
        """Word's inserted field: accepted, its ``w:ins`` pieces unwrapped; rejected, the whole field goes.

        Its separator and result too (C01, C16), last first; a bookmark, a
        comment's range or a proofing mark between its pieces stays where it
        stood.
        """
        r_class = trees.cls("R")
        for element in reversed(t.span):
            owner = self.current_owner(t.owner, element)
            if trees.revision_of(element) is Revision.INS:
                if accept:
                    self.unwrap(element, owner)
                else:
                    self.remove_revision(element, owner)
            elif not accept and isinstance(element, r_class):
                self.remove(owner, element)

    def holder(self, t: HolderT, accept: bool) -> None:
        """A content control or custom XML element inserted or deleted.

        Inserted (check 30, C16b): accepted, its four markers go; rejected, the
        element goes and what is left of its content stands in its place ---
        nothing, once the insertions inside it have been rejected with it.
        Deleted (docx4j-core-ts check 54 cases 15 and 17; CR-013 Phase C):
        accepted, the element goes with what it holds, as Word's Accept All
        took it; rejected, its four markers go.
        """
        self.remove(t.owner, t.outer_start)
        self.remove(t.owner, t.outer_end)
        items = trees.children(t.holder)
        if items is not None:
            _remove_quietly(items, t.inner_end)
            _remove_quietly(items, t.inner_start)
        if not t.ins:
            if accept:
                self.remove(t.owner, t.holder)
            return
        if accept:
            return
        owner = self.current_owner(t.owner, t.holder)
        index = trees.index_of(owner, t.holder)
        if index < 0:
            return
        moved = list(items) if items is not None else []
        if items is not None:
            items.clear()
        trees.splice(owner, index, 1, moved)

    # -- math ------------------------------------------------------------------

    def math_run(self, t: MathRunT, accept: bool) -> None:
        """``w:ins`` / ``w:del`` inside an ``m:r``: kept, its content is unwrapped into the run.

        Not kept, it goes, and a run left with no text goes with it (CR-043
        section 3.2).
        """
        run = t.run
        content = run.content
        index = trees.index_of(content, t.item)
        if index < 0:
            return
        if t.ins == accept:
            item = t.item
            moved = list(item.content)
            item.content.clear()
            if not t.ins:
                moved = [_restored_text(entry) for entry in moved]
            # the w:ins's own m:rPr and w:rPr become the run's, where it has none
            math_r_pr = getattr(item, "r_pr", None)
            word_r_pr = getattr(item, "schemas_openxmlformats_org_wordprocessingml_2006_main_r_pr", None)
            if math_r_pr is not None and run.r_pr is None:
                item.r_pr = None
                run.r_pr = math_r_pr
                math_r_pr.parent = run
            if word_r_pr is not None and run.r_pr_1 is None:
                item.schemas_openxmlformats_org_wordprocessingml_2006_main_r_pr = None
                run.r_pr_1 = word_r_pr
                word_r_pr.parent = run
            del content[index]
            trees.splice(content, index, 0, moved)
        else:
            self.removals.discard(self.story, t.item)
            del content[index]
            track = trees.cls("CTMathRunTrackChange")
            has_text = any(
                trees.item_text(entry) is not None or isinstance(entry, track) for entry in content
            )
            if not has_text and t.run_owner is not None:
                holder = getattr(run, "parent", None)
                self.remove(t.run_owner, run)
                self.drop_emptied_equation(holder)

    def math_structure(self, t: MathStructT, accept: bool) -> None:
        """An equation structure inserted or deleted (its ``m:ctrlPr``'s ``w:ins`` / ``w:del``).

        Kept, the marker goes and its recorded run properties become the
        control's; not kept, the structure goes whole, from the list that holds
        it, as Word's Accept All and Reject All did (an inserted ``m:sSup``, a
        deleted ``m:acc``: ``tracked-changes-equations.docx``).
        """
        if t.ins == accept or t.struct_item is None:
            ctrl_pr = t.ctrl_pr
            marker = ctrl_pr.r_pr_or_ins_or_del
            if marker is t.marker:
                live = props.r_pr_of(getattr(marker, "r_pr", None))
                ctrl_pr.r_pr_or_ins_or_del = live
                if live is not None:
                    from docx4j_py.child import link_parents

                    link_parents(live)
                    live.parent = ctrl_pr
        else:
            holder = getattr(t.struct_item, "parent", None)
            self.remove(t.struct_owner, t.struct_item)
            self.drop_emptied_equation(holder)

    def drop_emptied_equation(self, holder: Any) -> None:
        """An ``m:oMath`` left with nothing in it goes, and an ``m:oMathPara`` left with no ``m:oMath`` with it.

        docx4j ``Resolver.dropEmptiedEquation`` (``ebb21c990``, on
        ``VERSION_17_3_5``; its CR-043 section 16), after a math run or a
        structure is removed: Word 365's Accept All of a whole equation deleted
        in Word's form (a ``w:del`` in each ``m:r`` and each ``m:ctrlPr``) removed
        both, where docx4j left ``<m:oMathPara><m:oMath/></m:oMathPara>``
        (docx4j-core-ts check 57, 23 October 2026); by symmetry Reject All of a
        whole equation inserted (unmeasured). An argument (``m:e``, ``m:sup``)
        left empty stays: its structure keeps its places. CR-013 section 13.7, F1.
        """
        from docx4j_py.math import CTOMath, CTOMathPara

        if not isinstance(holder, CTOMath) or holder.content:
            return
        above = getattr(holder, "parent", None)
        if isinstance(above, CTOMathPara):
            index = trees.index_of(above.o_math, holder)
            if index >= 0:
                self.removals.discard(self.story, holder)
                del above.o_math[index]
            if above.o_math:
                return
            holder = above
            above = getattr(above, "parent", None)
        items = trees.list_of(above) if above is not None else None
        if items is None or trees.index_of(items, holder) < 0:
            return
        self.remove(items, holder)

    # -- the operations on lists -------------------------------------------------

    def current_owner(self, owner: list | None, item: Any) -> list | None:
        """The list `item` is in now: `owner` as it was listed, else the one its parent keeps."""
        if trees.index_of(owner, item) >= 0:
            return owner
        actual = trees.list_of(getattr(item, "parent", None))
        if actual is not None and trees.index_of(actual, item) >= 0:
            return actual
        return owner

    def remove(self, owner: list | None, item: Any) -> None:
        """Takes `item` out of its list, recording what it took with it."""
        if owner is None:
            return
        items = self.current_owner(owner, item)
        index = trees.index_of(items, item)
        if index < 0:
            return
        self.removals.discard(self.story, item)
        del items[index]
        if isinstance(item, trees.cls("R")) or trees.revision_of(item) is not None:
            self.drop_emptied_bookmarks(items, index)

    def drop_emptied_bookmarks(self, items: list, at: int) -> None:
        """A bookmark around content that went, left with nothing in it, goes too (C11b).

        Only a bookmark whose start stood before the content and whose end after
        it: one that was empty already stays.
        """
        start = at - 1
        while start >= 0:
            candidate = items[start]
            if _has_text(candidate):
                break
            if trees.name(candidate) == "bookmarkStart":
                mark_id = trees.marker_id(candidate)
                for end in range(at, len(items)):
                    after = items[end]
                    if trees.name(after) == "bookmarkEnd" and trees.marker_id(after) == mark_id:
                        del items[end]
                        del items[start]
                        at -= 1
                        break
                    if _has_text(after):
                        break
            start -= 1

    def unwrap(self, item: Any, owner: list | None) -> None:
        """A ``w:ins`` or ``w:moveTo`` kept: its content takes its place."""
        items = self.current_owner(owner, item)
        index = trees.index_of(items, item)
        if index < 0:
            return
        inner = trees.run_items(item)
        moved = list(inner) if inner is not None else []
        if inner is not None:
            inner.clear()
        del items[index]
        trees.splice(items, index, 0, moved)

    def restore_deleted(self, item: Any, owner: list | None) -> None:
        """A ``w:del`` or ``w:moveFrom`` kept: its runs come back, deleted text become text again."""
        restore_text(trees.run_items(item))
        self.unwrap(item, owner)

    def remove_revision(self, item: Any, owner: list | None) -> None:
        """A revision taken away with its content (an insertion rejected, a deletion accepted), and its comments.

        docx4j ``Resolver.removeRevision`` (``d18693760`` and ``a8bc6cda9``, its
        CR-043 section 15): a comment whose range start was inside, or whose
        markers lie either side of it in its list, goes whole with the text it
        was on --- markers, reference and entry --- once nothing with text is
        left in its range. That is Word 365's rule: core-ts check 32, and Word
        365's Reject All of the editor's ``review-for-word`` and of this
        package's acceptance artefact 8, 21 October 2026 (CR-010 section 12.8,
        fix 2). Word 2010 and Word 15 kept the comment whose markers lie
        outside, losing only its range markers; this package no longer does.

        A comment with one marker in this list on the revision's side and the
        other in another paragraph --- a start before the revision with no end
        later in the list, an end after it with no start earlier --- and one
        whose start was inside with its end elsewhere, is judged when the
        resolution is flushed (:meth:`Removals.remove_comment_if_emptied`),
        once the paragraph marks taken away have joined what is left: Word
        365's Reject All of two inserted paragraphs commented whole removes the
        comment (CR-043 U23; CR-010 section 12.9).
        """
        items = self.current_owner(owner, item)
        index = trees.index_of(items, item)
        if index < 0:
            return
        inner = trees.run_items(item)
        kept: list[Any] = []
        if inner is not None:
            kept = [entry for entry in inner if trees.name(entry) in ("commentRangeStart", "commentRangeEnd")]
        inside = [trees.marker_id(entry) for entry in kept if trees.name(entry) == "commentRangeStart"]
        outside: list[int | None] = []
        for s in range(index):
            if trees.name(items[s]) != "commentRangeStart":
                continue
            comment_id = trees.marker_id(items[s])
            if marker_at(items, s + 1, len(items), "commentRangeEnd", comment_id) < 0:
                outside.append(comment_id)  # its end in another paragraph
            elif marker_at(items, index + 1, len(items), "commentRangeEnd", comment_id) >= 0:
                outside.append(comment_id)
        for e in range(index + 1, len(items)):
            if trees.name(items[e]) != "commentRangeEnd":
                continue
            comment_id = trees.marker_id(items[e])
            if marker_at(items, 0, e, "commentRangeStart", comment_id) < 0:
                outside.append(comment_id)  # its start in another paragraph
        if inner is not None:
            for entry in kept:
                position = trees.index_of(inner, entry)
                if position >= 0:
                    del inner[position]
        self.removals.discard(self.story, item)
        del items[index]
        trees.splice(items, index, 0, kept)
        self.drop_emptied_bookmarks(items, index)
        for comment_id in dict.fromkeys([*inside, *outside]):
            if comment_id is None:
                continue
            start = marker_at(items, 0, len(items), "commentRangeStart", comment_id)
            if start < 0 or marker_at(items, start + 1, len(items), "commentRangeEnd", comment_id) < 0:
                self.removals.remove_comment_if_emptied(comment_id)  # start and end in different paragraphs (U23)
            elif range_is_empty(items, start, comment_id):
                self.removals.remove_comment(comment_id)

    # -- paragraph marks ---------------------------------------------------------

    def drop_mark(self, para: Para, ins: bool) -> None:
        p_pr = para.p.p_pr
        if p_pr is None or p_pr.r_pr is None:
            return
        if ins:
            p_pr.r_pr.ins = None
        else:
            p_pr.r_pr.del_value = None
        props.prune_paragraph(para.p)

    def join_with_next(self, para: Para) -> None:
        """A paragraph mark taken away: the paragraph joins the next one in its list (:func:`join_with_next`)."""
        join_with_next(para, self.removals, self.story, self.current_owner)

    # -- helpers ---------------------------------------------------------------


def join_with_next(
    para: Para, removals: Removals, story: Any, current_owner: Any = None
) -> Any:
    """A paragraph mark taken away: the paragraph joins the next one in its list.

    The result keeps the next paragraph's properties and ``w14:paraId``,
    for markup Word did not write as well (checks 18, 22, 33; C10, C10b).
    Range markers between the two (a bookmark's end, a comment's) go into
    the joined paragraph where they stood. With no paragraph next in the
    same list --- a story's, a cell's, a control's last, or a table or a
    control next --- a join never crosses a container's boundary (U1) but
    for one into a table's first cell, and Word 365's Accept All does one of
    four things instead (:func:`_no_paragraph_next`; docx4j CR-043 section
    17, CR-010 sections 13 and 13.3a). Otherwise the mark is only dropped.

    A module function since CR-013 Phase C, so that a ``Range``'s untracked
    deletion across paragraphs joins as Accept All joins, as docx4j-core-ts's
    ``spanEdits`` reuses its ``joinWithNext``. `current_owner` is the
    resolution's own lookup of the list an item is in now; `story` is what
    `removals` records the first paragraph's properties against (a section
    break's header and footer references, C06). Returns the paragraph that
    holds the join (the next one, or a table's first cell's), `para`'s own
    ``w:p`` where nothing was joined, or None where the paragraph went.
    """
    container = current_owner(para.container, para.p) if current_owner is not None else para.container
    index = trees.index_of(container, para.p)
    if index < 0:
        return None
    j = index + 1
    while j < len(container) and _is_marker(container[j]):
        j += 1
    following = container[j] if j < len(container) else None
    if not isinstance(following, trees.cls("P")):
        return _no_paragraph_next(para.p, container, index, j, following, removals, story)
    # the properties the first paragraph loses: a section break's header and
    # footer references among them (C06), whose parts may now be unreferenced (D7)
    if para.p.p_pr is not None:
        removals.discard(story, para.p.p_pr)
    moved = list(para.p.content)
    para.p.content.clear()
    markers = [container[k] for k in range(index + 1, j)]
    for _ in markers:
        del container[index + 1]
    moved.extend(markers)
    trees.splice(following.content, 0, 0, moved)
    del container[index]
    return following


def _no_paragraph_next(
    p: Any, container: list, index: int, j: int, following: Any, removals: Removals, story: Any
) -> Any:
    """A mark taken away with no paragraph next in its list: Word 365's Accept All, else the mark dropped.

    docx4j ``Resolver.joinWithNext``'s branch with no paragraph next, as its
    CR-043 section 17 has it (``29dbf5964``, revised by ``e55a83b8e``; CR-010
    sections 13 and 13.3a), in docx4j's order, from Word 365's Accept All of
    docx4j-core-ts's check 54 and of docx4j's text-left probe:

    1. emptied --- nothing but range markers left --- with a table or a block
       content control (or custom XML element) next: the paragraph goes, and
       what is next follows what came before (check 54 case 9, the probe's
       case 3);
    2. text left, with a table next: the content joins the table's first
       cell's first paragraph, at its front, which keeps its own properties,
       the range markers between going with it, and the paragraph goes (the
       probe's case 1);
    3. text left, with a block control next: the paragraph takes a copy of
       the properties of the control's first paragraph, as a join would have
       given it, and nothing joins (check 54 case 19, the probe's case 5);
    4. the last in a control's content, text left or not: a bare mark, its
       properties gone (check 54 case 13, the probe's case 2).

    Whichever resolution took the mark away, a deletion accepted or an
    insertion rejected, as docx4j's code has it; Word's Reject All of check 54
    and of the probe left nothing these rules reach. A kept block-level
    ``mc:AlternateContent`` next is read as the branch Word takes reads
    (:func:`_as_word_reads`; CR-012 Phase A), so the first block in that
    branch decides rules 1 and 3; no content joins into a kept branch, so rule
    2 does not reach one. docx4j reads no further than the element.

    Otherwise, as before (docx4j's U1), the mark is only dropped: a story's or
    a cell's last paragraph (the probe's case 4, where Word collapsed a
    one-cell table, is not followed, as docx4j does not follow it), and a
    paragraph before a table or control that holds no paragraph first.
    Returns the paragraph left holding the place: `p`, the cell's paragraph
    `p` joined, or None where `p` went.
    """
    emptied = not any(not _is_marker(item) for item in p.content)
    block = _as_word_reads(following)
    tbl_class = trees.cls("Tbl")
    if emptied and (isinstance(block, tbl_class) or (block is not None and trees.is_holder(block))):
        # check 54 case 9, the probe's case 3: gone, the table or control follows what came before
        if p.p_pr is not None:
            removals.discard(story, p.p_pr)
        del container[index]
        return None
    if isinstance(following, tbl_class):
        first = _first_cell_paragraph(following)
        if first is not None:
            # the probe's case 1: its content joins the first cell's first paragraph, which keeps its properties
            if p.p_pr is not None:
                removals.discard(story, p.p_pr)
            moved = list(p.content)
            p.content.clear()
            markers = [container[k] for k in range(index + 1, j)]
            for _ in markers:
                del container[index + 1]
            moved.extend(markers)
            trees.splice(first.content, 0, 0, moved)
            del container[index]
            return first
    if block is not None and trees.is_holder(block):
        first = _first_paragraph(trees.children(block))
        if first is not None:
            # check 54 case 19, the probe's case 5: the properties a join would have given it, the control's first paragraph's
            if p.p_pr is not None:
                removals.discard(story, p.p_pr)
            p.p_pr = deep_copy(first.p_pr, p) if first.p_pr is not None else None
            return p
    if following is None and isinstance(getattr(p, "parent", None), trees.cls("SdtContentBlock")):
        # check 54 case 13, the probe's case 2: a control's last paragraph keeps a bare mark, its text with it
        if p.p_pr is not None:
            removals.discard(story, p.p_pr)
        p.p_pr = None
        return p
    r_pr = p.p_pr.r_pr if p.p_pr is not None else None
    if r_pr is not None:
        r_pr.ins = None
        r_pr.del_value = None
    props.prune_paragraph(p)
    return p


def _as_word_reads(item: Any) -> Any:
    """A block as Word reads it: a kept block-level ``mc:AlternateContent`` as the first block of the branch it takes.

    Past range markers; None for an empty branch, and at the end of a list
    (CR-012 Phase A: the readers read Word's branch at block level). Any
    other block is itself.
    """
    if item is None or not is_alternate_content(item):
        return item
    for entry in trees.branch_items(item) or ():
        if not _is_marker(entry):
            return entry
    return None


def _first_cell_paragraph(tbl: Any) -> Any:
    """The first paragraph of a table's first cell, where a paragraph joining the table lands; None where it has none.

    docx4j ``Resolver.firstCellParagraph`` (``e55a83b8e``): the first row's
    first cell, and in it the first paragraph before any table.
    """
    tr_class, tc_class = trees.cls("Tr"), trees.cls("Tc")
    for row in tbl.content:
        if not isinstance(row, tr_class):
            continue
        for cell in row.content:
            if isinstance(cell, tc_class):
                return _first_paragraph(cell.content)
        return None
    return None


def _first_paragraph(items: list | None) -> Any:
    """The first paragraph of a list, before any table; None where a table comes first (docx4j ``Resolver.firstParagraph``)."""
    p_class, tbl_class = trees.cls("P"), trees.cls("Tbl")
    for item in items or ():
        if isinstance(item, p_class):
            return item
        if isinstance(item, tbl_class):
            return None
    return None


def _is_marker(item: Any) -> bool:
    local = trees.name(item)
    return local.endswith(("RangeStart", "RangeEnd")) or local in (
        "bookmarkStart",
        "bookmarkEnd",
        "permStart",
        "permEnd",
        "proofErr",
    )


def _remove_quietly(items: list, item: Any) -> None:
    index = trees.index_of(items, item)
    if index >= 0:
        del items[index]


def _has_text(item: Any) -> bool:
    """True when run-level content shows text in the accepted view.

    A wildcard node too, by element name, for a comment's range in a text box,
    which is a wildcard tree again once its island is written back
    (:func:`range_is_empty` at the flush). Since CR-012 Phase C a VML text box
    is typed and makes no island; one below a wildcard node still does.
    """
    if is_any_element(item):
        return _wildcard_has_text(item)
    if isinstance(item, trees.cls("R")):
        return trees.has_text(item)
    revision = trees.revision_of(item)
    if revision in (Revision.DEL, Revision.MOVE_FROM):
        return False
    items = trees.run_items(item)
    if items is None:
        return False
    return any(_has_text(entry) for entry in items)


#: The WordprocessingML run holders a wildcard paragraph's text can stand in.
_WILDCARD_HOLDERS = frozenset(
    {"p", "ins", "moveTo", "hyperlink", "sdt", "sdtContent", "smartTag", "customXml", "fldSimple", "dir", "bdo"}
)


def _wildcard_has_text(item: Any) -> bool:
    """:func:`_has_text` of a wildcard node: a ``w:r`` with text, or a run holder holding one."""
    qname = item.qname
    if not isinstance(qname, str) or not qname.startswith(f"{{{trees.W_NS}}}"):
        return False
    local = qname.rpartition("}")[2]
    children = [child for child in item.children or () if not isinstance(child, str)]
    if local == "r":
        return any(trees.item_text(child) is not None for child in children)
    if local not in _WILDCARD_HOLDERS:
        return False
    return any(_wildcard_has_text(child) for child in children if is_any_element(child))


def marker_at(items: list, start: int, stop: int, local: str, comment_id: int | None) -> int:
    """Where in ``items[start:stop]`` a comment's range start or end (`local`) stands, or -1 (docx4j ``markerAt``)."""
    for index in range(start, stop):
        entry = items[index]
        if not isinstance(entry, str) and trees.name(entry) == local and trees.marker_id(entry) == comment_id:
            return index
    return -1


def range_is_empty(items: list, start: int, comment_id: int | None) -> bool:
    """True when nothing with text stands between a comment's range start, at `start`, and its end in the same list.

    False when text does, or the end is elsewhere (docx4j ``rangeIsEmpty``).
    """
    for index in range(start + 1, len(items)):
        entry = items[index]
        if isinstance(entry, str):
            continue
        if trees.name(entry) == "commentRangeEnd" and trees.marker_id(entry) == comment_id:
            return True
        if _has_text(entry):
            return False
    return False  # the end is elsewhere: text may be


# ---------------------------------------------------------------------------
# w:t <-> w:delText
# ---------------------------------------------------------------------------


def _restored_text(item: Any) -> Any:
    """``w:delText`` as ``w:t``, ``w:delInstrText`` as ``w:instrText``; anything else as it is."""
    local = trees.name(item)
    if local not in ("delText", "delInstrText"):
        return item
    from docx4j_py.wml import el

    made = getattr(el, "t" if local == "delText" else "instrText")(getattr(item, "value", "") or "")
    space = getattr(item, "space", None)
    if space is not None:
        made.space = space
    return made


def _deleted_text(item: Any) -> Any:
    """``w:t`` as ``w:delText``, ``w:instrText`` as ``w:delInstrText``."""
    local = trees.name(item)
    if local not in ("t", "instrText") or trees.is_math(item):
        return item
    from docx4j_py.wml import el

    made = getattr(el, "delText" if local == "t" else "delInstrText")(getattr(item, "value", "") or "")
    space = getattr(item, "space", None)
    if space is not None:
        made.space = space
    return made


def restore_text(items: list | None) -> None:
    """Deleted text become text again in every run under `items`."""
    _rename_text(items, _restored_text)


def delete_text(items: list | None) -> None:
    """Text become deleted text in every run under `items`."""
    _rename_text(items, _deleted_text)


def _rename_text(items: list | None, rename: Any) -> None:
    if not items:
        return
    r_class = trees.cls("R")
    for item in items:
        if isinstance(item, r_class):
            content = item.content
            for index, entry in enumerate(list(content)):
                renamed = rename(entry)
                if renamed is not entry:
                    content[index] = renamed
        elif trees.revision_of(item) is None:
            _rename_text(trees.run_items(item), rename)


# ---------------------------------------------------------------------------
# moves
# ---------------------------------------------------------------------------


class _Piece:
    """One piece of move markup: a range start or end, a half, a moved mark, or what stands inside a range."""

    __slots__ = ("from_", "id", "item", "kind", "name", "names", "owner", "para")

    def __init__(self, kind: str, from_: bool) -> None:
        self.kind = kind
        self.from_ = from_
        self.id: int | None = None
        self.name: str | None = None
        self.item: Any = None
        self.owner: list | None = None
        self.names: list[str] = []
        self.para: Para | None = None


class _Moves:
    """Moves (core-ts ``resolveMove``, check 18 and check 20).

    Word pairs the halves by the ``w:name`` on their range starts; a half
    belongs to the move whose range of its side is open where it stands.
    Either half accepts or rejects the whole move.
    """

    @staticmethod
    def pieces_of(story_root: list) -> list[_Piece]:
        """Every piece of move markup in a story, in document order."""
        out: list[_Piece] = []
        _Moves._visit(story_root, False, {}, {}, out)
        return out

    @staticmethod
    def _visit(items: list, run_level: bool, open_from: dict, open_to: dict, out: list[_Piece]) -> None:
        p_class = trees.cls("P")
        for item in items:
            local = trees.name(item)
            if local.startswith(("moveFromRange", "moveToRange")):
                from_ = local.startswith("moveFrom")
                start = local.endswith("Start")
                piece = _Piece("START" if start else "END", from_)
                piece.id = trees.marker_id(item)
                piece.item = item
                piece.owner = items
                if start:
                    piece.name = getattr(item, "name", None) or ""
                    (open_from if from_ else open_to)[piece.id] = piece.name
                else:
                    (open_from if from_ else open_to).pop(piece.id, None)
                out.append(piece)
                continue
            revision = trees.revision_of(item)
            if revision in (Revision.MOVE_FROM, Revision.MOVE_TO):
                from_ = revision is Revision.MOVE_FROM
                piece = _Piece("HALF", from_)
                piece.item = item
                piece.owner = items
                piece.names = list((open_from if from_ else open_to).values())
                out.append(piece)
            elif run_level:
                for from_ in (True, False):
                    opened = open_from if from_ else open_to
                    if not opened:
                        continue
                    piece = _Piece("INNER", from_)
                    piece.item = item
                    piece.owner = items
                    piece.names = list(opened.values())
                    out.append(piece)
            nested = trees.children(item)
            if nested is not None:
                _Moves._visit(nested, isinstance(item, p_class), open_from, open_to, out)
            if isinstance(item, p_class):
                p_pr = item.p_pr
                r_pr = p_pr.r_pr if p_pr is not None else None
                for from_ in (True, False):
                    if r_pr is None or (r_pr.move_from if from_ else r_pr.move_to) is None:
                        continue
                    piece = _Piece("MARK", from_)
                    piece.para = Para(item, items)
                    piece.names = list((open_from if from_ else open_to).values())
                    out.append(piece)

    @staticmethod
    def _half_of(pieces: list[_Piece], item: Any) -> _Piece | None:
        for piece in pieces:
            if piece.kind == "HALF" and piece.item is item:
                return piece
        return None

    @staticmethod
    def resolve(op: _Op, t: RunT, accept: bool) -> bool:
        """Accepting keeps the destination and removes the source, paragraph marks and all.

        Rejecting the other way round, taking with it what was typed into the
        destination (check 20's "keep original location text"); the four range
        markers go either way. A destination whose source is gone is kept either
        way (check 20 case 12). False for a half in no named range.
        """
        pieces = _Moves.pieces_of(op.story.root)
        half = _Moves._half_of(pieces, t.item)
        if half is None or not half.names:
            return False
        name = half.names[0]
        ids = {piece.id for piece in pieces if piece.kind == "START" and piece.name == name}
        mine = [
            piece
            for piece in pieces
            if (piece.kind == "START" and piece.name == name)
            or (piece.kind == "END" and piece.id in ids)
            or (piece.kind in ("HALF", "MARK", "INNER") and name in piece.names)
        ]
        any_from = any(piece.from_ for piece in mine)
        any_to = any(not piece.from_ for piece in mine)
        orphan = not any_from
        kept_from = not (accept or orphan)
        broken = orphan or not any_to

        # the markers first: a range end between two paragraphs would stop a join below
        for piece in mine:
            if piece.kind in ("START", "END"):
                owner = op.current_owner(piece.owner, piece.item)
                if owner is not None:
                    _remove_quietly(owner, piece.item)
        if kept_from:
            for piece in mine:
                if piece.kind == "INNER" and not piece.from_:
                    op.remove(piece.owner, piece.item)
        for piece in mine:
            if piece.kind != "HALF":
                continue
            if piece.from_ == kept_from:
                if piece.from_:
                    op.restore_deleted(piece.item, piece.owner)
                else:
                    op.unwrap(piece.item, piece.owner)
            else:
                op.remove(piece.owner, piece.item)
        # the paragraph marks last, last first, so that a join never disturbs one still to do
        marks = [piece for piece in mine if piece.kind == "MARK"]
        for piece in reversed(marks):
            _Moves._resolve_mark(op, piece, piece.from_ == kept_from)
        if broken:
            _Moves.dissolve_broken(op)
        return True

    @staticmethod
    def _resolve_mark(op: _Op, mark: _Piece, kept: bool) -> None:
        """A moved paragraph's mark: kept, the marker goes; not kept, the paragraph goes when empty, else joins the next."""
        para = mark.para
        p = para.p
        r_pr = p.p_pr.r_pr if p.p_pr is not None else None
        if r_pr is not None:
            if mark.from_:
                r_pr.move_from = None
            else:
                r_pr.move_to = None
        if kept:
            props.prune_paragraph(p)
        elif not p.content:
            op.remove(para.container, p)  # the moved paragraph, emptied: gone, as in Word
        else:
            op.join_with_next(para)

    @staticmethod
    def resolve_rangeless(op: _Op, t: RunT, accept: bool) -> bool:
        """A move half in no named range (check 20 cases 13 and 14).

        A ``w:moveTo`` resolved as an insertion, a ``w:moveFrom`` as a deletion,
        its paragraph's mark with it; then what is left of the broken pair is
        dissolved.
        """
        pieces = _Moves.pieces_of(op.story.root)
        half = _Moves._half_of(pieces, t.item)
        if half is None or half.names:
            return False
        keep = (not half.from_) == accept
        if not keep:
            op.remove(t.owner, t.item)
        elif not half.from_:
            op.unwrap(t.item, t.owner)
        else:
            op.restore_deleted(t.item, t.owner)
        para = op.change.para
        if para is not None:
            for piece in pieces:
                if piece.kind != "MARK" or piece.from_ != half.from_ or piece.para.p is not para.p or piece.names:
                    continue
                _Moves._resolve_mark(op, piece, keep)
                break
        _Moves.dissolve_broken(op)
        return True

    @staticmethod
    def dissolve_broken(op: _Op) -> None:
        """Every move half and moved mark that cannot be paired, turned into the plain ``w:ins`` or ``w:del`` it stands for.

        As Word saved each broken pair in check 20.
        """
        pieces = _Moves.pieces_of(op.story.root)
        named_from = {piece.name for piece in pieces if piece.kind == "START" and piece.from_}
        named_to = {piece.name for piece in pieces if piece.kind == "START" and not piece.from_}
        for piece in pieces:
            other = named_to if piece.from_ else named_from
            if any(name in other for name in piece.names):
                continue
            if piece.kind == "HALF":
                _to_plain_revision(piece.item, piece.owner, piece.from_)
            elif piece.kind == "MARK":
                r_pr = piece.para.p.p_pr.r_pr
                if piece.from_ and r_pr.move_from is not None:
                    r_pr.del_value = r_pr.move_from
                    r_pr.move_from = None
                if not piece.from_ and r_pr.move_to is not None:
                    r_pr.ins = r_pr.move_to
                    r_pr.move_to = None


def _to_plain_revision(item: Any, owner: list | None, from_: bool) -> None:
    """A ``w:moveFrom`` become the plain ``w:del`` of the same runs, or a ``w:moveTo`` the plain ``w:ins``; it keeps its id."""
    index = trees.index_of(owner, item)
    if index < 0:
        return
    from docx4j_py.wml import RunDel, RunIns

    inner = trees.run_items(item) or []
    moved = list(inner)
    inner.clear()
    if from_:
        delete_text(moved)
    plain = (RunDel if from_ else RunIns)(id=item.id, author=item.author, date=item.date, date_utc=item.date_utc)
    trees.splice(plain.content, 0, 0, moved)
    owner[index] = trees.fit(plain, owner)
