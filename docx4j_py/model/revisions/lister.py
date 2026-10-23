"""Lists a package's tracked changes. docx4j ``org.docx4j.model.revisions.Lister``.

CR-010 section 3.1; CR-043 section 3.1. Each story walked in document order,
its single changes ("pieces") grouped as Office JS groups them: pieces that
touch, of one type, by one author, whatever their dates, are one change; a
different author or type splits; a revision nested in another is listed apart;
moves are never grouped; a group's date is its first piece's. A port of
docx4j's ``Lister`` (itself core-ts's ``Body.getTrackedChanges`` and
``TrackedChange.mts``'s collecting half), extended to every story part, text
boxes in both ``mc:AlternateContent`` branches (as typed islands, CR-010 section
3.5) and math.
"""

from __future__ import annotations

import dataclasses
import enum
import re
from typing import Any

from docx4j_py.child import (
    MC_CHOICE,
    MC_FALLBACK,
    is_alternate_content,
    is_any_element,
    iter_children,
    mce_branch,
)
from docx4j_py.model.revisions import tables, trees
from docx4j_py.model.revisions.change import (
    CellT,
    FieldT,
    GroupKind,
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
from docx4j_py.model.revisions.islands import Island, holds_revisions, is_wildcard_txbx
from docx4j_py.model.revisions.story import Story
from docx4j_py.model.revisions.trees import Revision, View

__all__ = ["BREAK", "Lister", "box_refs", "groupable", "package_stories", "story_parts"]


class _Break:
    __slots__ = ()

    def __repr__(self) -> str:
        return "BREAK"


#: Between two changes: something unrevised that keeps them apart.
BREAK = _Break()

#: The run children a text box can be inside.
_BOX_HOLDERS = frozenset({"drawing", "pict", "object", "AlternateContent"})

#: Revision markup, by local name, in a part's bytes: a part with none is not
#: unmarshalled for the listing, so that it is still written back byte for byte.
_REVISION_BYTES = re.compile(
    rb"<(?:[A-Za-z_][\w.-]*:)?(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange"
    rb"|tblPrChange|tblGridChange|tblPrExChange|trPrChange|tcPrChange|cellIns|cellDel|cellMerge"
    rb"|customXmlInsRangeStart|customXmlInsRangeEnd|customXmlDelRangeStart|customXmlDelRangeEnd"
    rb"|numberingChange)[\s/>]"
)


# ---------------------------------------------------------------------------
# the stories
# ---------------------------------------------------------------------------


def _may_hold_revisions(part: Any) -> bool:
    """True when a part may hold revision markup: always once unmarshalled; else its bytes say."""
    if getattr(part, "is_unmarshalled", True):
        return True
    try:
        data = part._raw_bytes()
    except Exception:  # noqa: BLE001 - no bytes to read: let the walk decide
        return True
    return _REVISION_BYTES.search(data) is not None


def _headers_and_footers(main: Any) -> list[Any]:
    """The headers and footers in the order the sections reference them, then any the main part relates to."""
    from docx4j_py.openpackaging.parts.wml import FooterPart, HeaderPart

    rels = main.relationships_part
    if rels is None:
        return []
    out: dict[int, Any] = {}
    for sect_pr in _sections(main.contents.body):
        for ref in sect_pr.header_reference_or_footer_reference or ():
            rel_id = getattr(ref, "id", None)
            part = rels.get_part(rel_id) if rel_id else None
            if isinstance(part, (HeaderPart, FooterPart)):
                out.setdefault(id(part), part)
    for part in (*main.header_parts(), *main.footer_parts()):
        out.setdefault(id(part), part)
    return list(out.values())


def _sections(body: Any) -> list[Any]:
    """Every ``w:sectPr`` of a body in document order: the paragraphs' and the body's own."""
    out: list[Any] = []
    if body is None:
        return out
    p_class = trees.cls("P")

    def visit(items: list) -> None:
        for item in items:
            if isinstance(item, p_class):
                p_pr = item.p_pr
                if p_pr is not None and p_pr.sect_pr is not None:
                    out.append(p_pr.sect_pr)
                continue
            nested = trees.block_items(item)
            if nested is not None and not isinstance(item, trees.cls("Tbl")):
                visit(nested)

    visit(body.content)
    if body.sect_pr is not None:
        out.append(body.sect_pr)
    return out


def story_parts(package: Any) -> list[Any]:
    """Every part whose content is a story's: main, headers, footers, footnotes, endnotes, comments."""
    main = getattr(package, "main_document_part", None)
    if main is None:
        return []
    out = [main]
    for name in ("footnotes_part", "endnotes_part", "comments_part"):
        part = getattr(main, name, None)
        if part is not None:
            out.append(part)
    try:
        hf = _headers_and_footers(main)
    except Exception:  # noqa: BLE001 - a main part we cannot read has none
        hf = []
    out.extend(hf)
    return out


def package_stories(package: Any, *, only_with_revisions: bool = True) -> list[Story]:
    """The stories of a package: main document, footnotes, endnotes, comments, headers and footers.

    With `only_with_revisions` (the listing's case), a part that has not been
    unmarshalled and whose bytes hold no revision markup is left out --- and
    stays unmarshalled, so that it is written back byte for byte.
    """
    out: list[Story] = []
    main = getattr(package, "main_document_part", None)
    if main is None:
        return out
    body = main.contents.body
    if body is not None:
        out.append(Story(package, main, body.content, body))
    for name, field in (("footnotes_part", "footnote"), ("endnotes_part", "endnote"), ("comments_part", "comment")):
        part = getattr(main, name, None)
        if part is None or (only_with_revisions and not _may_hold_revisions(part)):
            continue
        root = part.contents
        for note in getattr(root, field, None) or ():
            out.append(Story(package, part, note.content, note))
    for part in _headers_and_footers(main):
        if only_with_revisions and not _may_hold_revisions(part):
            continue
        root = part.contents
        items = trees.block_items(root)
        if items is not None:
            out.append(Story(package, part, items, root))
    return out


# ---------------------------------------------------------------------------
# the lister
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class _Box:
    """A text box found in a story, with its twins in the other branches."""

    story: Story
    twins: list[Story]


class _Ctx:
    """What the walk of one story carries."""

    __slots__ = ("boxes", "story")

    def __init__(self, story: Story) -> None:
        self.story = story
        self.boxes: list[_Box] = []


class _FieldState(enum.Enum):
    OPEN = 1
    CLOSED = 2
    OTHER = 3


class Lister:
    """Lists the changes of some stories (by default, every story of a package)."""

    __slots__ = ("_islands", "_stories", "attach_twins", "package")

    def __init__(
        self,
        package: Any,
        attach_twins: bool,
        stories: list[Story] | None = None,
        islands: dict[int, Island] | None = None,
    ) -> None:
        """Build the lister.

        Args:
            package: the package.
            attach_twins: True for the public listing, where a text box's
                ``mc:Fallback`` changes are attached to the ``mc:Choice``'s as
                twins; False for ``accept_all``'s, where every branch's are
                listed.
            stories: the stories to list; every story of the package when None.
            islands: the typed islands of wildcard text boxes, by ``id`` of
                the wildcard node, to take and add to: a reader that walks the
                same boxes (the text model, CR-011) then sees the very objects
                the listing names. None for islands of the lister's own.
        """
        self.package = package
        self.attach_twins = attach_twins
        self._stories = stories
        self._islands = islands

    def stories(self) -> list[Story]:
        """The stories this lister walks."""
        if self._stories is not None:
            return self._stories
        return package_stories(self.package)

    def list(self) -> list[TrackedChange]:
        """Every change, story by story; a text box's after the story it is in."""
        out: list[TrackedChange] = []
        for story in self.stories():
            self._list_story(story, out)
        return out

    def _change(self, target: Any, para: Para | None, ctx: _Ctx) -> TrackedChange:
        return TrackedChange(target, para, ctx.story)

    def _list_story(self, story: Story, out: list[TrackedChange]) -> None:
        ctx = _Ctx(story)
        tokens: list[Any] = []
        self._visit_blocks(story.root, tokens, ctx)
        owner = story.root_owner
        if isinstance(owner, trees.cls("Body")):
            sect_pr = owner.sect_pr
            if sect_pr is not None and sect_pr.sect_pr_change is not None:
                tokens.append(BREAK)
                tokens.append(
                    self._change(SectionT(sect_pr, _section_text(story.root, len(story.root))), None, ctx)
                )
        out.extend(_group_touching(tokens))
        # a text box is a story of its own (Word's text frame story), listed after the one it is in
        for box in ctx.boxes:
            primary: list[TrackedChange] = []
            self._list_story(box.story, primary)
            for twin in box.twins:
                twin_changes: list[TrackedChange] = []
                self._list_story(twin, twin_changes)
                if self.attach_twins:
                    _pair_twins(primary, twin_changes)
                else:
                    out.extend(twin_changes)
            out.extend(primary)

    # -- block level ---------------------------------------------------------

    def _visit_blocks(self, items: list, into: list, ctx: _Ctx) -> None:
        p_class, tbl_class = trees.cls("P"), trees.cls("Tbl")
        index = 0
        while index < len(items):
            item = items[index]
            if isinstance(item, p_class):
                self._paragraph_tokens(Para(item, items), into, ctx)
                index += 1
                continue
            if is_alternate_content(item):
                # block level: the branch Word reads, its changes kept apart (CR-012 Phase A)
                self._alternate_level(item, self._visit_blocks, into, ctx)
                index += 1
                continue
            # a block control or custom XML element that is itself an insertion
            # (check 30): one piece, grouping with the insertions around and in it
            inserted = self._holder_insertion_at(items, index, None, ctx)
            if inserted is not None:
                into.append(inserted)
                self._visit_blocks(trees.children(items[index + 1]) or [], into, ctx)
                index += 3
                continue
            if isinstance(item, tbl_class):
                self._table_tokens(item, items, into, ctx)
                index += 1
                continue
            if trees.revision_of(item) is not None:
                # run-level markup standing at block level (EG_BlockLevelElts admits it)
                into.append(BREAK)
                self._revision_tokens(None, items, index, into, ctx)
                into.append(BREAK)
                index += 1
                continue
            # a block-level control or custom XML element keeps what is inside it
            # apart; a marker between paragraphs keeps nothing apart
            nested = trees.block_items(item)
            if nested is not None:
                into.append(BREAK)
                self._visit_blocks(nested, into, ctx)
                into.append(BREAK)
            index += 1

    def _table_tokens(self, tbl: Any, tbl_owner: list, into: list, ctx: _Ctx) -> None:
        # its property records are one change of their own where the table
        # stands, kept apart; records that change nothing are not listed, and
        # the sweep drops them (U20)
        record = tables.record_of(tbl)
        if record is not None:
            into.append(BREAK)
            into.append(self._change(TableT(tbl, record), None, ctx))
            into.append(BREAK)
        self._row_items(tbl.content, tbl, tbl_owner, into, ctx)

    def _row_items(self, items: list, tbl: Any, tbl_owner: list, into: list, ctx: _Ctx) -> None:
        """A table's rows in order, through row-level controls and custom XML.

        docx4j ``Lister.rowItems`` (its CR-051 phase 2): a row-level
        ``mc:AlternateContent`` kept whole lists the branch Word reads, with the
        other branches' changes as twins (:meth:`_alternate_level`).
        """
        tr_class, tc_class = trees.cls("Tr"), trees.cls("Tc")
        for item in items:
            if isinstance(item, tr_class):
                self._row_tokens(item, items, tbl, tbl_owner, into, ctx)
            elif is_alternate_content(item):
                self._alternate_level(
                    item, lambda branch, tokens, c: self._row_items(branch, tbl, tbl_owner, tokens, c), into, ctx
                )
            elif not isinstance(item, tc_class):
                nested = trees.block_items(item)
                if nested is not None:
                    self._row_items(nested, tbl, tbl_owner, into, ctx)

    def _row_tokens(self, tr: Any, owner: list, tbl: Any, tbl_owner: list, into: list, ctx: _Ctx) -> None:
        # a table is no break of its own: an inserted or deleted row groups with
        # the rows and the inserted text touching it, as Office JS lists them (check 31)
        tokens: list[Any] = []
        self._cell_items(tr.content, tr, tokens, ctx)
        inner = _group_touching(tokens)
        # a row that is itself a revision reports as one change, as Office JS
        # does: the markup its cells carry belongs to that change
        row_changes: list[TrackedChange] = []
        tr_pr = tr.tr_pr
        if tr_pr is not None and tr_pr.ins is not None:
            row_changes.append(self._change(RowT(True, tr, owner, tr_pr.ins, inner, tbl, tbl_owner), None, ctx))
        if tr_pr is not None and tr_pr.del_value is not None:
            row_changes.append(self._change(RowT(False, tr, owner, tr_pr.del_value, inner, tbl, tbl_owner), None, ctx))
        if row_changes:
            into.extend(row_changes)
        else:
            into.append(BREAK)
            for change in inner:
                into.append(change)
                into.append(BREAK)

    def _cell_items(self, items: list, tr: Any, tokens: list, ctx: _Ctx) -> None:
        """A row's cells in order, through cell-level controls and custom XML.

        docx4j ``Lister.cellItems`` (its CR-051 phase 2): a cell-level
        ``mc:AlternateContent`` kept whole lists the branch Word reads, with the
        other branches' changes as twins (:meth:`_alternate_level`).
        """
        tc_class = trees.cls("Tc")
        for item in items:
            if isinstance(item, tc_class):
                self._cell_tokens(item, tr, tokens, ctx)
            elif is_alternate_content(item):
                self._alternate_level(item, lambda branch, out, c: self._cell_items(branch, tr, out, c), tokens, ctx)
            else:
                nested = trees.block_items(item)
                if nested is not None:
                    self._cell_items(nested, tr, tokens, ctx)

    def _cell_tokens(self, tc: Any, tr: Any, tokens: list, ctx: _Ctx) -> None:
        tc_pr = tc.tc_pr
        marker = getattr(tc_pr, "cell_ins_or_cell_del_or_cell_merge", None) if tc_pr is not None else None
        local = trees.name(marker) if marker is not None else ""
        if local in ("cellIns", "cellDel"):
            tokens.append(self._change(CellT(local == "cellIns", tc, tr, marker), None, ctx))
            tokens.append(BREAK)
        self._visit_blocks(tc.content, tokens, ctx)
        tokens.append(BREAK)

    # -- paragraphs and run level ---------------------------------------------

    def paragraph_changes(self, p: Any, container: list, story: Story) -> list[TrackedChange]:
        """One paragraph's changes: those starting in it, cut at its end (check 22).

        Office JS's ``Paragraph.getTrackedChanges``: the paragraph's own pieces
        grouped, less the leading ones that continue a group the paragraph
        before began (core-ts ``trackedChangesOfParagraph``).
        """
        ctx = _Ctx(story)
        tokens: list[Any] = []
        self._paragraph_tokens(Para(p, container), tokens, ctx)
        before = _mark_before(p, container)
        if before is not None:
            index = 0
            while index < len(tokens):
                token = tokens[index]
                if token is BREAK:
                    break
                kind = groupable(token)
                if kind is None:
                    index += 1
                    continue
                if kind is not before[0] or token.author != before[1]:
                    break
                del tokens[index]
        return _group_touching(tokens)

    def _paragraph_tokens(self, para: Para, out: list, ctx: _Ctx) -> None:
        """A paragraph's single changes in document order, its mark last, with a BREAK wherever unrevised content stands."""
        p = para.p
        p_pr = p.p_pr
        if p_pr is not None and p_pr.p_pr_change is not None:
            out.append(self._change(ParaPropsT(p_pr.p_pr_change), para, ctx))
        self._run_level(para, p.content, out, ctx)
        r_pr = p_pr.r_pr if p_pr is not None else None
        # a mark whose formatting changed is a piece, grouping with its runs' (check 27)
        if r_pr is not None and r_pr.r_pr_change is not None:
            out.append(self._change(MarkPropsT(r_pr.r_pr_change), para, ctx))
        if r_pr is not None and r_pr.ins is not None:
            out.append(self._change(MarkT(True, r_pr.ins), para, ctx))
        if r_pr is not None and r_pr.del_value is not None:
            out.append(self._change(MarkT(False, r_pr.del_value), para, ctx))
        if r_pr is None or (r_pr.ins is None and r_pr.del_value is None and r_pr.r_pr_change is None):
            out.append(BREAK)
        sect_pr = p_pr.sect_pr if p_pr is not None else None
        if sect_pr is not None and sect_pr.sect_pr_change is not None:
            at = trees.index_of(para.container, p)
            out.append(self._change(SectionT(sect_pr, _section_text(para.container, at + 1) + "\f"), para, ctx))

    def _run_level(self, para: Para | None, items: list | None, out: list, ctx: _Ctx) -> None:
        if not items:
            return
        r_class = trees.cls("R")
        index = 0
        while index < len(items):
            item = items[index]
            # an inline control or custom XML element that is itself an insertion
            inserted = self._holder_insertion_at(items, index, para, ctx)
            if inserted is not None:
                out.append(inserted)
                self._run_level(para, trees.children(items[index + 1]), out, ctx)
                index += 3
                continue
            # a field Word inserted, its separator and result plain between two insertions
            field = self._inserted_field_at(items, index, para, ctx)
            if field is not None:
                out.append(field)
                index += len(field.target.span)
                continue
            if trees.revision_of(item) is not None:
                self._revision_tokens(para, items, index, out, ctx)
                index += 1
                continue
            if isinstance(item, r_class):
                self._find_text_boxes(item, ctx)
                r_pr = item.r_pr
                # a run whose formatting changed is a piece of its own; one that did
                # not keeps pieces apart if it has text (a comment reference does not)
                if r_pr is not None and r_pr.r_pr_change is not None:
                    out.append(self._change(RunPropsT(item, r_pr.r_pr_change), para, ctx))
                elif trees.has_text(item):
                    out.append(BREAK)
                index += 1
                continue
            if trees.is_math(item):
                self._math(para, item, items, [], out, ctx)
                index += 1
                continue
            if is_alternate_content(item):
                # paragraph level: the branch Word reads, its changes kept apart (CR-012 Phase A)
                self._alternate_level(item, lambda branch, tokens, c: self._run_level(para, branch, tokens, c), out, ctx)
                index += 1
                continue
            # a run holder that is not itself an insertion keeps the insertions
            # inside it apart from those outside (check 30)
            nested = trees.run_items(item)
            if nested is not None:
                out.append(BREAK)
                self._run_level(para, nested, out, ctx)
                out.append(BREAK)
            index += 1

    def _revision_tokens(self, para: Para | None, items: list, index: int, out: list, ctx: _Ctx) -> None:
        """``items[index]``, a run-level revision: its piece, then what is nested in it."""
        item = items[index]
        revision = trees.revision_of(item)
        # moved text is content Office JS cannot list (checks 16, 17), and is kept apart
        move = revision in (Revision.MOVE_FROM, Revision.MOVE_TO)
        if move:
            out.append(BREAK)
        out.append(self._change(RunT(revision, item, items), para, ctx))
        if move:
            out.append(BREAK)
        nested = trees.run_items(item)
        if nested is None:
            return
        # a revision inside a revision (a deletion of inserted text, check 22) is listed apart
        inner: list[Any] = []
        self._run_level(para, nested, inner, ctx)
        if any(isinstance(t, TrackedChange) and isinstance(t.target, RunT) for t in inner):
            out.append(BREAK)
            out.extend(inner)
            out.append(BREAK)
        else:
            out.extend(t for t in inner if t is not BREAK)

    # -- content controls and fields inserted ----------------------------------

    def _holder_insertion_at(self, items: list, index: int, para: Para | None, ctx: _Ctx) -> TrackedChange | None:
        """Word's form for a content control or custom XML element inserted or deleted under tracking.

        Inserted (check 30): ``w:customXmlInsRangeStart``, the element --- whose
        content starts with that marker's ``w:customXmlInsRangeEnd`` and ends
        with a second start --- and that second start's end. Deleted, the same
        form of ``w:customXmlDelRangeStart`` and ``End``, as Word writes a
        control a tracked deletion took whole (docx4j-core-ts check 54 cases 15
        and 17, its ``holderChangeAt``; CR-013 Phase C).
        """
        start = items[index]
        local = trees.name(start)
        if local == "customXmlInsRangeStart":
            start_name, end_name, ins = "customXmlInsRangeStart", "customXmlInsRangeEnd", True
        elif local == "customXmlDelRangeStart":
            start_name, end_name, ins = "customXmlDelRangeStart", "customXmlDelRangeEnd", False
        else:
            return None
        if index + 2 >= len(items):
            return None
        holder = items[index + 1]
        outer_end = items[index + 2]
        if not trees.is_holder(holder) or trees.name(outer_end) != end_name:
            return None
        inner = trees.children(holder)
        if inner is None or len(inner) < 2:
            return None
        inner_end, inner_start = inner[0], inner[-1]
        if trees.name(inner_end) != end_name or trees.name(inner_start) != start_name:
            return None
        if trees.marker_id(inner_end) != trees.marker_id(start) or trees.marker_id(outer_end) != trees.marker_id(
            inner_start
        ):
            return None
        return self._change(HolderT(holder, items, start, outer_end, inner_end, inner_start, ins), para, ctx)

    def _inserted_field_at(self, items: list, index: int, para: Para | None, ctx: _Ctx) -> TrackedChange | None:
        """Word's form for a field inserted under tracking, when ``items[index]`` starts it (check 30, C01, C16).

        A ``w:ins`` that opens a field, then plain runs --- the separator and the
        result --- and markers, then a ``w:ins`` by the same author in which that
        field ends.
        """
        first = items[index]
        if trees.revision_of(first) is not Revision.INS:
            return None
        author = first.author
        depth = [0]
        if _scan(first, depth) is not _FieldState.OPEN or depth[0] == 0:
            return None
        r_class = trees.cls("R")
        for j in range(index + 1, len(items)):
            element = items[j]
            kind = trees.revision_of(element)
            if kind is Revision.INS:
                if element.author != author:
                    return None
                state = _scan(element, depth)
                if state is _FieldState.OTHER:
                    return None
                if state is _FieldState.CLOSED:
                    return self._change(FieldT(list(items[index : j + 1]), items), para, ctx)
                continue
            if kind is not None:
                return None
            if isinstance(element, r_class):
                r_pr = element.r_pr
                if (r_pr is not None and r_pr.r_pr_change is not None) or _scan(element, depth) is not _FieldState.OPEN:
                    return None
                continue
            if trees.run_items(element) is not None:
                return None
        return None

    # -- math --------------------------------------------------------------------

    def _math(self, para: Para | None, value: Any, owner: list | None, stack: list, out: list, ctx: _Ctx) -> None:
        """An equation's pieces (CR-043 section 3.2).

        ``w:ins`` / ``w:del`` inside an ``m:r``, a structure inserted or deleted
        (its ``m:ctrlPr``'s ``w:ins`` / ``w:del``), and run-level revisions
        holding math. `stack` holds ``(value, owner list)`` of the objects above.
        """
        ctr, ctrl_pr_class = trees.cls("CTR"), trees.cls("CTCtrlPr")
        if isinstance(value, ctr):
            track = trees.cls("CTMathRunTrackChange")
            for child in value.content:
                if isinstance(child, track):
                    out.append(self._change(MathRunT(trees.name(child) == "ins", value, child, owner), para, ctx))
                elif trees.item_text(child) is not None:
                    out.append(BREAK)
            return
        if isinstance(value, ctrl_pr_class):
            marker = value.r_pr_or_ins_or_del
            local = trees.name(marker) if marker is not None else ""
            if local not in ("ins", "del"):
                return
            struct_item = struct_owner = None
            parent = stack[-1] if stack else None
            if parent is not None and not isinstance(parent[0], trees.cls("CTOMathArg")) and len(stack) >= 2:
                # m:ctrlPr in m:sSupPr in m:sSup: the structure is two up
                struct_item, struct_owner = stack[-2]
            out.append(self._change(MathStructT(local == "ins", value, marker, struct_item, struct_owner), para, ctx))
            return
        stack.append((value, owner))
        try:
            r_class = trees.cls("R")
            for field in dataclasses.fields(value):
                found = getattr(value, field.name, None)
                if found is None or isinstance(found, str):
                    continue
                if isinstance(found, list):
                    k = 0
                    while k < len(found):
                        child = found[k]
                        if trees.revision_of(child) is not None:
                            self._revision_tokens(para, found, k, out, ctx)
                        elif isinstance(child, r_class):
                            self._run_level(para, [child], out, ctx)
                        elif trees.is_math(child) and dataclasses.is_dataclass(child):
                            self._math(para, child, found, stack, out, ctx)
                        k += 1
                elif trees.is_math(found) and dataclasses.is_dataclass(found):
                    self._math(para, found, None, stack, out, ctx)
        finally:
            stack.pop()

    # -- mc:AlternateContent at block, paragraph, row and cell level (CR-012 Phases A and A2) --

    def _alternate_level(self, ac: Any, walk: Any, into: list, ctx: _Ctx) -> None:
        """A kept ``mc:AlternateContent``: the branch Word reads, and its twins.

        docx4j ``Lister`` since its CR-051 phase 1 (``alternateBlocks``), and
        since its phase 2 among a table's rows and a row's cells
        (``alternateRows``, ``alternateCells``; ``TrackedChangesAlternateContentTest``):
        the changes of the branch a reader of text takes
        (:func:`docx4j_py.child.mce_branch`) are listed where the element
        stands, kept apart from what is around it, as a block control's are.
        Each other branch's changes are that branch's twins, the n-th of a kind
        with the n-th of that kind (:func:`_pair_twins`, as a text box's
        ``mc:Fallback``), so resolving one resolves the other; a row or cell
        that is itself a revision in another branch is the twin of the read
        branch's, and a text box in another branch the twin of the read
        branch's at the same position. ``accept_all``'s listing
        (``attach_twins`` False) lists every branch's changes, the others'
        marked :attr:`TrackedChange.twin` so that they are not counted.

        `walk` tokenises a branch's live list as the element's level does:
        ``walk(items, tokens, ctx)``, blocks, a paragraph's runs, rows or cells.
        """
        branches = _branch_objects(ac)
        selected = mce_branch(ac)
        if selected is None:
            return

        def changes_of(branch: Any, context: _Ctx) -> list[TrackedChange]:
            tokens: list[Any] = []
            walk(list_of_branch(branch), tokens, context)
            return _group_touching(tokens)

        first_box = len(ctx.boxes)
        primary = changes_of(selected, ctx)
        primary_boxes = ctx.boxes[first_box:]
        into.append(BREAK)
        for change in primary:
            into.append(change)
            into.append(BREAK)
        for branch in branches:
            if branch is selected:
                continue
            other = _Ctx(ctx.story)
            twins = changes_of(branch, other)
            if self.attach_twins:
                _pair_twins(primary, twins)
                for position, box in enumerate(other.boxes):
                    if position < len(primary_boxes):
                        primary_boxes[position].twins.append(box.story)
                        primary_boxes[position].twins.extend(box.twins)
                continue
            for change in twins:
                change.twin = True
                into.append(change)
                into.append(BREAK)
            for box in other.boxes:
                box.story.twin = True
                ctx.boxes.append(box)

    # -- text boxes ------------------------------------------------------------

    def _find_text_boxes(self, run: Any, ctx: _Ctx) -> None:
        """The text boxes in a run, each a story; in ``mc:AlternateContent``, the first branch's, the others' as twins."""
        # a text box lives in a w:drawing, a w:pict, a w:object or an
        # mc:AlternateContent; a run with none of them is not walked
        if not any(trees.name(item) in _BOX_HOLDERS for item in run.content):
            return
        for ref in box_refs(run, alternate=True):
            if ref[0] == "ac":
                self._alternate(ref[1], ctx)
                continue
            story = self._box_story(ref, ctx, ctx.story.twin)
            if story is not None:
                ctx.boxes.append(_Box(story, []))

    def _alternate(self, ac: Any, ctx: _Ctx) -> None:
        branches = [box_refs(branch, alternate=False) for branch in _branches(ac)]
        primary = next((branch for branch in branches if branch), None)
        if primary is None:
            return
        for index, ref in enumerate(primary):
            story = self._box_story(ref, ctx, ctx.story.twin)
            twins: list[Story] = []
            for branch in branches:
                if branch is primary or len(branch) <= index:
                    continue
                twin = self._box_story(branch[index], ctx, True)
                if twin is not None:
                    twins.append(twin)
            if story is None:
                if not twins:
                    continue
                story = twins.pop(0)
                story.twin = ctx.story.twin
            ctx.boxes.append(_Box(story, twins))

    def _box_story(self, ref: tuple, ctx: _Ctx, twin: bool) -> Story | None:
        kind, node, holder = ref
        part = ctx.story.part
        if kind == "typed":
            return Story(self.package, part, node.content, node, twin)
        if not holds_revisions(node):
            return None
        island = self._islands.get(id(node)) if self._islands is not None else None
        if island is None:
            try:
                island = Island(node, holder)
            except ValueError:
                return None
            if self._islands is not None:
                self._islands[id(node)] = island
        return Story(self.package, part, island.typed.content, island.typed, twin, island)


# ---------------------------------------------------------------------------
# text boxes: where they are
# ---------------------------------------------------------------------------


def _branch_objects(ac: Any) -> list[Any]:
    """The branches of an ``mc:AlternateContent``, typed or wildcard: its choices, then its fallback."""
    if is_any_element(ac):
        choices = [c for c in ac.children or () if is_any_element(c) and c.qname == MC_CHOICE]
        fallbacks = [c for c in ac.children or () if is_any_element(c) and c.qname == MC_FALLBACK]
        return [*choices, *fallbacks[:1]]
    out = list(ac.choice or ())
    if ac.fallback is not None:
        out.append(ac.fallback)
    return out


def list_of_branch(branch: Any) -> list:
    """A branch's live list: a typed one's ``any_element``, a wildcard one's ``children``."""
    if is_any_element(branch):
        return branch.children
    return branch.any_element


def _branches(ac: Any) -> list[list]:
    """The content of each branch of an ``mc:AlternateContent``, typed or wildcard: its choices, then its fallback."""
    if is_any_element(ac):
        choices = [c.children for c in ac.children or () if is_any_element(c) and c.qname == MC_CHOICE]
        fallbacks = [c.children for c in ac.children or () if is_any_element(c) and c.qname == MC_FALLBACK]
        return [*choices, *fallbacks[:1]]
    out = [choice.any_element for choice in ac.choice or ()]
    if ac.fallback is not None:
        out.append(ac.fallback.any_element)
    return out


def box_refs(root: Any, *, alternate: bool) -> list[tuple]:
    """The text boxes under `root` in document order, and (with `alternate`) the ``mc:AlternateContent`` met.

    Each is ``("typed", txbx, None)``, ``("wild", node, holder_list)`` or
    ``("ac", node, None)``; nothing under a box or an ``mc:AlternateContent``
    is looked into.
    """
    out: list[tuple] = []
    txbx_class = trees.cls("CTTxbxContent")

    def visit_list(items: list, wildcard_list: bool) -> None:
        for item in list(items):
            visit(item, items if wildcard_list else None)

    def visit(node: Any, holder: list | None) -> None:
        if isinstance(node, str):
            return
        if is_alternate_content(node):
            if alternate:
                out.append(("ac", node, None))
                return
            for branch in _branches(node):
                visit_list(branch, True)
            return
        if isinstance(node, txbx_class):
            out.append(("typed", node, None))
            return
        if is_wildcard_txbx(node):
            if holder is not None:
                out.append(("wild", node, holder))
            return
        if is_any_element(node):
            visit_list(node.children or [], True)
            return
        if not dataclasses.is_dataclass(node):
            return
        for _qname, child in iter_children(node, mce="all", wildcards=True):
            if is_any_element(child):
                visit(child, _holder_of(node, child))
            else:
                visit(child, None)

    if isinstance(root, list):
        visit_list(root, True)
    else:
        visit(root, None)
    return out


def _holder_of(typed: Any, child: Any) -> list | None:
    """The list of a typed node that holds a wildcard child (its ``any_element`` field, usually)."""
    for field in dataclasses.fields(typed):
        found = getattr(typed, field.name, None)
        if isinstance(found, list) and any(item is child for item in found):
            return found
    return None


# ---------------------------------------------------------------------------
# grouping
# ---------------------------------------------------------------------------


def groupable(change: TrackedChange) -> GroupKind | None:
    """What a piece groups as; None for moves, a paragraph's properties and the records."""
    t = change.target
    if isinstance(t, (MarkT, RowT, MathRunT, MathStructT)):
        return GroupKind.INS if t.ins else GroupKind.DEL
    if isinstance(t, RunT):
        if t.revision is Revision.INS:
            return GroupKind.INS
        return GroupKind.DEL if t.revision is Revision.DEL else None
    if isinstance(t, (RunPropsT, MarkPropsT)):
        return GroupKind.FORMAT
    if isinstance(t, HolderT):
        return GroupKind.INS if t.ins else GroupKind.DEL
    if isinstance(t, FieldT):
        return GroupKind.INS
    return None


def _group_touching(tokens: list) -> list[TrackedChange]:
    """Tokens made into the changes Office JS lists (checks 22, 23, 28).

    Consecutive pieces of one kind by one author, nothing unrevised between
    them, are one change; a change that does not group is listed where it
    stands and does not break a group.
    """
    out: list[Any] = []
    open_group: list[TrackedChange] | None = None
    for token in tokens:
        if token is BREAK:
            open_group = None
            continue
        kind = groupable(token)
        if kind is None:
            out.append(token)
            continue
        last = open_group[-1] if open_group else None
        if last is not None and groupable(last) is kind and last.author == token.author:
            open_group.append(token)
            continue
        open_group = [token]
        out.append(open_group)
    result: list[TrackedChange] = []
    for entry in out:
        if isinstance(entry, TrackedChange):
            result.append(entry)
            continue
        first = entry[0]
        if len(entry) == 1:
            result.append(first)
        else:
            result.append(TrackedChange(GroupT(groupable(first), entry), first.para, first.story))
    return result


def _pair_twins(primary: list[TrackedChange], twins: list[TrackedChange]) -> None:
    """A text box's ``mc:Fallback`` changes attached to the ``mc:Choice``'s, by position.

    The n-th change of a kind with the n-th of that kind, never by id (Word
    gives the twins different ids, C22; CR-043 section 10 item 2).
    """
    by_kind: dict[Any, list[TrackedChange]] = {}
    for twin in twins:
        by_kind.setdefault(twin.kind, []).append(twin)
    seen: dict[Any, int] = {}
    for change in primary:
        n = seen.get(change.kind, 0)
        seen[change.kind] = n + 1
        found = by_kind.get(change.kind)
        if found is not None and n < len(found):
            change.twins.append(found[n])


def _mark_before(p: Any, container: list) -> tuple[GroupKind, Any] | None:
    """The last revision on the mark of the paragraph before this one in its container, if any."""
    index = trees.index_of(container, p)
    previous = container[index - 1] if index > 0 else None
    if previous is None or not isinstance(previous, trees.cls("P")):
        return None
    p_pr = previous.p_pr
    r_pr = p_pr.r_pr if p_pr is not None else None
    if r_pr is None:
        return None
    if r_pr.del_value is not None:
        return GroupKind.DEL, r_pr.del_value.author
    if r_pr.ins is not None:
        return GroupKind.INS, r_pr.ins.author
    if r_pr.r_pr_change is not None:
        return GroupKind.FORMAT, r_pr.r_pr_change.author
    return None


def _section_text(items: list, end: int) -> str:
    """The text of the section ending at `end` (exclusive): its paragraphs', ``\\r`` apart, back to the previous break."""
    p_class = trees.cls("P")
    start = end - 1
    while start > 0:
        previous = items[start - 1]
        if isinstance(previous, p_class) and previous.p_pr is not None and previous.p_pr.sect_pr is not None:
            break
        start -= 1
    texts = [
        trees.paragraph_text(item, View.ACCEPTED) for item in items[max(start, 0) : end] if isinstance(item, p_class)
    ]
    return "\r".join(texts)


def _scan(element: Any, depth: list[int]) -> _FieldState:
    """The field characters of an element's runs in order: CLOSED when the field the span opened ends."""
    r_class = trees.cls("R")
    runs = [element] if isinstance(element, r_class) else list(trees.run_items(element) or ())
    for run in runs:
        if not isinstance(run, r_class):
            continue
        for item in run.content:
            if trees.name(item) != "fldChar":
                continue
            kind = getattr(getattr(item, "fld_char_type", None), "value", None)
            if kind == "begin":
                depth[0] += 1
            elif kind == "end":
                depth[0] -= 1
                if depth[0] < 0:
                    return _FieldState.OTHER
                if depth[0] == 0:
                    return _FieldState.CLOSED
    return _FieldState.OPEN
