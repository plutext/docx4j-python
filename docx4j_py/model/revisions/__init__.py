"""Tracked changes listed, accepted and rejected as Word resolves them. CR-010.

The port of docx4j's ``org.docx4j.model.revisions`` (docx4j CR-043, commit
``00614748e`` on ``VERSION_17_3_3``): every story --- the main document, the
footnotes, endnotes and comments, the headers and footers --- and inside them
content controls, text boxes in both branches of ``mc:AlternateContent`` (as
typed islands, CR-010 section 3.5) and equations, with **Word 365's own Accept
All and Reject All as the oracle**. The rules, and the Word check that measured
each, are docx4j CR-043 section 3.2; CR-010 section 10 records where this port
builds one differently.

::

    from docx4j_py.model import revisions

    for change in revisions.tracked_changes(pkg):
        print(change.kind, change.author, repr(change.text))
    revisions.accept_all(pkg)          # or reject_all; each returns the count

The content API's :class:`~docx4j_py.model.content.tracked_change.TrackedChange`
is the view over these (CR-010 section 3.2): ``pkg.get_tracked_changes()``,
``Body`` / ``Paragraph`` / ``Range.get_tracked_changes()``, and
``Body.accept_all()`` / ``reject_all()``.

Phase B (CR-010 section 5, docx4j CR-043 phase 2) resolves the table, grid,
row, cell and section records too, sweeps the legacy forms nothing lists, and
adds :class:`TrackedChangeFilter` and :func:`report`::

    revisions.reject_all(pkg, TrackedChangeFilter(authors=["Author B"]))
    revisions.report(pkg).to_dict()   # counts by kind, author and part

Imported on first use: nothing under :mod:`docx4j_py.model.content` imports this
at module level, and this imports no model class at module level.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docx4j_py.model.revisions.change import Kind, TrackedChange
from docx4j_py.model.revisions.filter import Filter, TrackedChangeFilter, matcher
from docx4j_py.model.revisions.generation import Generation
from docx4j_py.model.revisions.report import TrackedChangesReport
from docx4j_py.model.revisions.story import Story

__all__ = [
    "Filter",
    "Kind",
    "Story",
    "TrackedChange",
    "TrackedChangeFilter",
    "TrackedChangesReport",
    "accept_all",
    "body_stories",
    "reject_all",
    "report",
    "resolve_all",
    "tracked_changes",
    "tracked_changes_of",
]


def tracked_changes(package: Any, filter: Filter | None = None) -> list[TrackedChange]:
    """Every tracked change, story by story, each in document order, grouped as Office JS groups them.

    The main document, footnotes, endnotes, comments, then headers and footers
    in the order the sections reference them; a text box's changes after the
    story it is in, listed once from its ``mc:Choice`` (resolving one resolves
    its ``mc:Fallback`` twin too). The views are valid until the next accept or
    reject on the package. With a `filter`, the changes any of whose pieces it
    matches.
    """
    match = matcher(filter, package)
    changes = tracked_changes_of(package, None)
    if match is None:
        return changes
    return [change for change in changes if any(match(piece) for piece in change.pieces)]


def tracked_changes_of(package: Any, stories: list[Story] | None, key: Any = None) -> list[TrackedChange]:
    """The listing of some stories (every story of the package when None), stamped valid from now."""
    from docx4j_py.model.revisions.lister import Lister

    changes = Lister(package, True, stories).list()
    generation = Generation.of(key if key is not None else package)
    for change in changes:
        _stamp(change, generation)
    return changes


def _stamp(change: TrackedChange, generation: Generation) -> None:
    """A change, and the pieces it groups (each a view that can be resolved on its own), valid from now."""
    change.generation = generation
    pieces = change.pieces
    if len(pieces) > 1 or pieces[0] is not change:
        for piece in pieces:
            _stamp(piece, generation)


def accept_all(package: Any, filter: Filter | None = None) -> int:
    """Accepts every change in every story; returns how many the listing had.

    With a `filter` (a :class:`TrackedChangeFilter`, or a callable over single
    changes), only the pieces it matches are accepted, and nothing is swept;
    the count is of the changes it matched.
    """
    return resolve_all(package, True, filter)


def reject_all(package: Any, filter: Filter | None = None) -> int:
    """Rejects every change in every story; returns how many the listing had. The mirror of :func:`accept_all`."""
    return resolve_all(package, False, filter)


def report(package: Any) -> TrackedChangesReport:
    """The listing counted by kind, author and part (docx4j ``TrackedChanges.report``)."""
    return TrackedChangesReport(tracked_changes(package))


def _matching(change: TrackedChange, match: Callable[[TrackedChange], bool] | None) -> list[TrackedChange]:
    """The pieces of a change the filter matches (all of them with no filter)."""
    if match is None:
        return change.pieces
    return [piece for piece in change.pieces if match(piece)]


def resolve_all(
    package: Any,
    accept: bool,
    filter: Filter | None = None,
    *,
    stories: list[Story] | None = None,
    key: Any = None,
) -> int:
    """Accept or reject every change of some stories (every story when None); returns the count listed.

    Last change first, so that a paragraph join never disturbs a change still
    to do; then the list is read again and the pass repeated until nothing is
    left, since resolving a move can carry a deletion nested in it out as a
    change of its own, and resolving a broken move leaves plain ``w:ins`` /
    ``w:del`` (core-ts ``Body.resolveAll``); then, with no filter, what is
    never listed is swept. Both branches of a text box's
    ``mc:AlternateContent`` are walked. A filter's range is read once, before
    anything is resolved.
    """
    from docx4j_py.model.revisions.lister import Lister
    from docx4j_py.model.revisions.removals import Removals
    from docx4j_py.model.revisions.resolver import resolve, resolve_pieces

    match = matcher(filter, package)
    count = 0
    removals = Removals(package)
    left = None
    stalled = 0
    for round_ in range(20):
        if stalled >= 2:
            break
        changes = Lister(package, False, stories).list()
        if round_ == 0:
            # how many the public listing has: there a text box's mc:Fallback
            # twins are attached to the mc:Choice's changes, not listed
            count = sum(1 for change in changes if not change.story.twin and _matching(change, match))
        matched = 0
        for change in reversed(changes):
            pieces = _matching(change, match)
            if not pieces:
                continue
            matched += 1
            if len(pieces) == len(change.pieces):
                resolve(change, accept, removals)
            else:
                resolve_pieces(pieces, accept, removals)
        # the typed islands go back into their wildcard trees before the next listing reads them
        removals.write_back()
        if matched == 0:
            break
        stalled = stalled + 1 if left is not None and matched >= left else 0
        left = matched
    if match is None:
        _sweep(package, stories, accept, removals)
    removals.flush()
    Generation.of(key if key is not None else package).advance()
    return count


def _sweep(package: Any, stories: list[Story] | None, accept: bool, removals: Any) -> None:
    """What is never listed (core-ts ``resolveUnlisted``, docx4j ``TrackedChanges.sweep``).

    * a table's records that change nothing, dropped either way (U20);
    * a ``w:cellMerge``, which Office JS cannot list: a rejected one gives the
      merge it recorded back (U8);
    * a ``w:numberingChange`` (in ``w:numPr`` and ``w:fldChar``), which records
      the old number's text and not its definition: dropped either way;
    * a ``w:numPr/w:ins``: accepted, the marker goes; rejected, the
      ``w:numPr`` with it (U9);
    * a ``w:customXmlInsRangeStart`` or ``End`` left without the form of an
      inserted control, which marks nothing that can be found.

    A record's own content is what was, not what is, and is not looked into.
    """
    from docx4j_py.model.revisions.lister import package_stories

    done: set[int] = set()
    for story in stories if stories is not None else package_stories(package):
        if id(story.root) in done:
            continue
        done.add(id(story.root))
        if _sweep_content(story.root, accept, removals):
            removals.touched(story)


def _stray(item: Any) -> bool:
    from docx4j_py.model.revisions import trees

    return trees.name(item) in ("customXmlInsRangeStart", "customXmlInsRangeEnd")


#: The run children a text box can be inside.
_BOX_HOLDERS = frozenset({"drawing", "pict", "object", "AlternateContent"})


def _sweep_content(items: list, accept: bool, removals: Any) -> bool:
    """Sweeps a content list and the content lists under it; True when anything changed.

    The forms swept stand in content lists (markers), on what they hold
    (tables, cells, paragraphs) or in a run (``w:fldChar``), so only those are
    walked, and what a run holds only where a text box (a drawing, a picture,
    ``mc:AlternateContent``) or an equation may hold content of its own.
    """
    from docx4j_py.model.revisions import tables, trees

    r_class, p_class = trees.cls("R"), trees.cls("P")
    tbl_class, tc_class = trees.cls("Tbl"), trees.cls("Tc")
    changed = False
    stack = [items]
    while stack:
        current = stack.pop()
        for index in range(len(current) - 1, -1, -1):
            item = current[index]
            if _stray(item):
                del current[index]
                changed = True
                continue
            if isinstance(item, r_class):
                for child in item.content:
                    local = trees.name(child)
                    if local in _BOX_HOLDERS:
                        if removals.remove_everywhere(child, _stray):
                            changed = True
                        if _sweep_boxes(child, accept, removals):
                            changed = True
                    elif local == "fldChar" and _drop_numbering_change(child):
                        changed = True
                continue
            if trees.is_math(item):
                if removals.remove_everywhere(item, _stray):
                    changed = True
                continue
            if isinstance(item, p_class):
                changed = _sweep_numbering(item.p_pr, accept) or changed
            elif isinstance(item, tbl_class) and tables.has_records(item):
                tables.accept(item)
                changed = True
            elif isinstance(item, tc_class) and tables.cell_marker(item.tc_pr, "cellMerge") is not None:
                tables.resolve_cell_merge(item.tc_pr, accept)
                changed = True
            nested = trees.children(item)
            if nested is not None:
                stack.append(nested)
    return changed


def _sweep_numbering(p_pr: Any, accept: bool) -> bool:
    """A paragraph's ``w:numPr/w:numberingChange`` dropped, and its ``w:numPr/w:ins`` resolved (U9)."""
    num_pr = p_pr.num_pr if p_pr is not None else None
    if num_pr is None:
        return False
    changed = False
    if num_pr.numbering_change is not None:
        num_pr.numbering_change = None
        changed = True
    if num_pr.ins is not None:
        if accept:
            num_pr.ins = None
        else:
            p_pr.num_pr = None
        changed = True
    return changed


def _drop_numbering_change(fld_char: Any) -> bool:
    """A ``w:fldChar``'s ``w:numberingChange`` dropped (one choice with ``w:fldData`` and ``w:ffData``)."""
    from docx4j_py.model.revisions import trees

    value = getattr(fld_char, "fld_data_or_ff_data_or_numbering_change", None)
    if value is None or trees.name(value) != "numberingChange":
        return False
    fld_char.fld_data_or_ff_data_or_numbering_change = None
    return True


def _sweep_boxes(node: Any, accept: bool, removals: Any) -> bool:
    """The text boxes under a run's drawing, picture or ``mc:AlternateContent``, every branch's, swept.

    A text box kept as a wildcard tree is swept as a typed island (CR-010
    section 3.5), and only when it holds revision markup; it is written back
    only when the sweep changed it.
    """
    from docx4j_py.model.revisions.islands import Island, holds_revisions
    from docx4j_py.model.revisions.lister import box_refs

    changed = False
    for kind, box, holder in box_refs(node, alternate=False):
        if kind == "typed":
            if _sweep_content(box.content, accept, removals):
                changed = True
            continue
        if kind != "wild" or not holds_revisions(box):
            continue
        try:
            island = Island(box, holder)
        except ValueError:
            continue
        if _sweep_content(island.typed.content, accept, removals):
            island.dirty = True
            island.write_back()
            changed = True
    return changed


def body_stories(body: Any) -> list[Story]:
    """The stories a content-API ``Body`` is over.

    The main document's ``w:body`` is one story; a notes or comments part is a
    story per note or comment, as docx4j lists them; anything else --- a
    header, a footer, a cell, a content control's content --- is one story over
    the body's own list.
    """
    from docx4j_py.model.revisions import trees

    package = body.package
    part = body.part
    container = body.container
    content = body.content
    for field in ("footnote", "endnote", "comment"):
        notes = getattr(container, field, None)
        if isinstance(notes, list) and notes is content:
            return [Story(package, part, note.content, note) for note in notes]
    if isinstance(container, trees.cls("Body")):
        return [Story(package, part, content, container)]
    return [Story(package, part, content, container)]
