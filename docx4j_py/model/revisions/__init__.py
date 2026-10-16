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

Phase A (CR-010 section 4) lists the table, cell and section property records
and leaves them to Phase B: ``accept_all`` and ``reject_all`` resolve every
other kind and leave those listed, so the listing never says a document is
clean while Word still shows a revision.

Imported on first use: nothing under :mod:`docx4j_py.model.content` imports this
at module level, and this imports no model class at module level.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docx4j_py.model.revisions.change import Kind, TrackedChange, is_resolvable
from docx4j_py.model.revisions.generation import Generation
from docx4j_py.model.revisions.story import Story

__all__ = [
    "Kind",
    "Story",
    "TrackedChange",
    "accept_all",
    "body_stories",
    "is_resolvable",
    "reject_all",
    "resolve_all",
    "tracked_changes",
    "tracked_changes_of",
]

#: A filter: a callable taking one piece (a single change) and answering whether
#: it is resolved. CR-010 Phase B gives the named ``TrackedChangeFilter``.
Filter = Callable[[TrackedChange], bool]


def tracked_changes(package: Any) -> list[TrackedChange]:
    """Every tracked change, story by story, each in document order, grouped as Office JS groups them.

    The main document, footnotes, endnotes, comments, then headers and footers
    in the order the sections reference them; a text box's changes after the
    story it is in, listed once from its ``mc:Choice`` (resolving one resolves
    its ``mc:Fallback`` twin too). The views are valid until the next accept or
    reject on the package.
    """
    return tracked_changes_of(package, None)


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

    With a `filter`, only the pieces it matches are accepted, and nothing is
    swept. Table, cell and section records are left listed (CR-010 Phase B).
    """
    return resolve_all(package, True, filter)


def reject_all(package: Any, filter: Filter | None = None) -> int:
    """Rejects every change in every story; returns how many the listing had."""
    return resolve_all(package, False, filter)


def _matching(change: TrackedChange, filter: Filter | None) -> list[TrackedChange]:
    """The pieces of a change this release resolves and the filter matches (all of them with no filter)."""
    return [
        piece
        for piece in change.pieces
        if is_resolvable(piece.kind) and (filter is None or filter(piece))
    ]


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
    ``mc:AlternateContent`` are walked.
    """
    from docx4j_py.model.revisions.lister import Lister
    from docx4j_py.model.revisions.removals import Removals
    from docx4j_py.model.revisions.resolver import resolve, resolve_pieces

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
            count = sum(1 for change in changes if not change.story.twin and _matching(change, filter))
        matched = 0
        for change in reversed(changes):
            pieces = _matching(change, filter)
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
    if filter is None:
        _sweep(package, stories, removals)
    removals.flush()
    Generation.of(key if key is not None else package).advance()
    return count


def _sweep(package: Any, stories: list[Story] | None, removals: Any) -> None:
    """What is never listed: a ``w:customXmlInsRangeStart`` or ``End`` left without the form of an inserted control.

    It marks nothing that can be found (docx4j phase 1's sweep; the legacy
    forms of phase 2 are CR-010 Phase B's).
    """
    from docx4j_py.model.revisions import trees
    from docx4j_py.model.revisions.lister import package_stories

    def stray(item: Any) -> bool:
        return trees.name(item) in ("customXmlInsRangeStart", "customXmlInsRangeEnd")

    done: set[int] = set()
    for story in stories if stories is not None else package_stories(package):
        if id(story.root) in done:
            continue
        done.add(id(story.root))
        _sweep_content(story.root, stray, removals)


def _sweep_content(items: list, match: Callable[[Any], bool], removals: Any) -> None:
    """Removes what `match` matches from a content list and the content lists under it.

    Range markers stand only in content lists --- block, row, cell and run
    level --- so only those are walked, and what a run holds only where a text
    box (a drawing, a picture, ``mc:AlternateContent``) or an equation may hold
    content of its own.
    """
    from docx4j_py.model.revisions import trees

    r_class = trees.cls("R")
    stack = [items]
    while stack:
        current = stack.pop()
        for index in range(len(current) - 1, -1, -1):
            item = current[index]
            if match(item):
                del current[index]
                continue
            if isinstance(item, r_class):
                for child in item.content:
                    if trees.name(child) in ("drawing", "pict", "object", "AlternateContent"):
                        removals.remove_everywhere(child, match)
                continue
            if trees.is_math(item):
                removals.remove_everywhere(item, match)
                continue
            nested = trees.children(item)
            if nested is not None:
                stack.append(nested)


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
