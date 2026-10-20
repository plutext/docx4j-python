"""``TrackedChange``: Office JS ``Word.TrackedChange``, the view over the revisions engine.

CR-003 section 3.8 (Phase F) and CR-010 section 3.2. The reading half; the
writing half is :mod:`.tracking`, and the engine underneath both readers is
:mod:`docx4j_py.model.revisions` --- the port of docx4j's
``org.docx4j.model.revisions`` (docx4j CR-043), which lists, accepts and
rejects tracked changes **as Word's Review tab resolves them**, in every story.

::

    for change in pkg.body.get_tracked_changes():
        print(change.type, change.author, change.text)
    pkg.body.accept_all()

What one change is follows Office JS's grouping, measured against Word (CR-043
section 3.1): touching insertions, deletions or formatting changes by one
author are **one** change, whatever their dates; another author or type splits;
a revision nested in another is listed apart; a move is never grouped. A
paragraph's mark is a piece whose text is ``\\r``, so a deleted paragraph is one
change, ``"Two.\\r"``. Office JS's members keep their shape (``type``,
``author``, ``date``, ``text``, ``get_range()``, ``accept()``, ``reject()``,
``to_dict()``); the extensions are ``id``, ``element``, ``kind`` (docx4j's finer
kind, :class:`~docx4j_py.model.revisions.Kind`), ``target``, ``address``,
``part`` and ``pieces``.

**A view is valid until the next accept or reject on its package**, as Office
JS's is: a join removes a paragraph. A stale view's ``accept()`` and
``reject()`` raise :class:`~.errors.TrackedChangeError` with the code
``tracking.gone``; list again.
"""

from __future__ import annotations

import datetime
from typing import TYPE_CHECKING, Any

from docx4j_py.model.content.enums import TrackedChangeTypeValue
from docx4j_py.model.content.text_model import item_text
from docx4j_py.model.content.tracking import REVISION_NAMES
from docx4j_py.traversal import element_name, run_items_of
from docx4j_py.wml import R

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "TrackedChange",
    "accept_or_reject_body",
    "tracked_changes_of_body",
    "tracked_changes_of_package",
    "tracked_changes_of_paragraph",
]

#: How much of the text a ``repr`` shows.
_PREVIEW = 30


class TrackedChange:
    """A subset of Office JS ``Word.TrackedChange``, over one change the engine listed."""

    __slots__ = ("_paragraph", "body", "change")

    def __init__(self, change: Any, body: Body | None = None) -> None:
        """Build the view over the engine's change and the body it was listed from."""
        #: The engine's change (:class:`docx4j_py.model.revisions.TrackedChange`).
        self.change = change
        #: The body the change was listed from, for :attr:`paragraph` and
        #: :attr:`address`; None for a change in a text box, whose content is
        #: a typed island rather than the live tree (CR-010 section 3.5).
        self.body = body
        self._paragraph: Any = False

    # -- identity ----------------------------------------------------------

    def _anchors(self) -> tuple[int, ...]:
        return tuple(id(piece.target.value()) for piece in self.change.pieces)

    def __eq__(self, other: object) -> bool:
        """Two views of the same markup are equal (CR-003 section 3.1)."""
        return isinstance(other, TrackedChange) and other._anchors() == self._anchors()

    def __hash__(self) -> int:
        """Hashes by the markup's identity."""
        return hash(self._anchors())

    def __str__(self) -> str:
        """The text the change covers."""
        return self.text

    def __repr__(self) -> str:
        """``<TrackedChange Added 'Claude' body/3 'the new text'>``."""
        text = self.text
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        where = self.address
        return (
            f"<TrackedChange {self.type} {self.author!r}"
            f"{' ' + where if where else ''} {preview!r}>"
        )

    # -- Office JS's members ----------------------------------------------

    @property
    def type(self) -> TrackedChangeTypeValue:
        """Office JS ``Word.TrackedChangeType``: ``"Added"``, ``"Deleted"`` or ``"Formatted"``."""
        return self.change.type  # type: ignore[no-any-return]

    @property
    def author(self) -> str:
        """``w:author``: who made the change; a group's is its first piece's."""
        return str(self.change.author or "")

    @property
    def date(self) -> datetime.datetime | None:
        """When, in UTC: ``w16du:dateUtc`` where the markup has it, else ``w:date`` read as local time.

        ``w:date`` is Word's local wall-clock time whatever its ``Z`` says, read
        in this process's zone (CR-010 section 12.9). None when neither.
        """
        return self.change.date  # type: ignore[no-any-return]

    @property
    def text(self) -> str:
        """The text the change covers, as Office JS gives it.

        A deletion's is the **deleted** text; a paragraph mark is ``\\r``; a row
        its cells a tab apart and then ``\\r\\n``; a field's is its result.
        """
        return self.change.text  # type: ignore[no-any-return]

    def get_range(self) -> Range | None:
        """The range the change covers in its first paragraph, or None when it has none.

        A deletion is a **collapsed** range where it sits in the accepted text,
        because its text is not in that view at all; a mark, a ``w:pPrChange``
        or an equation's change is the whole paragraph; a row revision, a
        table's or a section's records and a text box's change have no range.
        A change over several paragraphs gives its range in the first.
        """
        from docx4j_py.model.content.range import Range
        from docx4j_py.model.revisions.change import (
            FieldT,
            HolderT,
            MarkPropsT,
            MarkT,
            RunPropsT,
            RunT,
        )

        paragraph = self.paragraph
        if paragraph is None:
            return None
        element = paragraph.element
        spans: list[tuple[int, int]] = []
        whole = False
        length = len(paragraph.text)
        for piece in self.change.pieces:
            if piece.para is None or piece.para.p is not element:
                continue
            t = piece.target
            if isinstance(t, RunT):
                spans.append(_span_of(paragraph, t.item))
            elif isinstance(t, FieldT):
                first, last = _span_of(paragraph, t.span[0]), _span_of(paragraph, t.span[-1])
                spans.append((first[0], last[1]))
            elif isinstance(t, HolderT):
                spans.append(_span_of(paragraph, t.holder))
            elif isinstance(t, RunPropsT):
                covered = [s for s in paragraph.segments() if s.run is t.run]
                at = _accepted_offset_of(element, t.run)
                spans.append((covered[0].start, covered[-1].end) if covered else (at, at))
            elif isinstance(t, (MarkT, MarkPropsT)):
                spans.append((length, length))
            else:
                whole = True
        if whole or not spans or all(isinstance(p.target, (MarkT, MarkPropsT)) for p in self.change.pieces):
            return paragraph.get_range("Whole")
        return Range(paragraph, min(s for s, _ in spans), max(e for _, e in spans))

    def accept(self) -> None:
        """Keep the change, as Word's Review tab does (docx4j CR-043 section 3.2).

        An insertion stays and loses its markup, a deletion goes (a deleted
        paragraph mark joins its paragraph with the next, which keeps its
        properties and ``w14:paraId``), a formatting change's record is
        dropped; either half of a move keeps the whole move. What goes takes
        with it the notes it references, the bookmarks and comments touching
        it, and the parts nothing else references.

        Raises:
            TrackedChangeError: ``tracking.gone`` when the package has been
                changed by an accept or reject since this was listed; list again.
        """
        self.change.accept()
        self._tidy()

    def reject(self) -> None:
        """Put back what was there: an insertion goes, a deletion stays and loses its markup.

        Raises:
            TrackedChangeError: ``tracking.gone``, as :meth:`accept`.
        """
        self.change.reject()
        self._tidy()

    def _tidy(self) -> None:
        """The halves of a run this session split, side by side again, joined (CR-003 section 16.12)."""
        body = self.body
        if body is not None and getattr(body.package, "_split_runs", None):
            body._tidy_runs()

    # -- the extensions ----------------------------------------------------

    @property
    def id(self) -> int | None:
        """The annotation id (``w:id``); a group's is its first piece's. Extension."""
        return self.change.id  # type: ignore[no-any-return]

    @property
    def kind(self) -> str:
        """docx4j's finer kind (``"run_insertion"``, ``"paragraph_mark_deletion"``, ``"move"``, ...). Extension.

        A :class:`~docx4j_py.model.revisions.Kind`, which is a ``str``; a
        group of pieces of different kinds is ``"insertion"``, ``"deletion"``
        or ``"formatting"``.
        """
        return self.change.kind  # type: ignore[no-any-return]

    @property
    def target(self) -> Any:
        """What the engine's change is over (its target). Extension."""
        return self.change.target

    @property
    def element(self) -> Any:
        """The markup this change is over: a ``w:ins``, a mark's record, a run, a row. Extension."""
        from docx4j_py.model.revisions.change import RowT, RunPropsT

        t = self.change.pieces[0].target
        if isinstance(t, RunPropsT):
            return t.run
        if isinstance(t, RowT):
            return t.tr
        return t.value()

    @property
    def part(self) -> Any:
        """The part the change is in: the main document part, a header, the notes, the comments. Extension."""
        return self.change.part

    @property
    def pieces(self) -> list[TrackedChange]:
        """The single changes this one groups, each a view that can be accepted or rejected on its own. Extension."""
        pieces = self.change.pieces
        if len(pieces) == 1 and pieces[0] is self.change:
            return [self]
        return [TrackedChange(piece, self.body) for piece in pieces]

    @property
    def paragraph(self) -> Paragraph | None:
        """The paragraph the change starts in, or None for a row, a table's or section's records, a text box's."""
        if self._paragraph is not False:
            return self._paragraph  # type: ignore[no-any-return]
        found = None
        para = self.change.para
        if para is not None and self.body is not None and self.change.story.island is None:
            from docx4j_py.model.content.paragraph import Paragraph

            found = Paragraph(para.p, para.container, self.body)
        self._paragraph = found
        return found

    @property
    def address(self) -> str:
        """The address of the paragraph (or the row, or the table) the change is in. Extension."""
        paragraph = self.paragraph
        if paragraph is not None:
            return paragraph.address
        if self.body is None or self.change.story.island is not None:
            return ""
        from docx4j_py.model.content.addresses import ordinal_of
        from docx4j_py.model.revisions.change import CellT, RowT, TableT

        t = self.change.pieces[0].target
        element = t.tr if isinstance(t, (RowT, CellT)) else t.tbl if isinstance(t, TableT) else None
        if element is None:
            return ""
        return ordinal_of(self.body, element) or ""

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about one revision."""
        date = self.date
        part = self.part
        return {
            "type": self.type,
            "author": self.author,
            "date": date.isoformat() if date is not None else None,
            "text": self.text,
            "id": self.id,
            "kind": str(self.kind),
            "address": self.address,
            "part": str(part.part_name) if part is not None and hasattr(part, "part_name") else None,
        }


# ---------------------------------------------------------------------------
# listing and resolving, over the engine
# ---------------------------------------------------------------------------


def _key(body: Body) -> Any:
    """What a body's views go stale with: its package, or its container when it has none."""
    return body.package if body.package is not None else body.container


def tracked_changes_of_body(body: Body) -> list[TrackedChange]:
    """Every tracked change in a body's story, in document order, and its text boxes'."""
    from docx4j_py.model import revisions

    changes = revisions.tracked_changes_of(body.package, revisions.body_stories(body), key=_key(body))
    return [TrackedChange(change, body) for change in changes]


def count_tracked_changes(body: Body) -> int:
    """How many changes :func:`tracked_changes_of_body` lists, without the views."""
    from docx4j_py.model import revisions
    from docx4j_py.model.revisions.lister import Lister

    return len(Lister(body.package, True, revisions.body_stories(body)).list())


def accept_or_reject_body(body: Body, accept: bool) -> int:
    """Resolve a body's story completely, and Word's side effects in other parts; the count listed."""
    from docx4j_py.model import revisions

    return revisions.resolve_all(body.package, accept, stories=revisions.body_stories(body), key=_key(body))


def tracked_changes_of_paragraph(paragraph: Paragraph) -> list[TrackedChange]:
    """The changes starting in one paragraph, cut at its end (Office JS ``Paragraph.getTrackedChanges``)."""
    from docx4j_py.model import revisions
    from docx4j_py.model.revisions.generation import Generation
    from docx4j_py.model.revisions.lister import Lister

    body = paragraph.parent_body
    stories = revisions.body_stories(body)
    story = _story_holding(stories, paragraph.element)
    lister = Lister(body.package, True, stories)
    changes = lister.paragraph_changes(paragraph.element, paragraph.container, story)
    generation = Generation.of(_key(body))
    out = []
    for change in changes:
        revisions._stamp(change, generation)
        view = TrackedChange(change, body)
        view._paragraph = paragraph
        out.append(view)
    return out


def _story_holding(stories: list, element: Any) -> Any:
    """The story whose root holds `element`, by its parents; the first when none says."""
    if len(stories) == 1:
        return stories[0]
    owners = {id(story.root_owner): story for story in stories}
    current = element
    for _ in range(64):
        if current is None:
            break
        found = owners.get(id(current))
        if found is not None:
            return found
        current = getattr(current, "parent", None)
    return stories[0]


def tracked_changes_of_package(package: Any) -> list[TrackedChange]:
    """Every tracked change in every story of a package (CR-010 section 3.2)."""
    from docx4j_py.model import revisions

    main = getattr(package, "main_document_part", None)
    bodies: dict[int, Any] = {}

    def body_for(part: Any) -> Any:
        if part is None:
            return None
        found = bodies.get(id(part))
        if found is None:
            found = package.body if part is main else getattr(part, "body", None)
            bodies[id(part)] = found
        return found

    out = []
    for change in revisions.tracked_changes(package):
        story = change.story
        body = None if story.island is not None else body_for(story.part)
        out.append(TrackedChange(change, body))
    return out


# ---------------------------------------------------------------------------
# where a revision sits in the accepted text
# ---------------------------------------------------------------------------


def _span_of(paragraph: Paragraph, element: Any) -> tuple[int, int]:
    """A run-level item's span in the paragraph's accepted text; collapsed where it is not in it."""
    from docx4j_py.model.content.tracking import revision_of

    inside = [
        segment
        for segment in paragraph.segments()
        if _within(segment.run, element)
        or ((found := revision_of(segment.run)) is not None and found.element is element)
    ]
    if inside:
        return inside[0].start, inside[-1].end
    at = _accepted_offset_of(paragraph.element, element)
    return at, at


def _within(run: Any, element: Any) -> bool:
    """True when `run` is `element` or stands somewhere under it."""
    current = run
    for _ in range(16):
        if current is None:
            return False
        if current is element:
            return True
        current = getattr(current, "parent", None)
    return False


def _accepted_offset_of(paragraph: Any, target: Any) -> int:
    """How much accepted text comes before `target` in the paragraph."""
    position = 0
    found = -1

    def visit(items: list | None) -> None:
        nonlocal position, found
        if not items or found >= 0:
            return
        for item in items:
            if found >= 0:
                return
            if item is target:
                found = position
                return
            if isinstance(item, R):
                position += len(_accepted_run_text(item))
                continue
            name = element_name(item)
            if name in REVISION_NAMES and (name or "").rpartition("}")[2] in ("del", "moveFrom"):
                continue
            visit(run_items_of(item))

    visit(run_items_of(paragraph))
    return found if found >= 0 else position


def _accepted_run_text(run: R) -> str:
    """One run's contribution to the accepted view (its ``w:delText`` left out)."""
    out: list[str] = []
    for child in getattr(run, "content", None) or ():
        text = item_text(element_name(child), child)
        if text is not None:
            out.append(text)
    return "".join(out)
