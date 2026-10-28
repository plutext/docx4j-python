"""One tracked change, as Word's Review tab and Office JS present it. docx4j ``TrackedChange``.

CR-010 section 3.1. Touching insertions, deletions or formatting changes by one
author are one change, whatever their dates; a move is one change from either
half. The engine's change is a **view**: it holds the markup it is over, and is
valid until the next accept or reject on its package (a paragraph join can
remove the paragraph it holds), after which :meth:`TrackedChange.accept` and
:meth:`TrackedChange.reject` raise
:class:`~docx4j_py.model.content.errors.TrackedChangeError` (``tracking.gone``).

The content API's :class:`~docx4j_py.model.content.tracked_change.TrackedChange`
is the view over this one (CR-010 section 3.2).
"""

from __future__ import annotations

import datetime
import enum
from typing import Any

from docx4j_py.model.revisions.trees import Revision

__all__ = [
    "Kind",
    "TrackedChange",
    "date_of",
    "utc_date",
    "word_date",
]


class Kind(enum.StrEnum):
    """docx4j's finer kind of change (``TrackedChange.Kind``), in snake case."""

    RUN_INSERTION = "run_insertion"
    """``w:ins`` around runs, at any depth of run holder; and inside an ``m:r``."""
    RUN_DELETION = "run_deletion"
    """``w:del`` around runs; and inside an ``m:r``."""
    MOVE = "move"
    """``w:moveFrom`` and ``w:moveTo``, with their range markers: either half is the whole move."""
    PARAGRAPH_MARK_INSERTION = "paragraph_mark_insertion"
    """``w:pPr/w:rPr/w:ins``: an inserted paragraph mark."""
    PARAGRAPH_MARK_DELETION = "paragraph_mark_deletion"
    """``w:pPr/w:rPr/w:del``: a deleted paragraph mark."""
    RUN_PROPERTIES = "run_properties"
    """A run's ``w:rPrChange``."""
    PARAGRAPH_MARK_PROPERTIES = "paragraph_mark_properties"
    """A paragraph mark's ``w:rPrChange`` (``w:pPr/w:rPr/w:rPrChange``)."""
    PARAGRAPH_PROPERTIES = "paragraph_properties"
    """``w:pPrChange``."""
    ROW_INSERTION = "row_insertion"
    """``w:trPr/w:ins``."""
    ROW_DELETION = "row_deletion"
    """``w:trPr/w:del``."""
    FIELD_INSERTION = "field_insertion"
    """A field Word inserted: ``w:ins`` around its begin and instruction and around its end."""
    CONTENT_CONTROL_INSERTION = "content_control_insertion"
    """A content control or custom XML element inserted (``w:customXmlInsRangeStart`` around it)."""
    CONTENT_CONTROL_DELETION = "content_control_deletion"
    """A content control or custom XML element deleted (``w:customXmlDelRangeStart`` around it).

    Word's form for a control a tracked deletion took whole (docx4j-core-ts
    check 54 cases 15 and 17; CR-013 Phase C): an extension of docx4j's kinds,
    as core-ts's ``CONTENT_CONTROL_DELETION`` is."""
    MATH_STRUCTURE_INSERTION = "math_structure_insertion"
    """An equation's structure (``m:sSup``, ``m:acc``, ...) inserted: a ``w:ins`` in its ``m:ctrlPr``."""
    MATH_STRUCTURE_DELETION = "math_structure_deletion"
    """An equation's structure deleted: a ``w:del`` in its ``m:ctrlPr``."""
    INSERTION = "insertion"
    """A group of insertions of different kinds (text, marks, rows, fields)."""
    DELETION = "deletion"
    """A group of deletions of different kinds."""
    FORMATTING = "formatting"
    """A group of run and mark formatting changes."""
    TABLE_PROPERTIES = "table_properties"
    """A table's property records (``w:tblPrChange``, ``w:tblGridChange``, ``w:trPrChange``,
    ``w:tcPrChange``, ``w:tblPrExChange``), which Word writes as one set: one change per table."""
    SECTION_PROPERTIES = "section_properties"
    """``w:sectPrChange``: a section's properties (the body's last section, or a paragraph's)."""
    CELL_INSERTION = "cell_insertion"
    """``w:cellIns`` (Word no longer writes it, but reads it)."""
    CELL_DELETION = "cell_deletion"
    """``w:cellDel`` (Word no longer writes it, but reads it)."""


class GroupKind(enum.Enum):
    """What a piece groups as."""

    INS = "ins"
    DEL = "del"
    FORMAT = "format"


# ---------------------------------------------------------------------------
# what a change is over
# ---------------------------------------------------------------------------


class Para:
    """A paragraph and the live list it stands in."""

    __slots__ = ("container", "p")

    def __init__(self, p: Any, container: list) -> None:
        """Hold the ``w:p`` and its list."""
        self.p = p
        self.container = container


class Target:
    """What a change is over: the markup, and where it stands."""

    __slots__ = ()

    def value(self) -> Any:
        """What carries ``w:id``, ``w:author`` and ``w:date``."""
        raise NotImplementedError


class RunT(Target):
    """``w:ins``, ``w:del``, ``w:moveFrom`` or ``w:moveTo`` around run-level content."""

    __slots__ = ("item", "owner", "revision")

    def __init__(self, revision: Revision, item: Any, owner: list) -> None:
        self.revision = revision
        self.item = item
        self.owner = owner

    def value(self) -> Any:
        return self.item


class MarkT(Target):
    """``w:pPr/w:rPr/w:ins`` or ``w:del``."""

    __slots__ = ("ins", "marker")

    def __init__(self, ins: bool, marker: Any) -> None:
        self.ins = ins
        self.marker = marker

    def value(self) -> Any:
        return self.marker


class RunPropsT(Target):
    """A run's ``w:rPrChange``."""

    __slots__ = ("record", "run")

    def __init__(self, run: Any, record: Any) -> None:
        self.run = run
        self.record = record

    def value(self) -> Any:
        return self.record


class ParaPropsT(Target):
    """A paragraph's ``w:pPrChange``."""

    __slots__ = ("record",)

    def __init__(self, record: Any) -> None:
        self.record = record

    def value(self) -> Any:
        return self.record


class MarkPropsT(Target):
    """A paragraph mark's ``w:rPrChange``."""

    __slots__ = ("record",)

    def __init__(self, record: Any) -> None:
        self.record = record

    def value(self) -> Any:
        return self.record


class RowT(Target):
    """A row's ``w:trPr/w:ins`` or ``w:del``, with the revisions of its own cells."""

    __slots__ = ("inner", "ins", "marker", "owner", "tbl", "tbl_owner", "tr")

    def __init__(
        self, ins: bool, tr: Any, owner: list, marker: Any, inner: list, tbl: Any, tbl_owner: list
    ) -> None:
        self.ins = ins
        self.tr = tr
        self.owner = owner
        self.marker = marker
        self.inner = inner
        self.tbl = tbl
        self.tbl_owner = tbl_owner

    def value(self) -> Any:
        return self.marker


class GroupT(Target):
    """Touching pieces of one type by one author: Office JS's one change."""

    __slots__ = ("kind", "pieces")

    def __init__(self, kind: GroupKind, pieces: list[TrackedChange]) -> None:
        self.kind = kind
        self.pieces = pieces

    def value(self) -> Any:
        return self.pieces[0].target.value()


class FieldT(Target):
    """A field Word inserted: the items of `owner` from the first ``w:ins`` to the last."""

    __slots__ = ("owner", "span")

    def __init__(self, span: list, owner: list) -> None:
        self.span = span
        self.owner = owner

    def value(self) -> Any:
        return self.span[0]


class HolderT(Target):
    """A content control or custom XML element inserted or deleted: its two pairs of ``w:customXmlIns*`` or ``Del*`` markers.

    `ins` is True for Word's form of an inserted control (check 30), False for
    a deleted one's (docx4j-core-ts check 54 cases 15 and 17; CR-013 Phase C).
    """

    __slots__ = ("holder", "inner_end", "inner_start", "ins", "outer_end", "outer_start", "owner")

    def __init__(
        self,
        holder: Any,
        owner: list,
        outer_start: Any,
        outer_end: Any,
        inner_end: Any,
        inner_start: Any,
        ins: bool = True,
    ) -> None:
        self.holder = holder
        self.owner = owner
        self.outer_start = outer_start
        self.outer_end = outer_end
        self.inner_end = inner_end
        self.inner_start = inner_start
        self.ins = ins

    def value(self) -> Any:
        return self.outer_start


class MathRunT(Target):
    """``w:ins`` or ``w:del`` inside an ``m:r``."""

    __slots__ = ("ins", "item", "run", "run_owner")

    def __init__(self, ins: bool, run: Any, item: Any, run_owner: list | None) -> None:
        self.ins = ins
        self.run = run
        self.item = item
        self.run_owner = run_owner

    def value(self) -> Any:
        return self.item


class MathStructT(Target):
    """A ``w:ins`` or ``w:del`` in an equation structure's ``m:ctrlPr``."""

    __slots__ = ("ctrl_pr", "ins", "marker", "struct_item", "struct_owner")

    def __init__(
        self, ins: bool, ctrl_pr: Any, marker: Any, struct_item: Any, struct_owner: list | None
    ) -> None:
        self.ins = ins
        self.ctrl_pr = ctrl_pr
        self.marker = marker
        self.struct_item = struct_item
        self.struct_owner = struct_owner

    def value(self) -> Any:
        return self.marker


class TableT(Target):
    """A table's property records."""

    __slots__ = ("record", "tbl")

    def __init__(self, tbl: Any, record: Any) -> None:
        self.tbl = tbl
        self.record = record

    def value(self) -> Any:
        return self.record


class SectionT(Target):
    """``w:sectPrChange``."""

    __slots__ = ("sect_pr", "text")

    def __init__(self, sect_pr: Any, text: str) -> None:
        self.sect_pr = sect_pr
        self.text = text

    def value(self) -> Any:
        return self.sect_pr.sect_pr_change


class CellT(Target):
    """``w:cellIns`` or ``w:cellDel``."""

    __slots__ = ("ins", "marker", "tc", "tr")

    def __init__(self, ins: bool, tc: Any, tr: Any, marker: Any) -> None:
        self.ins = ins
        self.tc = tc
        self.tr = tr
        self.marker = marker

    def value(self) -> Any:
        return self.marker


# ---------------------------------------------------------------------------
# the view
# ---------------------------------------------------------------------------


class TrackedChange:
    """One tracked change: the engine's, which the content API's view wraps."""

    __slots__ = ("generation", "para", "story", "target", "twin", "twins")

    def __init__(self, target: Target, para: Para | None, story: Any) -> None:
        """Build the change over its target, the paragraph it is in and its story."""
        #: What the change is over.
        self.target = target
        #: The paragraph the change is in; None for a row, a table's or a
        #: section's records, or markup at block level.
        self.para = para
        #: The story the change is in.
        self.story = story
        #: The changes in a text box's ``mc:Fallback`` twin this one resolves
        #: too (CR-043 section 3.1), and in the branches of a block- or
        #: paragraph-level ``mc:AlternateContent`` Word does not read (CR-012
        #: Phase A).
        self.twins: list[TrackedChange] = []
        #: True for a change in a branch of a block- or paragraph-level
        #: ``mc:AlternateContent`` Word does not read, listed on its own
        #: (``accept_all``'s listing, which resolves every branch): a twin, not
        #: counted.
        self.twin = False
        #: The package's generation when this was listed; None for a piece
        #: resolved through its group.
        self.generation: Any = None

    # -- Office JS's members ----------------------------------------------

    @property
    def type(self) -> str:
        """Office JS ``Word.TrackedChangeType``: ``"Added"``, ``"Deleted"`` or ``"Formatted"``."""
        t = self.target
        if isinstance(t, RunT):
            return "Added" if t.revision in (Revision.INS, Revision.MOVE_TO) else "Deleted"
        if isinstance(t, (MarkT, RowT, CellT, MathRunT, MathStructT)):
            return "Added" if t.ins else "Deleted"
        if isinstance(t, GroupT):
            if t.kind is GroupKind.INS:
                return "Added"
            return "Deleted" if t.kind is GroupKind.DEL else "Formatted"
        if isinstance(t, HolderT):
            return "Added" if t.ins else "Deleted"
        if isinstance(t, FieldT):
            return "Added"
        return "Formatted"

    @property
    def kind(self) -> Kind:
        """docx4j's finer kind; a group's is its pieces' when they are of one kind."""
        t = self.target
        if isinstance(t, RunT):
            if t.revision is Revision.INS:
                return Kind.RUN_INSERTION
            if t.revision is Revision.DEL:
                return Kind.RUN_DELETION
            return Kind.MOVE
        if isinstance(t, MarkT):
            return Kind.PARAGRAPH_MARK_INSERTION if t.ins else Kind.PARAGRAPH_MARK_DELETION
        if isinstance(t, RunPropsT):
            return Kind.RUN_PROPERTIES
        if isinstance(t, MarkPropsT):
            return Kind.PARAGRAPH_MARK_PROPERTIES
        if isinstance(t, ParaPropsT):
            return Kind.PARAGRAPH_PROPERTIES
        if isinstance(t, RowT):
            return Kind.ROW_INSERTION if t.ins else Kind.ROW_DELETION
        if isinstance(t, FieldT):
            return Kind.FIELD_INSERTION
        if isinstance(t, HolderT):
            return Kind.CONTENT_CONTROL_INSERTION if t.ins else Kind.CONTENT_CONTROL_DELETION
        if isinstance(t, MathRunT):
            return Kind.RUN_INSERTION if t.ins else Kind.RUN_DELETION
        if isinstance(t, MathStructT):
            return Kind.MATH_STRUCTURE_INSERTION if t.ins else Kind.MATH_STRUCTURE_DELETION
        if isinstance(t, TableT):
            return Kind.TABLE_PROPERTIES
        if isinstance(t, SectionT):
            return Kind.SECTION_PROPERTIES
        if isinstance(t, CellT):
            return Kind.CELL_INSERTION if t.ins else Kind.CELL_DELETION
        assert isinstance(t, GroupT)
        first = t.pieces[0].kind
        for piece in t.pieces:
            if piece.kind != first:
                if t.kind is GroupKind.INS:
                    return Kind.INSERTION
                return Kind.DELETION if t.kind is GroupKind.DEL else Kind.FORMATTING
        return first

    @property
    def author(self) -> str | None:
        """``w:author``; a group's is its first piece's (one author by definition)."""
        value = self.target.value()
        return getattr(value, "author", None) if value is not None else None

    @property
    def date(self) -> datetime.datetime | None:
        """When the change was made, in UTC: ``w16du:dateUtc`` where present, else ``w:date``.

        ``w:date`` is read as Word means it, local wall-clock time
        (:func:`date_of`), in the package's zone (``pkg.date_time_zone``, else
        the process's; CR-025, docx4j ``TrackedChange.getDate()`` since D12); a
        date alone as local midnight (CR-025 Phase B2). A group's is its first
        piece's; None when there is none.
        """
        from docx4j_py.model.dates import package_zone

        return date_of(self.target.value(), package_zone(getattr(self.story, "package", None)))

    @property
    def id(self) -> int | None:
        """The annotation id (``w:id``); a group's is its first piece's."""
        value = self.target.value()
        found = getattr(value, "id", None) if value is not None else None
        return found if isinstance(found, int) else None

    @property
    def text(self) -> str:
        """The text the change covers, as Office JS gives it.

        A paragraph mark is ``\\r``, a row its cells a tab apart and then
        ``\\r\\n``, a section's its text then ``\\f``; a deletion's is the
        deleted text; a field's is its result.
        """
        from docx4j_py.model.revisions.texts import text_of

        return text_of(self)

    @property
    def part(self) -> Any:
        """The part the change is in: the main document part, a header, the notes, the comments."""
        return self.story.part

    @property
    def pieces(self) -> list[TrackedChange]:
        """The single changes this one groups, in document order; a change that is no group is its own one piece."""
        if isinstance(self.target, GroupT):
            return list(self.target.pieces)
        return [self]

    def anchor(self) -> Any:
        """Where the change stands, for a range filter (docx4j ``TrackedChange.anchor``).

        The markup object its position is taken from: a ``w:ins``, a mark's
        record, a run, a row, a table, a section's properties; a group's is its
        first piece's.
        """
        t = self.target
        if isinstance(t, RunT):
            return t.item
        if isinstance(t, RunPropsT):
            return t.run
        if isinstance(t, RowT):
            return t.tr
        if isinstance(t, FieldT):
            return t.span[0]
        if isinstance(t, HolderT):
            return t.holder
        if isinstance(t, MathStructT):
            return t.ctrl_pr
        if isinstance(t, TableT):
            return t.tbl
        if isinstance(t, SectionT):
            return t.sect_pr
        if isinstance(t, CellT):
            return t.tc
        if isinstance(t, GroupT):
            return t.pieces[0].anchor()
        return t.value()

    def accept(self) -> None:
        """Keep the change, as Word's Review tab does.

        An insertion stays and loses its markup, a deletion goes (a deleted
        paragraph mark joins its paragraph with the next), a formatting change's
        record is dropped; either half of a move keeps the whole move. What goes
        takes with it the notes it references, the bookmarks and comments
        touching it, and the parts nothing else references (CR-043 section 3.2).

        Raises:
            TrackedChangeError: ``tracking.gone`` when the package has been
                changed by an accept or reject since this was listed.
        """
        self._resolve(True)

    def reject(self) -> None:
        """Put back what was there; either half of a move rejects the whole move."""
        self._resolve(False)

    def _resolve(self, accept: bool) -> None:
        from docx4j_py.model.content.errors import TrackedChangeError
        from docx4j_py.model.revisions.removals import Removals
        from docx4j_py.model.revisions.resolver import resolve

        if self.generation is None or not self.generation.is_current():
            raise TrackedChangeError(
                "this tracked change was listed before the last accept or reject on its "
                "package, so it cannot be "
                + ("accepted" if accept else "rejected"),
                code="tracking.gone",
                hint="call get_tracked_changes() again after accepting or rejecting others",
            )
        removals = Removals.for_story(self.story)
        resolve(self, accept, removals)
        for twin in self.twins:
            resolve(twin, accept, removals)
        removals.flush()
        self.generation.advance()

    def __repr__(self) -> str:
        """``<TrackedChange run_insertion by 'Author A' in /word/document.xml 'text'>``."""
        part = getattr(self.story.part, "part_name", None)
        text = self.text.replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
        return f"<TrackedChange {self.kind} by {self.author!r} in {part} {text!r}>"


# ---------------------------------------------------------------------------
# dates
# ---------------------------------------------------------------------------


def _calendar(value: Any) -> Any:
    """A ``w:date`` or a ``dateUtc`` as an ``XmlDateTime``, or None: the model's value, a string or a ``datetime``.

    :func:`docx4j_py.model.values.date_time`, Word's reading of the form: a
    date alone (``2024-02-03``, which the parse keeps as its text) is midnight
    on that day (CR-025 Phase B2; docx4j's date-forms probe, R01), as docx4j's
    ``XMLGregorianCalendar`` reads it.
    """
    from docx4j_py.model.values import date_time

    return date_time(value)


def _to_datetime(calendar: Any) -> datetime.datetime | None:
    try:
        return calendar.to_datetime()  # type: ignore[no-any-return]
    except Exception:  # noqa: BLE001 - a date we cannot read is no date
        return None


def utc_date(value: Any) -> datetime.datetime | None:
    """An ``xsd:dateTime`` meant as UTC --- ``w16du:dateUtc``, ``w16cex:dateUtc`` --- in UTC.

    A missing offset is read as UTC, as docx4j reads one; a date alone is
    midnight UTC, as docx4j's ``dateOf`` reads one (CR-025 Phase B2; Word
    writes neither).
    """
    calendar = _calendar(value)
    found = _to_datetime(calendar) if calendar is not None else None
    if found is None:
        return None
    if found.tzinfo is None:
        found = found.replace(tzinfo=datetime.UTC)
    return found.astimezone(datetime.UTC)


def word_date(value: Any, zone: datetime.tzinfo | None = None) -> datetime.datetime | None:
    """A ``w:date`` as Word means it, in UTC: its local wall-clock time, the ``Z`` notwithstanding.

    Word writes ``w:date`` on the author's local wall clock with a ``Z`` that
    does not mean UTC, and the true UTC beside it in ``w16du:dateUtc`` (Word
    365) or ``w16cex:dateUtc`` (a comment). Without those, the wall-clock time is
    read in `zone`, the package's (``pkg.date_time_zone``; CR-025, docx4j's
    ``TrackedChange.dateOf(value, zone)`` since D12), or for None in **this
    process's** time zone, as docx4j reads it in the JVM's by default (check
    36) and Office JS in the add-in's. An explicit non-zero offset, which Word
    does not write, is read as it says (CR-010 section 12.9).

    Word reads every form so (docx4j CR-067 section 12.2, its date-forms
    probe, Word 365 in Sydney): ``2025-01-02T03:04:05Z`` as 3:04 AM local, the
    ``Z`` ignored; ``2025-06-07T08:09:10``, with no offset, as 8:09 AM local, as
    written; and a date alone, ``2024-02-03``, as **local midnight**
    (:func:`docx4j_py.model.values.date_time`; CR-025 Phase B2: it was None).
    """
    calendar = _calendar(value)
    found = _to_datetime(calendar) if calendar is not None else None
    if found is None:
        return None
    if calendar.offset not in (None, 0):
        return found.astimezone(datetime.UTC)
    from docx4j_py.model.dates import instant_of

    wall_clock = found.replace(tzinfo=None)
    try:
        return instant_of(wall_clock, zone)
    except (OverflowError, OSError, ValueError):  # pragma: no cover - out of the platform's range
        return wall_clock.replace(tzinfo=datetime.UTC)


def date_of(value: Any, zone: datetime.tzinfo | None = None) -> datetime.datetime | None:
    """A revision's date, in UTC: ``w16du:dateUtc`` where present (Word 365 writes both), else ``w:date``.

    docx4j ``TrackedChange.dateOf`` (CR-010 section 12.9): ``w16du:dateUtc`` is
    the true UTC, whatever `zone`; ``w:date`` is Word's local wall-clock time,
    read in `zone`, else this process's time zone (:func:`word_date`). This
    package's tracker writes both, as Word does, so what it writes reads back as
    the same instant.
    """
    if value is None:
        return None
    utc = utc_date(getattr(value, "date_utc", None))
    if utc is not None:
        return utc
    return word_date(getattr(value, "date", None), zone)
