"""The change tracker: Word's revision markup, written by every mutation.

CR-003 section 3.8 and section 4's "Tracking rules, Word's not just the markup",
Phase F. This is the writing half; :mod:`.tracked_change` is the reading half
(``TrackedChange``, ``accept()`` and ``reject()``).

::

    pkg.author = Author("Claude", initials="C")
    pkg.change_tracking_mode = "TrackAll"
    pkg.body.paragraphs[0].replace_text("colour", "color")   # a w:del and a w:ins
    pkg.body.get_tracked_changes()

What the markup is (ECMA-376 17.13.5): ``w:ins`` and ``w:del`` around runs, with
a deleted run's ``w:t`` turned into ``w:delText`` and its ``w:instrText`` into
``w:delInstrText``; ``w:pPr/w:rPr/w:ins`` and ``w:pPr/w:rPr/w:del`` for a
paragraph mark; ``w:trPr/w:ins`` and ``w:trPr/w:del`` for a table row;
``w:rPrChange`` and ``w:pPrChange`` for formatting, each recorded **once**, on
the first write, with the properties as they stood before it.

The rules of section 4 that land here:

* a run already inside a ``w:ins`` **by the same author** is extended rather
  than nested in another one, and deleting text that author inserted takes it
  back rather than nesting a ``w:del``;
* a replacement is the ``w:del`` first and the ``w:ins`` after it;
* runs are split at the span's boundaries, so a partly deleted run is not
  wholly deleted;
* revision ids come from a per-package counter one above the highest ``w:id``
  on any ``CTMarkup`` in the parts already unmarshalled --- the one annotation
  id space of 17.13.5.4 --- **excluding** ``w:comment/@w:id``, which is Phase
  G's separate space (:func:`~docx4j_py.model.content.comments.next_comment_id`);
* the author is ``pkg.author`` (:func:`.comments.author_of`), the same identity
  a comment carries, and only its ``name`` reaches ``w:author``;
* ``TrackMineOnly`` is stored as ``TrackAll``: ``w:trackRevisions`` is a flag
  and a file cannot tell the two apart;
* editing deleted text raises :class:`~.errors.TrackedChangeError`, naming the
  author who deleted it.
"""

from __future__ import annotations

import dataclasses
import datetime
from typing import Any

from docx4j_py.child import ChildList, deep_copy, deep_copy_as, link_parents
from docx4j_py.model.content.enums import (
    CHANGE_TRACKING_MODES,
    ChangeTracking,
    ChangeTrackingModeValue,
)
from docx4j_py.model.content.errors import ContentError, TrackedChangeError
from docx4j_py.model.content.text_model import RUN_HOLDER_NAMES
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of
from docx4j_py.wml import (
    CTMarkup,
    CTPPrChange,
    CTRPrChange,
    CTTrackChange,
    P,
    ParaRPr,
    PPrBase,
    R,
    RPr,
    Tbl,
    Tr,
    TrPr,
    el,
    rpr_to_elements,
)

__all__ = [
    "CHANGE_TRACKING_MODES",
    "REVISION_NAMES",
    "RUN_ATTRIBUTES",
    "ChangeTracker",
    "ChangeTracking",
    "ChangeTrackingModeValue",
    "FontTracking",
    "Revision",
    "copy_r_pr",
    "create_settings_part",
    "date_of",
    "delete_row_content",
    "insertion_mark_of",
    "mark_deleted",
    "mark_inserted",
    "merge_runs",
    "mode_of",
    "para_r_pr_of",
    "prune_paragraph_properties",
    "restore_p_pr",
    "restore_r_pr",
    "revision_of",
    "row_pr_of",
    "set_mode",
    "settings_part_for_write",
    "tidy_runs",
    "to_deleted_text",
    "to_restored_text",
    "track_deleted_row",
    "track_inserted_blocks",
    "track_inserted_paragraph",
    "track_inserted_paragraph_into",
    "track_inserted_table",
    "tracker_of",
    "wrap_new_runs",
    "wrap_run",
    "xml_date",
]


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


W_INS = _w("ins")
W_DEL = _w("del")
W_MOVE_FROM = _w("moveFrom")
W_MOVE_TO = _w("moveTo")
W_TRACK_REVISIONS = _w("trackRevisions")

#: The four run-level revision holders, by element name.
REVISION_NAMES: frozenset[str] = frozenset({W_INS, W_DEL, W_MOVE_FROM, W_MOVE_TO})

#: A ``w:r``'s own attributes: Word's revision save ids. A run split in two
#: gives both halves the original's, so that the halves are identical but for
#: their text --- which is what lets a rejected revision put them back together
#: (CR-003 section 16.12).
RUN_ATTRIBUTES: tuple[str, ...] = ("rsid_rpr", "rsid_del", "rsid_r")

#: What a deleted run's text items become, and the way back.
_DELETED_NAMES: dict[str, str] = {"t": "delText", "instrText": "delInstrText"}
_RESTORED_NAMES: dict[str, str] = {"delText": "t", "delInstrText": "instrText"}


# ---------------------------------------------------------------------------
# the revision a run is in
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class Revision:
    """The ``w:ins`` / ``w:del`` / ``w:moveFrom`` / ``w:moveTo`` a run sits in."""

    kind: str
    """``"ins"``, ``"del"``, ``"moveFrom"`` or ``"moveTo"``."""
    element: Any
    """The revision element: a ``RunIns``, ``RunDel``, ``MoveFrom2``, ``MoveTo2``."""
    owner: list
    """The live list holding :attr:`element`."""
    items: list
    """The revision's own run list."""
    parent: Any
    """What :attr:`element` hangs from: the ``w:p``, a hyperlink, ..."""

    @property
    def author(self) -> str:
        """``w:author`` on the revision, or ``""``."""
        return str(getattr(self.element, "author", "") or "")

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the kind, the author and the id."""
        return {"kind": self.kind, "author": self.author, "id": getattr(self.element, "id", None)}


def revision_of(run: Any) -> Revision | None:
    """The revision a run is inside, hyperlinks and controls seen through.

    The run's parent is the holder :func:`~docx4j_py.traversal.run_items_of`
    put it in, so this walks up through the run holders (a ``w:hyperlink``, a
    run-level ``w:sdt``) until it meets a revision or something that is not a
    holder at all.
    """
    holder = getattr(run, "parent", None)
    while holder is not None:
        name = element_name(holder)
        if name in REVISION_NAMES:
            outer = getattr(holder, "parent", None)
            owner = run_items_of(outer) if outer is not None else None
            return Revision(
                kind=(name or "").rpartition("}")[2],
                element=holder,
                owner=owner if owner is not None else [],
                items=run_items_of(holder) or [],
                parent=outer,
            )
        if name in RUN_HOLDER_NAMES:
            holder = getattr(holder, "parent", None)
            continue
        return None
    return None


# ---------------------------------------------------------------------------
# dates, as Word writes w:date
# ---------------------------------------------------------------------------


def xml_date(value: datetime.datetime) -> Any:
    """A ``w:date`` value: an ``xsd:dateTime`` in UTC, to the second, as Word writes it."""
    from docx4j_xsdata.models.datatype import XmlDateTime

    if value.tzinfo is None:
        value = value.replace(tzinfo=datetime.UTC)
    return XmlDateTime.from_datetime(value.astimezone(datetime.UTC).replace(microsecond=0))


def date_of(value: Any) -> datetime.datetime | None:
    """The :class:`datetime.datetime` of a ``w:date``, or None when there is none."""
    if value is None:
        return None
    to_datetime = getattr(value, "to_datetime", None)
    if to_datetime is not None:
        try:
            return to_datetime()
        except Exception:  # noqa: BLE001 - a date we cannot read is no date
            return None
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, str):
        try:
            return datetime.datetime.fromisoformat(value)
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# the tracker
# ---------------------------------------------------------------------------


class ChangeTracker:
    """The per-package writer of revision markup; None-shaped when the mode is off.

    Made by :func:`tracker_of` while ``pkg.change_tracking_mode`` is not
    ``"Off"``; every mutation of the content API asks its
    :class:`~docx4j_py.model.content.body.Body` for one and, when it gets one,
    writes revisions instead of editing in place.

    A departure from docx4j, which has no such class --- docx4j leaves revision
    markup to the caller. The name follows Office JS's vocabulary
    (``Word.ChangeTrackingMode``).
    """

    __slots__ = ("_counter", "mode", "package")

    def __init__(self, package: Any, mode: str = "TrackAll") -> None:
        """Build the tracker over a package and the mode it is in."""
        #: The package whose author, date and ids this writes.
        self.package = package
        #: ``"TrackAll"`` or ``"TrackMineOnly"``; never ``"Off"``.
        self.mode = mode
        self._counter: int | None = None

    def __repr__(self) -> str:
        """``<ChangeTracker TrackAll 'Claude'>``."""
        return f"<ChangeTracker {self.mode} {self.author!r}>"

    # -- who and when ------------------------------------------------------

    @property
    def author(self) -> str:
        """The ``w:author`` a new revision carries: ``pkg.author``'s name."""
        from docx4j_py.model.content.comments import author_of

        return author_of(self.package).name

    @property
    def date(self) -> datetime.datetime:
        """``pkg.tracked_change_date``, or now in UTC (CR-003 section 3.8)."""
        fixed = getattr(self.package, "tracked_change_date", None)
        if isinstance(fixed, datetime.datetime):
            return fixed
        return datetime.datetime.now(datetime.UTC)

    # -- the ids -----------------------------------------------------------

    def markup_roots(self) -> list[Any]:
        """The trees the id scan covers: the parts already unmarshalled.

        A part this package has not read is not unmarshalled for an id, which
        would cost a whole document's parse and mark the part for
        re-marshalling. The annotation ids of a header nobody has touched
        cannot collide with an edit to the body in any way Word minds.
        """
        from docx4j_py.model.content.addresses import package_parts

        out: list[Any] = []
        for part in package_parts(self.package):
            if getattr(part, "is_unmarshalled", False):
                contents = getattr(part, "contents", None)
                if contents is not None:
                    out.append(contents)
        return out

    def next_id(self) -> int:
        """The next annotation id: one above the highest in use, then counting.

        The one annotation id space of ECMA-376 17.13.5.4 --- revisions,
        bookmarks, comment range marks and permissions share it --- scanned
        once per package. ``w:comment/@w:id`` is **excluded**: it is a different
        space, and Phase G's :func:`.comments.next_comment_id` owns it.
        """
        if self._counter is None:
            from docx4j_py.traversal import iter_nodes
            from docx4j_py.wml import CommentsComment

            highest = 0
            for root in self.markup_roots():
                for node in iter_nodes(root, mce="all"):
                    if not isinstance(node, CTMarkup) or isinstance(node, CommentsComment):
                        continue
                    identifier = getattr(node, "id", None)
                    if isinstance(identifier, int) and identifier > highest:
                        highest = identifier
            self._counter = highest
        self._counter += 1
        return self._counter

    def markup(self) -> dict[str, Any]:
        """The ``w:id``, ``w:author`` and ``w:date`` a new revision carries."""
        return {"id": self.next_id(), "author": self.author, "date": xml_date(self.date)}

    def track_change(self) -> CTTrackChange:
        """A ``CT_TrackChange``: the id, this package's author and the date."""
        return CTTrackChange(**self.markup())

    # -- the run-level wrappers -------------------------------------------

    def ins(self, items: list) -> Any:
        """``w:ins`` around run-level items, their parents linked."""
        wrapper = el.ins(content=ChildList(list(items)), **self.markup())
        link_parents(wrapper)
        for item in items:
            item.parent = wrapper
        return wrapper

    def deletion(self, items: list) -> Any:
        """``w:del`` around run-level items whose text is already ``w:delText``."""
        wrapper = el.del_(content=ChildList(list(items)), **self.markup())
        link_parents(wrapper)
        for item in items:
            item.parent = wrapper
        return wrapper

    # -- the rules ---------------------------------------------------------

    def own_insertion(self, revision: Revision | None) -> bool:
        """True when this is an insertion by this author, to extend rather than nest."""
        return revision is not None and revision.kind == "ins" and revision.author == self.author

    def own_paragraph(self, paragraph: P) -> bool:
        """True when a paragraph's **mark** is an insertion by this author."""
        p_pr = paragraph.p_pr
        r_pr = p_pr.r_pr if p_pr is not None else None
        ins = r_pr.ins if r_pr is not None else None
        return ins is not None and str(getattr(ins, "author", "") or "") == self.author

    def inserted(self, paragraph: P) -> bool:
        """True when **this session** inserted the paragraph (CR-003 section 16.12).

        Its properties are this author's too, so a property write on it records
        no ``w:pPrChange``: there is nothing in the document yet to record. The
        mark cannot answer this since 16.10 --- an inserted paragraph at the end
        of a container carries none, and the mark it does carry may belong to
        the paragraph before it --- so the package remembers.
        """
        store = getattr(self.package, "_inserted_paragraphs", None)
        return bool(store) and id(paragraph) in store

    def remember_inserted(self, paragraph: P) -> None:
        """Record that this session inserted a paragraph."""
        store = getattr(self.package, "_inserted_paragraphs", None)
        if store is not None:
            store.add(id(paragraph))

    def assert_editable(self, revision: Revision | None) -> None:
        """Refuse an edit to deleted text, naming the author who deleted it.

        Raises:
            TrackedChangeError: the text is inside a ``w:del`` or a
                ``w:moveFrom``.
        """
        if revision is not None and revision.kind in ("del", "moveFrom"):
            raise TrackedChangeError(
                f"this text is inside a w:{revision.kind} "
                f"(deleted by {revision.author or 'another author'})",
                code="tracking.deleted_text",
                hint="reject the deletion first, then edit the text",
            )

    # -- the recorded originals -------------------------------------------

    def record_r_pr_change(self, r_pr: RPr) -> None:
        """Record a run's properties in ``w:rPrChange``, once (the first original wins)."""
        if getattr(r_pr, "r_pr_change", None) is not None:
            return
        from docx4j_py.wml import CtRprChangeRPr, rpr_from_elements

        original = rpr_from_elements(rpr_to_elements(r_pr), cls=CtRprChangeRPr)
        change = CTRPrChange(r_pr=original, **self.markup())
        r_pr.r_pr_change = change
        link_parents(change)
        change.parent = r_pr

    def record_p_pr_change(self, p_pr: Any) -> None:
        """Record a paragraph's properties in ``w:pPrChange``, once.

        The copy is made **as** ``PPrBase``, which is what ``w:pPrChange``
        declares: a ``PPr`` in there would marshal ``xsi:type="w:CT_PPr"``,
        which is valid and is not what Word writes (CR-003 section 10.2).
        """
        if getattr(p_pr, "p_pr_change", None) is not None:
            return
        original = deep_copy_as(p_pr, PPrBase)
        change = CTPPrChange(p_pr=original, **self.markup())
        p_pr.p_pr_change = change
        link_parents(change)
        change.parent = p_pr

    # -- the marks ---------------------------------------------------------

    def mark_paragraph_inserted(self, paragraph: P) -> None:
        """``w:pPr/w:rPr/w:ins``: the paragraph mark is an insertion."""
        r_pr = para_r_pr_of(paragraph)
        r_pr.ins = self.track_change()
        r_pr.ins.parent = r_pr

    def mark_paragraph_deleted(self, paragraph: P) -> None:
        """``w:pPr/w:rPr/w:del``: the mark is deleted; accepting joins with the next.

        Raises:
            TrackedChangeError: the mark is already marked deleted.
        """
        r_pr = para_r_pr_of(paragraph)
        if r_pr.del_value is not None:
            raise TrackedChangeError(
                "this paragraph mark is already marked deleted",
                code="tracking.already_deleted",
                hint="accept or reject the deletion before deleting the paragraph again",
            )
        r_pr.del_value = self.track_change()
        r_pr.del_value.parent = r_pr

    def mark_row_inserted(self, row: Tr) -> None:
        """``w:trPr/w:ins``: the table row is an insertion."""
        tr_pr = row_pr_of(row)
        tr_pr.ins = self.track_change()
        tr_pr.ins.parent = tr_pr

    def mark_row_deleted(self, row: Tr) -> None:
        """``w:trPr/w:del``: the row is deleted, and stays in the tree until accepted.

        Raises:
            TrackedChangeError: the row is already marked deleted.
        """
        tr_pr = row_pr_of(row)
        if tr_pr.del_value is not None:
            raise TrackedChangeError(
                "this row is already marked deleted",
                code="tracking.already_deleted",
                hint="accept or reject the deletion before deleting the row again",
            )
        tr_pr.del_value = self.track_change()
        tr_pr.del_value.parent = tr_pr


@dataclasses.dataclass(frozen=True, slots=True)
class FontTracking:
    """What a :class:`~docx4j_py.model.content.font.Font` needs to record a change.

    The tracker, and the runs that are **this author's own insertion**: a run
    inside a ``w:ins`` of ours takes the formatting outright, because there is
    nothing yet in the document to record as the original (CR-003 section 4).
    """

    tracker: ChangeTracker
    """The package's tracker."""
    own_insertions: frozenset[int]
    """``id()`` of each run inside a ``w:ins`` by this author."""

    def records(self, run: Any) -> bool:
        """True when a write to this run must record ``w:rPrChange`` first."""
        return id(run) not in self.own_insertions


def tracker_of(package: Any) -> ChangeTracker | None:
    """The package's tracker while its mode is on, or None when it is ``"Off"``.

    Cached on the package, so that a loop of edits reads the settings part once
    and scans for the highest annotation id once.
    """
    if package is None:
        return None
    found = getattr(package, "_change_tracker", None)
    if found is not None:
        return found
    # through the package's own property where it has one, because a
    # ``TrialPackage``'s answers about the trial without unmarshalling the real
    # document's settings part (CR-003 section 16)
    mode = getattr(package, "change_tracking_mode", None) or mode_of(package)
    if mode == "Off":
        return None
    tracker = ChangeTracker(package, mode)
    try:
        package._change_tracker = tracker
    except AttributeError:  # pragma: no cover - a package-like object with no slot
        return tracker
    return tracker


# ---------------------------------------------------------------------------
# whole paragraphs and whole tables
# ---------------------------------------------------------------------------


def track_inserted_paragraph(tracker: ChangeTracker, paragraph: P) -> None:
    """A paragraph inserted whole: its mark is marked inserted and its runs wrapped.

    Word, splitting a paragraph, marks the **first** paragraph's mark instead
    and leaves the new one the original mark; marking the new paragraph's own
    mark gives the same document once accepted or rejected and keeps the edit on
    the element that was added (CR-003 section 4).
    """
    tracker.remember_inserted(paragraph)
    tracker.mark_paragraph_inserted(paragraph)
    wrap_new_runs(tracker, paragraph)


def wrap_new_runs(tracker: ChangeTracker, paragraph: P) -> None:
    """Wrap every top-level run not already in a revision, in one ``w:ins`` per group."""
    content = getattr(paragraph, "content", None)
    if not content:
        return
    index = 0
    while index < len(content):
        if not _is_plain_run(content[index]):
            index += 1
            continue
        end = index
        while end < len(content) and _is_plain_run(content[end]):
            end += 1
        runs = list(content[index:end])
        del content[index:end]
        wrapper = tracker.ins(runs)
        content.insert(index, wrapper)
        wrapper.parent = paragraph
        index += 1


def _index_of(items: list, element: Any) -> int:
    """The index of an element **by identity**; ``-1`` when it is not there."""
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def wrap_run(tracker: ChangeTracker, run: R, paragraph: P) -> Any:
    """Wrap **one** run in a ``w:ins`` where it stands, and return the wrapper.

    What a verb that adds a single run to an existing paragraph needs ---
    :meth:`~docx4j_py.model.content.paragraph.Paragraph.insert_inline_picture`
    is the one. :func:`wrap_new_runs` would wrap the paragraph's **own** runs
    too, and rejecting the revision would then take text nobody inserted
    (CR-003 section 16.12).
    """
    holder = getattr(run, "parent", None)
    items = run_items_of(holder) if holder is not None else None
    if items is None:
        items, holder = paragraph.content, paragraph
    index = _index_of(items, run)
    if index < 0:
        return None
    del items[index]
    wrapper = tracker.ins([run])
    items.insert(index, wrapper)
    wrapper.parent = holder
    return wrapper


def _is_plain_run(item: Any) -> bool:
    """A ``w:r`` that is not itself a revision holder's child of another author."""
    return isinstance(item, R)


def track_inserted_blocks(tracker: ChangeTracker, elements: list, container: list) -> None:
    """Mark what one insert put into a container, the way **Word** marks it.

    In the middle of a container each new paragraph carries its own mark, as
    docx4j-core-ts's phase F decided (CR-003 section 16.2 item, corrected in
    16.10). At the **end** of one it must not: a container's final paragraph
    mark cannot be deleted, so Word never marks it inserted --- pressing Enter
    at the end of the last paragraph marks the *preceding* paragraph's mark and
    gives the new paragraph the original, unmarked final mark. A final mark
    marked inserted is something Word cannot reject, and it hangs.

    So a paragraph with nothing after it but the other paragraphs of this same
    insert **shifts its mark back one**: the paragraph before it takes it, which
    for a fragment of several is the one this call inserted just before. With no
    paragraph before it --- an empty container, or a table --- nothing is marked
    and only the runs are wrapped, which is the honest answer: there is no
    earlier mark for the break to live on.
    """
    inserted = {id(item) for item in elements}
    for item in elements:
        if isinstance(item, Tbl):
            track_inserted_table(tracker, item)
            continue
        if not isinstance(item, P):
            continue
        tracker.remember_inserted(item)
        wrap_new_runs(tracker, item)
        index = _index_of(container, item)
        if index < 0:
            tracker.mark_paragraph_inserted(item)
            continue
        if not _trails(container, index, inserted):
            tracker.mark_paragraph_inserted(item)
            continue
        previous = container[index - 1] if index > 0 else None
        if isinstance(previous, P) and not mark_inserted(previous):
            tracker.mark_paragraph_inserted(previous)


def _trails(container: list, index: int, inserted: set[int]) -> bool:
    """True when nothing follows ``container[index]`` but this insert's own paragraphs."""
    for item in container[index + 1 :]:
        if not isinstance(item, P) or id(item) not in inserted:
            return False
    return True


def track_inserted_paragraph_into(
    tracker: ChangeTracker, paragraph: P, container: list
) -> None:
    """One inserted paragraph: :func:`track_inserted_blocks` for a single element."""
    track_inserted_blocks(tracker, [paragraph], container)


def insertion_mark_of(paragraph: P, container: list) -> P | None:
    """The paragraph whose mark records **this** paragraph's insertion, or None.

    Its own mark normally; the **previous** paragraph's when this is the last of
    its container, which is where :func:`track_inserted_paragraph_into` puts it.
    """
    if mark_inserted(paragraph):
        return paragraph
    index = _index_of(container, paragraph)
    if index > 0 and index == len(container) - 1:
        previous = container[index - 1]
        if isinstance(previous, P) and mark_inserted(previous):
            return previous
    return None


def track_deleted_row(tracker: ChangeTracker, row: Tr, body: Any) -> bool:
    """Delete a row as Word deletes one: the row mark **and its content**.

    Word writes a deleted row as ``w:trPr/w:del`` *plus* every run in every cell
    in a ``w:del`` with ``w:delText`` and every cell paragraph's mark marked
    deleted --- the mirror of an inserted row. A row marked but not emptied is
    what Word shows in pink with no strikethrough (CR-003 section 16.10).

    A row **this author inserted** is taken back instead, as deleting text this
    author inserted is (section 4).

    Returns:
        True when the row was taken back and the caller should remove it.
    """
    tr_pr = getattr(row, "tr_pr", None)
    ins = tr_pr.ins if tr_pr is not None else None
    if ins is not None and str(getattr(ins, "author", "") or "") == tracker.author:
        return True
    tracker.mark_row_deleted(row)
    delete_row_content(tracker, row, body)
    return False


def delete_row_content(tracker: ChangeTracker, row: Tr, body: Any) -> None:
    """Every run of every cell into a ``w:del``, every cell mark marked deleted."""
    from docx4j_py.model.content.text_model import block_children_of, cells_of

    for cell, _owner in cells_of(row):
        blocks = block_children_of(cell)
        if blocks is not None:
            _delete_blocks(tracker, blocks, body)


def _delete_blocks(tracker: ChangeTracker, blocks: list, body: Any) -> None:
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.text_model import block_children_of, rows_of

    for block in blocks:
        if isinstance(block, P):
            view = Paragraph(block, blocks, body)
            length = len(view.text)
            if length:
                view.delete_text_tracked(tracker, 0, length)
            if not mark_deleted(block):
                tracker.mark_paragraph_deleted(block)
            continue
        if isinstance(block, Tbl):
            for nested, _owner in rows_of(block):
                if getattr(getattr(nested, "tr_pr", None), "del_value", None) is None:
                    track_deleted_row(tracker, nested, body)
            continue
        children = block_children_of(block)
        if children is not None:
            _delete_blocks(tracker, children, body)


def track_inserted_table(tracker: ChangeTracker, table: Tbl) -> None:
    """A table inserted whole: every row marked inserted, every paragraph an insertion."""
    from docx4j_py.model.content.text_model import cells_of, rows_of

    for row, _owner in rows_of(table):
        tracker.mark_row_inserted(row)
        for cell, _cell_owner in cells_of(row):
            _track_blocks(tracker, getattr(cell, "content", None) or [])


def _track_blocks(tracker: ChangeTracker, blocks: list) -> None:
    from docx4j_py.model.content.text_model import block_children_of

    for block in blocks:
        if isinstance(block, P):
            track_inserted_paragraph(tracker, block)
            continue
        if isinstance(block, Tbl):
            track_inserted_table(tracker, block)
            continue
        children = block_children_of(block)
        if children is not None:
            _track_blocks(tracker, children)


# ---------------------------------------------------------------------------
# the property holders
# ---------------------------------------------------------------------------


def merge_runs(
    items: list, at: int, split: dict[int, int] | None = None, spaces: set[int] | None = None
) -> int:
    """Join the runs on either side of one boundary in a run list; how many went.

    A revision that goes --- a ``w:del`` accepted, a ``w:ins`` rejected, a
    deletion restored --- leaves the two halves of a run **this API split** side
    by side: isolating ``document`` for a ``w:del`` splits
    ``<w:t>document.</w:t>`` in two, and putting the text back does not put the
    run back. Word merges them, and so does this, so that a document with every
    revision rejected is **canonically the document that was there** (CR-003
    section 16.12).

    It is deliberately **narrow**. Only the boundary a revision just vacated is
    looked at; the two runs must be identical but for their content --- the same
    ``w:rPr`` by value and the same ``w:rsidR`` / ``w:rsidRPr`` / ``w:rsidDel``,
    which a split gives both halves; and they must be **the two halves of one
    split this package made** (`split` is its ``_split_runs``, ``id(tail) ->
    id(head)``). Runs a document keeps apart stay apart even when they are
    identical, which they often are, and a half is never merged with a run it
    never belonged to. With no `split` --- a package reloaded since, which
    cannot know --- nothing is merged at all.
    """
    merged = 0
    while 0 < at < len(items) and _mergeable(items[at - 1], items[at], split):
        first, second = items[at - 1], items[at]
        for item in list(second.content):
            first.content.append(item)
            item.parent = first
        _join_text(first, spaces)
        del items[at]
        _joined(split, first, second)
        merged += 1
    return merged


def _joined(split: dict[int, int] | None, first: R, second: R) -> None:
    """Keep the split table right after two halves become one run."""
    if not split:
        return
    split.pop(id(second), None)
    for tail, head in list(split.items()):
        if head == id(second):
            split[tail] = id(first)


def tidy_runs(
    paragraph: P, split: dict[int, int] | None, spaces: set[int] | None = None
) -> int:
    """:func:`merge_runs` at every boundary of a paragraph; how many runs went.

    Run once after ``accept_all()`` / ``reject_all()``, because the order the
    changes come in decides what is mergeable **when**: a run whose
    ``w:rPrChange`` is rejected after its neighbour's deletion could not be
    joined at the moment that deletion went (CR-003 section 16.12). It is as
    narrow as :func:`merge_runs`: one of the two runs must be a run this
    session's split made.
    """
    merged = 0

    def visit(items: list | None) -> None:
        nonlocal merged
        if not items:
            return
        index = 1
        while index < len(items):
            gone = merge_runs(items, index, split, spaces)
            merged += gone
            if gone:
                continue
            item = items[index - 1]
            if not isinstance(item, R):
                nested = run_items_of(item)
                if nested is not None:
                    visit(nested)
            index += 1
        last = items[-1] if items else None
        if last is not None and not isinstance(last, R):
            nested = run_items_of(last)
            if nested is not None:
                visit(nested)

    visit(run_items_of(paragraph))
    return merged


def _mergeable(first: Any, second: Any, split: dict[int, int] | None) -> bool:
    """True when two adjacent runs are the two halves of one split."""
    if not isinstance(first, R) or not isinstance(second, R):
        return False
    if not split or split.get(id(second)) != id(first):
        return False
    if first.r_pr != second.r_pr:
        return False
    for name in RUN_ATTRIBUTES:
        if getattr(first, name, None) != getattr(second, name, None):
            return False
    return bool(getattr(first, "content", None)) and bool(getattr(second, "content", None))


def _join_text(run: R, spaces: set[int] | None = None) -> None:
    """Fold a run's adjacent ``w:t`` items into one, as Word writes them.

    ``xml:space`` is **recomputed** for the joined value, but only when the
    split is what put it there (`spaces`, the package's ``_split_spaces``): the
    halves of ``<w:t>document.</w:t>`` need it and the whole does not, and a
    stray ``xml:space="preserve"`` would be a difference where there is none ---
    while an attribute the **document** wrote, redundant or not, is left exactly
    as it was (CR-003 section 16.12).
    """
    from docx4j_py.model.content.text_model import W_T

    items = run.content
    index = 0
    while index < len(items) - 1:
        first, second = items[index], items[index + 1]
        if element_name(first) == W_T and element_name(second) == W_T:
            value = (first.value or "") + (second.value or "")
            ours = spaces is not None and id(first) in spaces
            first.value = value
            if value != value.strip() or "  " in value:
                first.space = "preserve"
            elif ours:
                first.space = None
            del items[index + 1]
            continue
        index += 1


def para_r_pr_of(paragraph: P) -> ParaRPr:
    """The paragraph mark's run properties (``w:pPr/w:rPr``), created when absent."""
    p_pr = paragraph.p_pr
    if p_pr is None:
        p_pr = el.pPr()
        paragraph.p_pr = p_pr
        p_pr.parent = paragraph
    r_pr = p_pr.r_pr
    if r_pr is None:
        r_pr = ParaRPr()
        p_pr.r_pr = r_pr
        r_pr.parent = p_pr
    return r_pr


def row_pr_of(row: Tr) -> TrPr:
    """The row's properties (``w:trPr``), created when absent."""
    tr_pr = row.tr_pr
    if tr_pr is None:
        tr_pr = TrPr()
        row.tr_pr = tr_pr
        tr_pr.parent = row
    return tr_pr


def prune_paragraph_properties(paragraph: P) -> None:
    """Drop ``w:pPr/w:rPr`` and ``w:pPr`` once nothing is left in them.

    So that accepting or rejecting a paragraph-mark revision leaves no husk of
    an element Word would have to ignore.
    """
    p_pr = paragraph.p_pr
    if p_pr is None:
        return
    if p_pr.r_pr is not None and _is_empty(p_pr.r_pr):
        p_pr.r_pr = None
    if _is_empty(p_pr):
        paragraph.p_pr = None


def _is_empty(value: Any) -> bool:
    """True when every field of a model object is None or an empty list."""
    for field in dataclasses.fields(value):
        found = getattr(value, field.name, None)
        if found is None:
            continue
        if isinstance(found, list) and not found:
            continue
        return False
    return True


def mark_deleted(paragraph: P) -> bool:
    """True when a paragraph's mark carries ``w:pPr/w:rPr/w:del``."""
    p_pr = paragraph.p_pr
    r_pr = p_pr.r_pr if p_pr is not None else None
    return r_pr is not None and r_pr.del_value is not None


def mark_inserted(paragraph: P) -> bool:
    """True when a paragraph's mark carries ``w:pPr/w:rPr/w:ins``."""
    p_pr = paragraph.p_pr
    r_pr = p_pr.r_pr if p_pr is not None else None
    return r_pr is not None and r_pr.ins is not None


# ---------------------------------------------------------------------------
# w:t <-> w:delText
# ---------------------------------------------------------------------------


def to_deleted_text(run: R) -> None:
    """Turn a run's ``w:t`` into ``w:delText`` and ``w:instrText`` into ``w:delInstrText``."""
    _rename(run, _DELETED_NAMES)


def to_restored_text(run: R) -> None:
    """The inverse, for ``reject()`` of a deletion."""
    _rename(run, _RESTORED_NAMES)


def _rename(run: R, table: dict[str, str]) -> None:
    content = getattr(run, "content", None)
    if not content:
        return
    for index, item in enumerate(list(content)):
        local = (element_name(item) or "").rpartition("}")[2]
        to = table.get(local)
        if to is None:
            continue
        made = getattr(el, to)(getattr(item, "value", "") or "")
        space = getattr(item, "space", None)
        if space is not None:
            made.space = space
        content[index] = made
        made.parent = run


# ---------------------------------------------------------------------------
# restoring what a revision recorded
# ---------------------------------------------------------------------------


def copy_r_pr(r_pr: Any) -> Any:
    """A copy of a run's direct formatting for a new run, without its ``w:rPrChange``."""
    if r_pr is None:
        return None
    copy = deep_copy(r_pr)
    copy.r_pr_change = None
    return copy


def restore_r_pr(holder: Any, original: Any) -> None:
    """Replace a run's direct formatting with `original`, keeping nothing of what is there."""
    if original is None or _is_empty(original):
        holder.r_pr = None
        return
    holder.r_pr = original
    link_parents(original)
    original.parent = holder


def restore_p_pr(paragraph: P, original: Any) -> None:
    """Replace a paragraph's properties with a recorded ``w:pPrChange/w:pPr``.

    The mark's own run properties and the section break stay: they are not part
    of what a ``w:pPrChange`` records (its child is a ``CT_PPrBase``).
    """
    p_pr = paragraph.p_pr
    if p_pr is None:
        return
    keep_r_pr, keep_sect_pr = p_pr.r_pr, p_pr.sect_pr
    fresh = el.pPr()
    if original is not None:
        for field in dataclasses.fields(original):
            value = getattr(original, field.name, None)
            if value is None or (isinstance(value, list) and not value):
                continue
            setattr(fresh, field.name, deep_copy(value))
    fresh.r_pr = keep_r_pr
    fresh.sect_pr = keep_sect_pr
    paragraph.p_pr = fresh
    link_parents(fresh)
    fresh.parent = paragraph
    if keep_r_pr is not None:
        keep_r_pr.parent = fresh
    if keep_sect_pr is not None:
        keep_sect_pr.parent = fresh
    prune_paragraph_properties(paragraph)


# ---------------------------------------------------------------------------
# the mode, over w:trackRevisions in the settings part
# ---------------------------------------------------------------------------


def mode_of(package: Any) -> str:
    """``pkg.change_tracking_mode``: ``"Off"`` or ``"TrackAll"``.

    **A read does not unmarshal the settings part.** ``w:trackRevisions`` is
    read with lxml from the bytes the part would be saved as, exactly as
    :func:`~docx4j_py.model.content.describe.describe`'s ``tracking_on`` reads
    it, so a document whose mode is only *read* keeps ``/word/settings.xml``
    byte for byte. The answer is cached on the package (CR-003 section 16).
    """
    cached = getattr(package, "_change_tracking_mode", None)
    if cached is not None:
        return str(cached)
    mode = "TrackAll" if _track_revisions_in(package) else "Off"
    try:
        package._change_tracking_mode = mode
    except AttributeError:  # pragma: no cover - a package-like object with no slot
        pass
    return mode


def _track_revisions_in(package: Any) -> bool:
    """``w:trackRevisions`` as the file has it, through the tree or through lxml."""
    part = getattr(package, "document_settings_part", None)
    if part is None:
        return False
    if getattr(part, "is_unmarshalled", False):
        value = getattr(part.contents, "track_revisions", None)
        return value is not None and getattr(value, "val", True) is not False
    from docx4j_py.model.content.describe import _root

    root = _root(part)
    if root is None:
        return False
    node = root.find(W_TRACK_REVISIONS)
    if node is None:
        return False
    return node.get(_w("val")) not in ("0", "false")


def set_mode(package: Any, value: str | None) -> list[str]:
    """Write ``w:trackRevisions``; returns the part names touched.

    ``"TrackMineOnly"`` is stored as ``"TrackAll"`` (CR-003 section 4).
    Writing **does** unmarshal the settings part, which is what changing it
    means, and creates one --- with its relationship and its content type ---
    when the document has none.
    """
    mode = "Off" if value is None else str(value)
    if mode not in CHANGE_TRACKING_MODES:
        raise ContentError(
            f"change_tracking_mode takes Off, TrackAll or TrackMineOnly, not {mode!r}",
            code="tracking.mode_invalid",
            hint="pkg.change_tracking_mode = 'TrackAll'",
        )
    touched: list[str] = []
    part = settings_part_for_write(package)
    if part is None:
        if mode == "Off":
            _remember(package, mode)
            return touched
        part = create_settings_part(package)
    settings = part.contents
    if mode == "Off":
        settings.track_revisions = None
    elif settings.track_revisions is None:
        settings.track_revisions = el.trackRevisions()
        settings.track_revisions.parent = settings
    else:
        settings.track_revisions.val = None
    touched.append(str(part.part_name))
    _remember(package, mode)
    return touched


def settings_part_for_write(package: Any) -> Any:
    """The settings part a write goes to: a trial's own copy, or the real part.

    Shared with :mod:`~docx4j_py.model.content.compatibility`, which writes into
    the same part.

    A :class:`~docx4j_py.model.content.trial.TrialPackage` answers with the copy
    it is about to edit, so that a dry run of "turn tracking on" leaves the real
    document's ``/word/settings.xml`` alone (CR-003 section 16).
    """
    writable = getattr(package, "writable_settings_part", None)
    if writable is not None:
        return writable()
    return getattr(package, "document_settings_part", None)


def _remember(package: Any, mode: str) -> None:
    """Cache the mode and throw away the tracker the old one made."""
    try:
        package._change_tracking_mode = mode
        package._change_tracker = None
    except AttributeError:  # pragma: no cover - a package-like object with no slot
        pass


def create_settings_part(package: Any) -> Any:
    """``/word/settings.xml``, created as ``create_package`` creates it.

    Shared with :mod:`~docx4j_py.model.content.compatibility`.
    """
    from docx4j_py.openpackaging.parts.wml import DocumentSettingsPart
    from docx4j_py.wml import Settings

    main = getattr(package, "main_document_part", None)
    if main is None:
        raise ContentError(
            "this package has no main document part, so it has no settings part",
            code="tracking.no_part",
            hint="load a .docx, or call create_package()",
        )
    part = DocumentSettingsPart()
    part.set_contents(Settings())
    main.add_target_part(part)
    return part
