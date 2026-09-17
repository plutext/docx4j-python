"""``TrackedChange``: Office JS ``Word.TrackedChange`` over Word's revision markup.

CR-003 section 3.8, Phase F. The reading half; :mod:`.tracking` is the writing
half. A view like :class:`~docx4j_py.model.content.paragraph.Paragraph` and
:class:`~docx4j_py.model.content.range.Range`: it holds the markup it is over,
nothing is cached, and two views of the same markup compare equal.

::

    for change in pkg.body.get_tracked_changes():
        print(change.type, change.author, change.text)
    pkg.body.accept_all()

What a change may be over (ECMA-376 17.13.5): a ``w:ins``, ``w:del``,
``w:moveFrom`` or ``w:moveTo`` around runs; a paragraph mark's
``w:pPr/w:rPr/w:ins`` or ``w:pPr/w:rPr/w:del``; a run's ``w:rPrChange`` or a
paragraph's ``w:pPrChange``; a row's ``w:trPr/w:ins`` or ``w:trPr/w:del``.

**Accepting** follows docx4j's
``org.docx4j.convert.out.common.preprocess.AcceptTrackedChanges``: a ``w:ins``
is unwrapped, a ``w:del`` removed, a deleted paragraph mark joins its paragraph
with the next --- the joined paragraph keeps the first one's content and
``w14:paraId`` and takes the second one's properties --- and a deleted row is
removed. Unlike that conversion preprocessor, which leaves formatting revisions
alone because the current properties are what the document shows, accepting here
also **drops** ``w:rPrChange`` and ``w:pPrChange``: this document is saved
again (CR-003 section 4). **Rejecting** is the mirror.
"""

from __future__ import annotations

import dataclasses
import datetime
from typing import TYPE_CHECKING, Any

from docx4j_py.child import link_parents
from docx4j_py.model.content.enums import TrackedChangeTypeValue
from docx4j_py.model.content.errors import TrackedChangeError
from docx4j_py.model.content.text_model import W_DEL_TEXT, block_children_of, item_text
from docx4j_py.model.content.tracking import (
    REVISION_NAMES,
    date_of,
    prune_paragraph_properties,
    restore_p_pr,
    restore_r_pr,
    to_restored_text,
)
from docx4j_py.traversal import element_name, run_items_of
from docx4j_py.wml import P, R, Tbl

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "TrackedChange",
    "TrackedChangeTarget",
    "changes_in_row",
    "folded_in_row",
    "join_with_next",
    "tracked_changes_of_body",
    "tracked_changes_of_paragraph",
    "tracked_changes_of_row",
]

#: How much of the text a ``repr`` shows.
_PREVIEW = 30


@dataclasses.dataclass(frozen=True, slots=True)
class TrackedChangeTarget:
    """Which piece of revision markup a :class:`TrackedChange` is over.

    An extension: Office JS has no such member, and docx4j-core-ts carries the
    same thing as a discriminated union.
    """

    kind: str
    """One of ``"run"``, ``"mark"``, ``"run_properties"``,
    ``"paragraph_properties"``, ``"row"``."""
    value: Any
    """What carries ``w:id``, ``w:author`` and ``w:date``: a ``CT_TrackChange``,
    a ``w:rPrChange`` or a ``w:pPrChange``."""
    element: Any = None
    """The element the change is on: the ``w:ins``, the ``w:tr``, the ``w:r``,
    the ``w:p``."""
    owner: list | None = None
    """The live list holding :attr:`element`, for the kinds that are removed."""
    mark: str | None = None
    """``"ins"`` or ``"del"`` for a mark or a row; the revision kind for a run."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the kind and, where there is one, the mark."""
        out: dict[str, Any] = {"kind": self.kind}
        if self.mark is not None:
            out["mark"] = self.mark
        return out


class TrackedChange:
    """A subset of Office JS ``Word.TrackedChange`` over one piece of revision markup."""

    __slots__ = ("body", "paragraph", "target")

    def __init__(
        self, target: TrackedChangeTarget, paragraph: Paragraph | None, body: Body | None = None
    ) -> None:
        """Build the view over the markup, the paragraph it is in and its body."""
        #: What the change is over (an extension).
        self.target = target
        #: The paragraph the change is in; None for a row revision.
        self.paragraph = paragraph
        #: The body it was collected from, for :attr:`address`; may be None.
        self.body = body if body is not None else getattr(paragraph, "parent_body", None)

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same markup are equal (CR-003 section 3.1)."""
        return isinstance(other, TrackedChange) and other.target.value is self.target.value

    def __hash__(self) -> int:
        """Hashes by the markup's identity."""
        return hash(id(self.target.value))

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
        """Office JS ``Word.TrackedChangeType``: ``"Added"``, ``"Deleted"`` or ``"Formatted"``.

        ``"None"`` is in Office JS's list and is never produced here: every
        piece of markup a change is over is one of the three.
        """
        target = self.target
        if target.kind == "run":
            return "Added" if target.mark in ("ins", "moveTo") else "Deleted"
        if target.kind in ("mark", "row"):
            return "Added" if target.mark == "ins" else "Deleted"
        return "Formatted"

    @property
    def author(self) -> str:
        """``w:author``: who made the change."""
        return str(getattr(self.target.value, "author", "") or "")

    @property
    def date(self) -> datetime.datetime | None:
        """``w:date``, when the markup carries one; None when it does not."""
        return date_of(getattr(self.target.value, "date", None))

    @property
    def text(self) -> str:
        """The text the change covers; ``""`` for a paragraph mark.

        A deletion's text is the **deleted** text (its ``w:delText``), which is
        what the original view shows and what ``to_markdown(view="markup")``
        renders between ``{--`` and ``--}``.
        """
        target = self.target
        if target.kind == "run":
            return _revision_text(target.element)
        if target.kind == "run_properties":
            return _revision_text_of_runs([target.element])
        if target.kind == "paragraph_properties":
            return self.paragraph.text if self.paragraph is not None else ""
        if target.kind == "row":
            return _row_text(target.element)
        return ""

    def get_range(self) -> Range | None:
        """The range the change covers, or None when it has no paragraph.

        A deletion is a **collapsed** range where it sits in the accepted text,
        because its text is not in that view at all; a mark or a ``w:pPrChange``
        is the whole paragraph; a row revision has no range.
        """
        from docx4j_py.model.content.range import Range

        paragraph = self.paragraph
        if paragraph is None:
            return None
        target = self.target
        if target.kind == "run":
            start, end = _span_of(paragraph, target.element)
            return Range(paragraph, start, end)
        if target.kind == "run_properties":
            covered = [s for s in paragraph.segments() if s.run is target.element]
            if not covered:
                return paragraph.get_range("Start")
            return Range(paragraph, covered[0].start, covered[-1].end)
        return paragraph.get_range("Whole")

    def accept(self) -> None:
        """Keep the change: a ``w:ins`` unwrapped, a ``w:del`` removed, a mark joined.

        Raises:
            TrackedChangeError: the markup is no longer in the tree.
        """
        target = self.target
        if target.kind == "run":
            if target.mark in ("ins", "moveTo"):
                _unwrap(target)
            else:
                _remove(target, "accept")
            return
        if target.kind == "mark":
            if target.mark == "ins":
                _drop_mark(self.paragraph, "ins")
            else:
                join_with_next(self._require_paragraph())
            return
        if target.kind == "run_properties":
            r_pr = getattr(target.element, "r_pr", None)
            if r_pr is not None:
                r_pr.r_pr_change = None
            return
        if target.kind == "paragraph_properties":
            paragraph = self._require_paragraph().element
            if paragraph.p_pr is not None:
                paragraph.p_pr.p_pr_change = None
            prune_paragraph_properties(paragraph)
            return
        if target.mark == "ins":
            # the cell-level insertions this row is made of are folded into it
            # (CR-003 section 16.10), so accepting the row accepts them
            for inner in reversed(folded_in_row(self)):
                inner.accept()
            tr_pr = getattr(target.element, "tr_pr", None)
            if tr_pr is not None:
                tr_pr.ins = None
        else:
            _remove_row(target, "accept")  # the content goes with the row

    def reject(self) -> None:
        """Put back what was there: a ``w:ins`` removed, a ``w:del`` restored.

        Raises:
            TrackedChangeError: the markup is no longer in the tree.
        """
        target = self.target
        if target.kind == "run":
            if target.mark in ("ins", "moveTo"):
                _remove(target, "reject")
            else:
                _restore_deleted(target)
            return
        if target.kind == "mark":
            if target.mark == "del":
                _drop_mark(self.paragraph, "del")
            else:
                paragraph = self._require_paragraph()
                join_with_next(
                    paragraph,
                    fallback_to_previous=True,
                    keep_properties=True,
                )
                # the break this mark recorded is gone, so the mark goes with
                # it: keeping the surviving paragraph's own properties (16.10)
                # means nothing else takes it away (CR-003 section 16.11)
                _drop_mark(paragraph, "ins")
            return
        if target.kind == "run_properties":
            from docx4j_py.wml import RPr, rpr_from_elements

            restore_r_pr(target.element, rpr_from_elements(target.value.r_pr, cls=RPr))
            return
        if target.kind == "paragraph_properties":
            restore_p_pr(self._require_paragraph().element, target.value.p_pr)
            return
        if target.mark == "del":
            for inner in reversed(folded_in_row(self)):
                inner.reject()
            tr_pr = getattr(target.element, "tr_pr", None)
            if tr_pr is not None:
                tr_pr.del_value = None
        else:
            _remove_row(target, "reject")  # the content goes with the row

    # -- the extensions ----------------------------------------------------

    @property
    def id(self) -> int | None:
        """The annotation id (``w:id``). Extension: Office JS has no id."""
        value = getattr(self.target.value, "id", None)
        return value if isinstance(value, int) else None

    @property
    def element(self) -> Any:
        """The markup this change is over. Extension."""
        return self.target.element if self.target.element is not None else self.target.value

    @property
    def kind(self) -> str:
        """``"run"``, ``"mark"``, ``"run_properties"``, ``"paragraph_properties"``, ``"row"``."""
        return self.target.kind

    @property
    def address(self) -> str:
        """The address of the paragraph (or the row) the change is in. Extension."""
        if self.paragraph is not None:
            return self.paragraph.address
        if self.body is not None and self.target.element is not None:
            from docx4j_py.model.content.addresses import ordinal_of

            return ordinal_of(self.body, self.target.element) or ""
        return ""

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about one revision."""
        date = self.date
        return {
            "type": self.type,
            "author": self.author,
            "date": date.isoformat() if date is not None else None,
            "text": self.text,
            "id": self.id,
            "kind": self.target.kind,
            "address": self.address,
        }

    # -- helpers -----------------------------------------------------------

    def _require_paragraph(self) -> Paragraph:
        if self.paragraph is None:
            raise TrackedChangeError(
                "this tracked change has no paragraph",
                code="tracking.no_paragraph",
                hint="collect it through body.get_tracked_changes(), which carries one",
            )
        return self.paragraph


# ---------------------------------------------------------------------------
# what accepting and rejecting do
# ---------------------------------------------------------------------------


def _index_of(items: list, element: Any) -> int:
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def _remove(target: TrackedChangeTarget, verb: str) -> None:
    """Take the element out of the list holding it."""
    owner = target.owner
    index = _index_of(owner, target.element) if owner is not None else -1
    if index < 0:
        raise TrackedChangeError(
            f"this tracked change is no longer in the document, so it cannot be {verb}ed",
            code="tracking.gone",
            hint="call get_tracked_changes() again after accepting or rejecting others",
        )
    del owner[index]


def _remove_row(target: TrackedChangeTarget, verb: str) -> None:
    """Take a row out, and the table with it when that was its last row.

    A departure from docx4j's ``AcceptTrackedChanges``, which is a conversion
    preprocessor and leaves the husk: a ``w:tbl`` with no ``w:tr`` in it is not
    valid WordprocessingML, and this document is saved again (CR-003 section
    16).
    """
    row = target.element
    table = getattr(row, "parent", None)
    _remove(target, verb)
    while table is not None and isinstance(table, Tbl):
        from docx4j_py.model.content.text_model import rows_of

        if rows_of(table):
            return
        holder = getattr(table, "parent", None)
        items = block_children_of(holder) if holder is not None else None
        index = _index_of(items, table) if items is not None else -1
        if index < 0:
            return
        del items[index]
        table = None


def _unwrap(target: TrackedChangeTarget) -> None:
    """A ``w:ins`` or ``w:moveTo`` accepted: its runs take its place."""
    owner = target.owner
    index = _index_of(owner, target.element) if owner is not None else -1
    if index < 0:
        raise TrackedChangeError(
            "this tracked change is no longer in the document, so it cannot be accepted",
            code="tracking.gone",
            hint="call get_tracked_changes() again after accepting or rejecting others",
        )
    items = list(run_items_of(target.element) or ())
    parent = getattr(target.element, "parent", None)
    owner[index : index + 1] = items
    for item in items:
        item.parent = parent


def _restore_deleted(target: TrackedChangeTarget) -> None:
    """A ``w:del`` or ``w:moveFrom`` rejected: its runs come back as ``w:t``."""
    for item in run_items_of(target.element) or ():
        if isinstance(item, R):
            to_restored_text(item)
    _unwrap(target)


def _drop_mark(paragraph: Paragraph | None, which: str) -> None:
    """Take ``w:ins`` or ``w:del`` off a paragraph mark and tidy up after it."""
    if paragraph is None:
        return
    p_pr = paragraph.element.p_pr
    r_pr = p_pr.r_pr if p_pr is not None else None
    if r_pr is not None:
        if which == "ins":
            r_pr.ins = None
        else:
            r_pr.del_value = None
    prune_paragraph_properties(paragraph.element)


def join_with_next(
    paragraph: Paragraph,
    *,
    fallback_to_previous: bool = False,
    keep_properties: bool = False,
) -> None:
    """Join a paragraph with the next one, as docx4j's ``AcceptTrackedChanges`` does.

    The paragraph takes the next one's content **and its properties**, so the
    mark that survives is the next one's and the ``w14:paraId`` that survives is
    this one's. Called when a deleted mark is accepted and when an inserted one
    is rejected.

    `keep_properties` is the **reject** direction's difference: the paragraph
    that survives is one the document already had --- the break after it is what
    was inserted --- so it keeps its own ``w:pPr`` rather than taking the new
    paragraph's, which would strip the formatting of a paragraph nobody edited
    (CR-003 section 16.10). Its caller then drops the ``w:ins`` from that
    ``w:pPr`` itself, since replacing the properties wholesale is what used to
    carry the mark away (section 16.11).

    `fallback_to_previous` is for the second case at the end of a container,
    where a paragraph this package inserted carries its own mark (CR-003 section
    4) and there is nothing after it to join with: the paragraph then goes, its
    content joining the one before. With neither neighbour the mark is dropped.
    """
    container = paragraph.container
    index = _index_of(container, paragraph.element)
    following = container[index + 1] if 0 <= index < len(container) - 1 else None
    element = paragraph.element
    if not isinstance(following, P):
        previous = container[index - 1] if fallback_to_previous and index > 0 else None
        if isinstance(previous, P):
            for item in list(element.content):
                previous.content.append(item)
                item.parent = previous
            link_parents(previous)
            del container[index]
            return
        _drop_mark(paragraph, "ins")
        _drop_mark(paragraph, "del")
        return
    for item in list(following.content):
        element.content.append(item)
        item.parent = element
    if not keep_properties:
        element.p_pr = following.p_pr
        if element.p_pr is not None:
            element.p_pr.parent = element
    link_parents(element)
    element.parent = getattr(following, "parent", None) or element.parent
    del container[index + 1]


# ---------------------------------------------------------------------------
# the text of a revision, and where it sits
# ---------------------------------------------------------------------------


def _revision_text(element: Any) -> str:
    """The text a run-level revision covers, its ``w:delText`` counted."""
    return _revision_text_of_runs(run_items_of(element) or [])


def _revision_text_of_runs(items: list) -> str:
    out: list[str] = []

    def visit(entries: list) -> None:
        for entry in entries:
            if isinstance(entry, R):
                for child in getattr(entry, "content", None) or ():
                    qname = element_name(child)
                    if qname == W_DEL_TEXT:
                        value = getattr(child, "value", None)
                        out.append(value if isinstance(value, str) else "")
                        continue
                    text = item_text(qname, child)
                    if text is not None:
                        out.append(text)
                continue
            nested = run_items_of(entry)
            if nested is not None:
                visit(nested)

    visit(items)
    return "".join(out)


def _span_of(paragraph: Paragraph, element: Any) -> tuple[int, int]:
    """A run-level revision's span in the paragraph's accepted text."""
    from docx4j_py.model.content.tracking import revision_of

    inside = [
        segment
        for segment in paragraph.segments()
        if (found := revision_of(segment.run)) is not None and found.element is element
    ]
    if inside:
        return inside[0].start, inside[-1].end
    at = _accepted_offset_of(paragraph.element, element)
    return at, at


def _accepted_offset_of(paragraph: P, target: Any) -> int:
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


# ---------------------------------------------------------------------------
# collecting
# ---------------------------------------------------------------------------


def tracked_changes_of_paragraph(paragraph: Paragraph) -> list[TrackedChange]:
    """The tracked changes of one paragraph, in document order."""
    out: list[TrackedChange] = []
    element = paragraph.element
    p_pr = element.p_pr
    if p_pr is not None and p_pr.p_pr_change is not None:
        out.append(
            TrackedChange(
                TrackedChangeTarget("paragraph_properties", p_pr.p_pr_change, element),
                paragraph,
            )
        )
    r_pr = p_pr.r_pr if p_pr is not None else None
    if r_pr is not None and r_pr.ins is not None:
        out.append(
            TrackedChange(TrackedChangeTarget("mark", r_pr.ins, element, mark="ins"), paragraph)
        )
    if r_pr is not None and r_pr.del_value is not None:
        out.append(
            TrackedChange(
                TrackedChangeTarget("mark", r_pr.del_value, element, mark="del"), paragraph
            )
        )
    _collect_run_level(paragraph, element.content, out)
    return out


def _collect_run_level(paragraph: Paragraph, items: list, out: list[TrackedChange]) -> None:
    for item in items:
        name = element_name(item)
        if name in REVISION_NAMES:
            kind = (name or "").rpartition("}")[2]
            out.append(
                TrackedChange(
                    TrackedChangeTarget("run", item, item, owner=items, mark=kind), paragraph
                )
            )
            nested = run_items_of(item)
            if nested is not None:
                _collect_run_level(paragraph, nested, out)
            continue
        if isinstance(item, R):
            r_pr = getattr(item, "r_pr", None)
            if r_pr is not None and r_pr.r_pr_change is not None:
                out.append(
                    TrackedChange(
                        TrackedChangeTarget("run_properties", r_pr.r_pr_change, item), paragraph
                    )
                )
            continue
        nested = run_items_of(item)
        if nested is not None:
            _collect_run_level(paragraph, nested, out)


def _row_text(row: Any) -> str:
    """A row's text, a cell paragraph per line, its ``w:delText`` counted."""
    from docx4j_py.model.content.text_model import cells_of

    lines: list[str] = []

    def visit(blocks: list) -> None:
        for block in blocks:
            if isinstance(block, P):
                lines.append(_revision_text_of_runs(block.content))
                continue
            children = block_children_of(block)
            if children is not None:
                visit(children)

    for cell, _owner in cells_of(row):
        children = block_children_of(cell)
        if children is not None:
            visit(children)
    return "\n".join(lines)


def changes_in_row(row: Any, body: Body | None = None) -> list[TrackedChange]:
    """Every tracked change in a row's cells, the row's own revisions apart."""
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.text_model import cells_of, rows_of

    out: list[TrackedChange] = []

    def visit(blocks: list) -> None:
        for block in blocks:
            if isinstance(block, P):
                out.extend(tracked_changes_of_paragraph(Paragraph(block, blocks, body)))
                continue
            if isinstance(block, Tbl):
                for nested, owner in rows_of(block):
                    out.extend(tracked_changes_of_row(nested, owner, body))
                    out.extend(changes_in_row(nested, body))
                continue
            children = block_children_of(block)
            if children is not None:
                visit(children)

    for cell, _owner in cells_of(row):
        children = block_children_of(cell)
        if children is not None:
            visit(children)
    return out


def folded_in_row(change: TrackedChange) -> list[TrackedChange]:
    """The cell-level changes a row revision is **made of**, in document order.

    CR-003 section 16.10: a row insertion or deletion is **one**
    :class:`TrackedChange`, as Word's Reviewing pane shows one, and the cell
    revisions of the same author and the same direction are folded into it ---
    they are that same edit. Anything else inside the row --- another author's
    change, a formatting revision --- is listed separately and is not touched by
    accepting or rejecting the row.
    """
    if change.target.kind != "row":
        return []
    wanted = "Added" if change.target.mark == "ins" else "Deleted"
    author = change.author
    return [
        inner
        for inner in changes_in_row(change.target.element, change.body)
        if inner.type == wanted and inner.author == author
    ]


def tracked_changes_of_row(row: Any, owner: list, body: Body | None = None) -> list[TrackedChange]:
    """The ``w:trPr/w:ins`` and ``w:trPr/w:del`` revisions of a table row."""
    tr_pr = getattr(row, "tr_pr", None)
    out: list[TrackedChange] = []
    if tr_pr is None:
        return out
    if tr_pr.ins is not None:
        out.append(
            TrackedChange(
                TrackedChangeTarget("row", tr_pr.ins, row, owner=owner, mark="ins"), None, body
            )
        )
    if tr_pr.del_value is not None:
        out.append(
            TrackedChange(
                TrackedChangeTarget("row", tr_pr.del_value, row, owner=owner, mark="del"),
                None,
                body,
            )
        )
    return out


def tracked_changes_of_body(body: Body) -> list[TrackedChange]:
    """Every tracked change in a body, in document order.

    Tables are descended into: a row's own revisions first, then the changes in
    its cells, which is the order accepting in reverse needs.
    """
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.text_model import cells_of, rows_of

    out: list[TrackedChange] = []

    def visit(items: list) -> None:
        # ``items`` is the **live** list, and the views built here hold it: a
        # paragraph join removes the next paragraph from the list it is in
        # (CR-003 section 4), so a copy would leave it behind.
        for item in list(items):
            if isinstance(item, P):
                out.extend(tracked_changes_of_paragraph(Paragraph(item, items, body)))
                continue
            if isinstance(item, Tbl):
                for row, owner in rows_of(item):
                    revisions = tracked_changes_of_row(row, owner, body)
                    out.extend(revisions)
                    if revisions:
                        # the row's own revision folds the cell-level ones it is
                        # made of; everything else in the row is still listed
                        # a set of views: two of the same markup are equal and
                        # hash alike, so this is an identity test on the markup
                        folded = {c for revision in revisions for c in folded_in_row(revision)}
                        out.extend(
                            inner for inner in changes_in_row(row, body) if inner not in folded
                        )
                        continue
                    for cell, _cell_owner in cells_of(row):
                        children = block_children_of(cell)
                        if children is not None:
                            visit(children)
                continue
            children = block_children_of(item)
            if children is not None:
                visit(children)

    visit(body.content)
    return out
