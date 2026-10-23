"""``Range.hyperlink``: Office JS's read and write of a link, over ``w:hyperlink``.

CR-003 section 21, the port of docx4j-core-ts CR-002 sections 20 and 25
(``Range.mts``, its ``d64ab67`` and the Word-measured corrections after it). The
member is Office JS's, a string ``address#location``:

* the **address** becomes an external relationship of type
  ``.../relationships/hyperlink`` of **the range's own part** (a header's
  paragraph relates from the header part) and the ``r:id`` of the
  ``w:hyperlink``;
* the **location** becomes its ``w:anchor``, so ``"#chapter1"`` is a link
  within the document, with no relationship, and
  ``"https://example.com/doc#s2"`` both.

Reading gives the first link the range touches, ``""`` where there is none.
Writing removes the links the range touches and wraps its runs in one new
``w:hyperlink``, split at the span's boundaries as :attr:`Range.font` splits
them; writing ``""`` (or None) only removes. What Word does, as docx4j-core-ts
measured it through Office JS and this follows (its ``test/README.md`` checks
15, 16 and 19):

* each run a link wraps takes the ``Hyperlink`` **character style**, in place of
  any character style it had, and loses its direct ``w:color``; the rest of its
  direct formatting stays. A link removed takes the style off its runs again,
  and an ``w:rPr`` left empty goes. The style's definition is added to
  ``styles.xml`` when the document lacks it (docx4j's ``KnownStyles.xml``);
* a range that crosses a link's **edge** takes the whole of that link into the
  new one; a range **inside** a link replaces it with a link over the range
  alone; removing from part of a link removes all of it;
* an **empty** range is refused, as Office JS refuses it;
* a new link's adjacent text-only runs with identical properties become one run,
  as Word writes the link;
* ``w:history`` is written, as Word and docx4j write it on every link they make.

A range that would put the new link across a **run holder** --- a content
control, a field, a tracked insertion --- or around a **tracked deletion** is
refused with :class:`~docx4j_py.model.content.errors.SpanError` before anything
changes. Under change tracking the restyle is a formatting revision
(``w:rPrChange``), as a :class:`~docx4j_py.model.content.font.Font` write is; the
wrap itself is not a revision, since Word has no markup for one.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.child import link_parents
from docx4j_py.model.content.errors import ContentError, SpanError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.text_model import W_R, W_T, runs_of
from docx4j_py.model.content.tracking import REVISION_NAMES, _is_empty, _joined
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = ["HYPERLINK_STYLE", "hyperlink_of", "links_of", "set_hyperlink"]

#: The character style Word gives the runs of a link it makes.
HYPERLINK_STYLE = "Hyperlink"

W_HYPERLINK = f"{{{WML_NS}}}hyperlink"


# ---------------------------------------------------------------------------
# reading
# ---------------------------------------------------------------------------


def links_of(paragraph: Paragraph) -> list[tuple[Any, int, int]]:
    """Each ``w:hyperlink`` holding text, with its ``[start, end)``, in document order.

    Read off the paragraph's segments, so a link is where its accepted-view text
    is; a link holding no text (or only deleted text) is in no range and is not
    listed, as in docx4j-core-ts's ``spanOf``.
    """
    out: list[list[Any]] = []
    found: dict[int, list[Any]] = {}
    stop = paragraph.element
    for segment in paragraph.segments():
        node = getattr(segment.run, "parent", None)
        while node is not None and node is not stop:
            if element_name(node) == W_HYPERLINK:
                entry = found.get(id(node))
                if entry is None:
                    entry = [node, segment.start, segment.end]
                    found[id(node)] = entry
                    out.append(entry)
                else:
                    entry[2] = segment.end
            node = getattr(node, "parent", None)
    return [(entry[0], entry[1], entry[2]) for entry in out]


def _touches(span_start: int, span_end: int, start: int, end: int) -> bool:
    """Whether a range ``[start, end)`` touches a link's text ``[span_start, span_end)``.

    Overlapping it; or, for an empty range, lying in it **or at either end of
    it** --- Office JS reads a caret just after a link as in that link
    (docx4j-core-ts check 19, case K; the start edge was not measured).
    """
    if start == end:
        return span_start <= start <= span_end
    return span_start < end and span_end > start


def touched_links(span: Range) -> list[tuple[Any, int, int]]:
    """The links :func:`links_of` lists that `span` touches."""
    return [
        link
        for link in links_of(span.paragraph)
        if _touches(link[1], link[2], span.start, span.end)
    ]


def hyperlink_of(span: Range) -> str:
    """``Range.hyperlink``'s read: the first link the span touches, as ``address#location``.

    A range across paragraphs reads the first link among its paragraphs'
    pieces (CR-013 Phase B), through :attr:`Range.hyperlink`.
    """
    if span.spans_paragraphs:
        return span.hyperlink
    found = touched_links(span)
    if not found:
        return ""
    link = found[0][0]
    anchor = str(getattr(link, "anchor", None) or "")
    rel_id = getattr(link, "id", None)
    address = _relationship_target(span, str(rel_id)) if rel_id else ""
    return f"{address}#{anchor}" if anchor else address


def _relationship_target(span: Range, rel_id: str) -> str:
    """The target of a relationship of the span's part, or ``""``."""
    part = span.paragraph.parent_body.part
    rels = getattr(part, "relationships_part", None) if part is not None else None
    rel = rels.get_relationship_by_id(rel_id) if rels is not None else None
    return str(getattr(rel, "target", None) or "") if rel is not None else ""


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------


def set_hyperlink(span: Range, value: str | None) -> None:
    """``Range.hyperlink``'s write. See this module's docstring for the rules.

    Raises:
        SpanError: the span is empty (``range.empty``), covers no text
            (``range.no_text``), or would put the link across a run holder or
            around a tracked change (``range.crosses_holder``). Nothing has
            changed when it raises.
        ContentError: an address was given and the span's body has no part to
            relate it from (``hyperlink.no_part``).
        SpanError: ``range.spans_paragraphs``, a range across paragraphs
            (CR-013 Phase B).
    """
    span.require_one_paragraph("Range.hyperlink")
    value = "" if value is None else str(value)
    if span.start == span.end:
        # Office JS refuses a hyperlink set on an empty range --- inside a link,
        # in plain text, just after one, and "" alike --- and changes nothing
        # (docx4j-core-ts CR-002 section 25, check 19)
        raise SpanError(
            f"a hyperlink cannot be set on an empty range ({span.start}:{span.end}), "
            "as Office JS refuses it",
            code="range.empty",
            hint="select the text to link, or the whole of a link to change it",
        )
    paragraph = span.paragraph
    body = paragraph.parent_body
    address = value.partition("#")[0]
    if address and body.part is None:
        raise ContentError(
            "this range's body belongs to no part, so an address cannot be related to it",
            code="hyperlink.no_part",
            hint="link text in a document's body, header or note, or set a #location only",
        )
    with recording(body, "hyperlink") as change:
        change.text(before=f"hyperlink={hyperlink_of(span)}")
        touched = touched_links(span)
        start, end = span.start, span.end
        for _link, link_start, link_end in touched:
            # a link the range lies within is replaced, not taken in (check 16, G);
            # one whose edge it crosses is taken in whole (check 16, F)
            if link_start <= span.start and span.end <= link_end:
                continue
            start = min(start, link_start)
            end = max(end, link_end)
        unwrapped = [link for link, _s, _e in touched]
        # the span and every link it takes in, before anything changes (CR-011 Phase F)
        paragraph.refuse_branch_text(start, end, "Range.hyperlink")
        if value:
            _check_wrappable(paragraph, start, end, unwrapped)
        for link in unwrapped:
            _unwrap(paragraph, link)
        if value:
            _wrap(span, start, end, value, change)
        change.touched(paragraph)
        change.text(after=f"hyperlink={value}")


def _check_wrappable(paragraph: Paragraph, start: int, end: int, unwrapped: list[Any]) -> None:
    """Refuse, before anything changes, a span one ``w:hyperlink`` cannot wrap.

    The span's runs must sit in **one** run-level list once the links being
    replaced are unwrapped --- the same rule, through the same walk, as
    :meth:`Range.require_one_holder`, with the four revision holders counted
    as holders too --- and nothing between its first run and its last may be a
    tracked change (a ``w:del`` between two runs has no text in the accepted
    view, so only this finds it).
    """
    from docx4j_py.model.content.range import Range

    extent = Range(paragraph, start, end)
    extent.require_one_holder("Range.hyperlink", revisions=True, unwrapped=unwrapped)

    seen_through = {id(link) for link in unwrapped}
    tops: list[Any] = []
    holder: Any = None
    for segment in paragraph.segments():
        if segment.end <= start or segment.start >= end:
            continue
        top = segment.run
        holder = getattr(top, "parent", None)
        while holder is not None and id(holder) in seen_through:
            top = holder
            holder = getattr(holder, "parent", None)
        if not tops or tops[-1] is not top:
            tops.append(top)
    if not tops:
        raise SpanError(
            f"the range {start}:{end} covers no text of this paragraph",
            code="range.no_text",
            hint="use search() or get_range() for a span of this paragraph's text",
        )
    owner = run_items_of(holder) if holder is not None else None
    if owner is None:  # pragma: no cover - a run always sits in a run-level list
        return
    first = _index_of(owner, tops[0])
    last = _index_of(owner, tops[-1])
    for item in owner[first : last + 1]:
        inner = run_items_of(item) if id(item) in seen_through else None
        for element in inner if inner is not None else (item,):
            name = element_name(element)
            if name in REVISION_NAMES:
                raise SpanError(
                    f"the range {start}:{end} holds a tracked change "
                    f"(w:{(name or '').rpartition('}')[2]}) between its runs, "
                    "which a hyperlink would have to wrap",
                    code="range.crosses_holder",
                    hint="accept or reject the tracked change first, or link a span on one side of it",
                )


def _unwrap(paragraph: Paragraph, link: Any) -> None:
    """Take a ``w:hyperlink`` away, its runs staying where they were, unstyled."""
    parent = getattr(link, "parent", None)
    owner = run_items_of(parent) if parent is not None else None
    if owner is None:  # pragma: no cover - a link always sits in a run-level list
        return
    at = _index_of(owner, link)
    if at < 0:  # pragma: no cover - the link was found through this tree
        return
    _restyle(paragraph, link, linked=False)
    items = list(link.content or ())
    owner[at : at + 1] = items
    for item in items:
        item.parent = parent


def _wrap(span: Range, start: int, end: int, value: str, change: Any) -> None:
    """The new ``w:hyperlink`` over ``[start, end)``: relationship, split, wrap, style."""
    from docx4j_py.model.content.styles import define_built_in
    from docx4j_py.wml import el

    paragraph = span.paragraph
    body = paragraph.parent_body
    address, _hash, location = value.partition("#")
    # w:history="1" is what Word writes on every link it makes and what docx4j
    # writes when it builds one ("add the target to the viewed hyperlinks",
    # ECMA-376 17.16.22); docx4j-core-ts CR-002 section 20
    link = el.hyperlink(history=True)
    if location:
        link.anchor = location
    if address:
        link.id = _relate(body, address, change)

    used_start = paragraph.split_at(start, prefer="back")
    used_end = paragraph.split_at(end, prefer="forward")
    # the range goes on covering what it covered, snapped as Font snaps it
    if start == span.start:
        span.start = used_start
    if end == span.end:
        span.end = used_end
    segments = [s for s in paragraph.segments() if s.end > used_start and s.start < used_end]
    first, last = segments[0].run, segments[-1].run
    owner = segments[0].run_owner
    parent = getattr(first, "parent", None) or paragraph.element
    begin = _index_of(owner, first)
    finish = _index_of(owner, last)
    if element_name(parent) in REVISION_NAMES:
        _wrap_in_revision(paragraph, link, parent, begin, finish)
    else:
        items = list(owner[begin : finish + 1])
        owner[begin : finish + 1] = [link]
        link.content.extend(items)
        link_parents(link)
        link.parent = parent

    _restyle(paragraph, link, linked=True)
    _merge_text_runs(link.content, getattr(body.package, "_split_runs", None))
    define_built_in(body.package, HYPERLINK_STYLE)


def _wrap_in_revision(
    paragraph: Paragraph, link: Any, revision: Any, begin: int, finish: int
) -> None:
    """Wrap runs that sit in a ``w:ins`` (or ``w:moveTo``): the link goes outside it.

    ``CT_RunTrackChange`` holds ``EG_ContentRunContent``, which has no
    ``w:hyperlink``, so a link cannot be **in** a revision; a revision can be in
    a link (``EG_PContent``). The revision is split at the span --- the part
    before keeps the element and its ``w:id``, the span's runs go into a revision
    of the same kind and attributes with a new id **inside** the link, and the
    part after into another --- which is docx4j-core-ts's form for a run holder
    in an insertion (its CR-002 section 37 item 4) and for a content control
    over inserted text (check 30). Rejecting the insertion then leaves the link
    empty, as Word leaves that form.
    """
    outer_parent = getattr(revision, "parent", None)
    outer = run_items_of(outer_parent) if outer_parent is not None else None
    if outer is None:  # pragma: no cover - a revision always sits in a run-level list
        raise SpanError(
            "this range is in a tracked change that sits in nothing",
            code="range.crosses_holder",
            hint="link a span outside the tracked change",
        )
    at = _index_of(outer, revision)
    items = list(revision.content)
    head, middle, tail = items[:begin], items[begin : finish + 1], items[finish + 1 :]
    if not head and not tail:
        # the whole revision: it goes into the link as it is, its id kept
        outer[at : at + 1] = [link]
        link.content.append(revision)
    else:
        # the element keeps what stays in front of the link (or after it, when
        # nothing does) **first**: assigning its content orphans every run it
        # held, so the clones adopt theirs only once that is done
        ids = _id_source(paragraph)
        revision.content[:] = head or tail
        inner = _revision_like(revision, middle, ids)
        link.content.append(inner)
        if head:
            pieces = [link] if not tail else [link, _revision_like(revision, tail, ids)]
            outer[at + 1 : at + 1] = pieces
        else:
            outer[at:at] = [link]
    link_parents(link)
    for item in outer:
        if item is link or element_name(item) in REVISION_NAMES:
            item.parent = outer_parent


def _id_source(paragraph: Paragraph) -> Any:
    """What allocates the new revisions' ``w:id``: the package's tracker, or one made for this.

    With tracking off the document still has revisions, so the ids come from
    the same annotation id space, one above the highest in use; one tracker for
    the whole call, so that two new revisions do not get the same id.
    """
    from docx4j_py.model.content.tracking import ChangeTracker

    tracker = paragraph.change_tracker
    package = paragraph.parent_body.package
    if tracker is None and package is not None:
        tracker = ChangeTracker(package)
    return tracker


def _revision_like(revision: Any, items: list, ids: Any) -> Any:
    """A revision of the same kind and attributes as `revision`, a new ``w:id``, holding `items`."""
    import dataclasses

    from docx4j_py.child import ChildList

    values = {
        field.name: getattr(revision, field.name)
        for field in dataclasses.fields(revision)
        if field.name != "content"
    }
    clone = type(revision)(**values, content=ChildList())
    if ids is not None:
        clone.id = ids.next_id()
    clone.content.extend(items)
    link_parents(clone)
    return clone


def _relate(body: Any, address: str, change: Any) -> str:
    """A new external hyperlink relationship of the body's part; its id.

    In a :func:`~docx4j_py.model.content.trial.dry_run` the relationship is
    added to the real part's ``.rels`` (a trial shares the part graph) and
    taken away again when the block ends, with the relationships part itself
    when this call created it.
    """
    from docx4j_py.model.content.trial import note_added_relationship
    from docx4j_py.openpackaging.parts.namespaces import Namespaces

    part = body.part  # set_hyperlink has refused a body with none
    had = getattr(part, "relationships_part", None) is not None
    rels = part.get_relationships_part(True)
    allocator = getattr(rels, "_next_id", None)
    rel = rels.add_external_relationship(Namespaces.HYPERLINK, address)
    note_added_relationship(body.package, rels, rel, created=not had, allocator=allocator)
    change.part(rels)
    return str(rel.id)


def _restyle(paragraph: Paragraph, link: Any, *, linked: bool) -> None:
    """The ``Hyperlink`` style on a link's runs and their direct colour off, or the style off.

    What Word does (docx4j-core-ts CR-002 section 25, check 15): a run's own
    character style is replaced, its ``w:color`` removed and the rest of its
    direct formatting kept; a link removed takes ``Hyperlink`` off the runs that
    carry it. Under change tracking each run's properties as they stood are
    recorded in ``w:rPrChange`` first, as :class:`~.font.Font` records them,
    unless the run is this author's own insertion.
    """
    from docx4j_py.wml import RPr, apply_run_options

    tracking = paragraph.font_tracking()
    runs: list[Any] = []
    seen: set[int] = set()
    for run in (*runs_of(link), *runs_of(link, view="original")):
        if id(run) not in seen:
            seen.add(id(run))
            runs.append(run)
    for run in runs:
        r_pr = getattr(run, "r_pr", None)
        if not linked and _style_of(r_pr) != HYPERLINK_STYLE:
            continue
        if r_pr is None:
            r_pr = RPr()
            run.r_pr = r_pr
            r_pr.parent = run
        if tracking is not None and tracking.records(run):
            tracking.tracker.record_r_pr_change(r_pr)
        if linked:
            apply_run_options(r_pr, style=HYPERLINK_STYLE, color="")
        else:
            apply_run_options(r_pr, style="")
        if _is_empty(r_pr):
            # Word writes no w:rPr where nothing is left (check 15, case E)
            run.r_pr = None
        else:
            link_parents(r_pr)
            r_pr.parent = run


def _style_of(r_pr: Any) -> Any:
    """``w:rStyle/@w:val``, or None."""
    style = getattr(r_pr, "r_style", None) if r_pr is not None else None
    return getattr(style, "val", None)


def _merge_text_runs(items: list, split: dict[int, int] | None) -> None:
    """Adjacent runs of nothing but ``w:t`` with identical properties become one.

    As Word writes a link it makes: linking "beta gamma" over the end of a link
    on "alpha beta" gives one run "alpha beta gamma" (docx4j-core-ts check 16,
    case F), where the split and the restyle left two. Runs whose properties
    differ at all --- a ``w:rPrChange`` under tracking among them --- stay apart.
    """
    from docx4j_py.wml import el

    index = len(items) - 1
    while index > 0:
        first, second = items[index - 1], items[index]
        if _is_text_run(first) and _is_text_run(second) and first.r_pr == second.r_pr:
            text = "".join(str(t.value or "") for t in (*first.content, *second.content))
            first.content[:] = [el.t(text)]
            del items[index]
            _joined(split, first, second)
        index -= 1


def _is_text_run(item: Any) -> bool:
    """A ``w:r`` holding ``w:t`` items and nothing else."""
    if element_name(item) != W_R:
        return False
    return all(element_name(child) == W_T for child in (getattr(item, "content", None) or ()))


def _index_of(items: list, element: Any) -> int:
    """The index of `element` in `items` by identity, or -1."""
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1
