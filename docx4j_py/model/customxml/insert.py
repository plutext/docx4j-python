"""``insert_content_control(kind)`` on a body, a paragraph and a range.

CR-003 section 3.2 (the member), section 4 (the two refusals) and Phase A's
``sdt`` / ``sdt_pr`` / ``next_sdt_id`` builders, which do the markup.

Three levels, three shapes:

``Body``
    wraps **everything** the body holds in one block-level ``w:sdt``; the body
    is then one block long.
``Paragraph``
    wraps the paragraph, in place, so the container is the same length.
``Range``
    wraps the runs the span covers, **splitting them at the boundaries** as
    ``Range.font`` does, giving a run-level ``w:sdt``. It refuses a span that
    crosses a run holder (``Range.require_one_holder``), and refuses
    ``RepeatingSection``, which Word has no run-level form of.

The ``w:id`` comes from :func:`~docx4j_py.wml.sdt.next_sdt_id` over the part's
whole tree --- not from the change tracker and not from the comment allocator,
which are two other id spaces entirely (CR-003 section 16.6) --- so the same
document and the same calls make the same bytes.

What change tracking does with one, recorded because it is a question a caller
asks: **a content control is not a revision.** Wrapping content in a ``w:sdt``
writes no ``w:ins``, exactly as ``ContentControl.delete()`` writes no ``w:del``
(CR-003 section 4). Paragraphs *inserted* while the mode is on are inserted
content in the ordinary way, and a control put around them holds them; the
control itself is invisible to ``get_tracked_changes()``.

**Across paragraphs** (CR-013 Phase C; docx4j-core-ts check 54 cases 36 to
43): :func:`wrap_paragraphs` puts one block-level control around the
paragraphs a range holds whole, as Word does, and refuses any other range, as
Word refuses it. Under tracking that control is an insertion in Word's own form,
the ``w:customXmlInsRangeStart`` and ``End`` pairs of check 30 around it (case
43), as Word's save after the run has it and core-ts writes it; the
paragraph's, the body's and a run-level control keep CR-003's rule above.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.child import link_parents
from docx4j_py.model.content.errors import BindingError, ContentError
from docx4j_py.traversal import element_name
from docx4j_py.model.content.reports import recording
from docx4j_py.wml.sdt import next_sdt_id, sdt

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.controls import ContentControl

__all__ = [
    "insert_content_control_in_body",
    "insert_content_control_in_paragraph",
    "insert_content_control_in_range",
    "wrap_paragraphs",
]


def _root_of(body: Any) -> Any:
    """The tree :func:`next_sdt_id` scans: the whole part, or the container."""
    part = getattr(body, "part", None)
    contents = getattr(part, "contents", None) if part is not None else None
    return contents if contents is not None else body.container


def _control(body: Any, element: Any, container: list) -> ContentControl:
    from docx4j_py.model.content.controls import ContentControl

    return ContentControl(element, container, body)


def _warn_if_w15(body: Any, kind: str, change: Any) -> None:
    """A repeating section is ``w15``: warn when the document is below mode 15.

    CR-003 section 17.11. Word 2010 does not know ``w15:repeatingSection``; the
    markup is written all the same, and the report says what it needs.
    """
    if kind != "RepeatingSection":
        return
    from docx4j_py.model.content.compatibility import warn_below

    warn_below(
        getattr(body, "package", None), change, "a repeating section (w15:repeatingSection)"
    )


def insert_content_control_in_body(body: Any, kind: str = "RichText") -> ContentControl:
    """Wrap everything the body holds in one content control.

    Office JS ``Body.insertContentControl``.

    Raises:
        ContentError: the body is empty; a control has to wrap something.
    """
    with recording(body, "insert_content_control") as change:
        content = body.content
        if not content:
            raise ContentError(
                "this body holds nothing to wrap in a content control",
                code="control.empty_body",
                hint="insert a paragraph first, then wrap it",
            )
        _warn_if_w15(body, kind, change)
        identifier = next_sdt_id(_root_of(body))
        items = list(content)
        control = sdt(items, kind=kind, id=identifier, form="block")
        del content[:]
        content.append(control)
        link_parents(control)
        view = _control(body, control, content)
        change.touched(f"{body.prefix}/0")
        change.text(after=view.text)
        return view


def insert_content_control_in_paragraph(paragraph: Any, kind: str = "RichText") -> ContentControl:
    """Wrap the paragraph in a content control, in place.

    Office JS ``Paragraph.insertContentControl``.

    Raises:
        ContentError: the paragraph is no longer in its container.
    """
    body = paragraph.parent_body
    with recording(body, "insert_content_control") as change:
        container = paragraph.container
        index = next((i for i, item in enumerate(container) if item is paragraph.element), -1)
        if index < 0:
            raise ContentError(
                "this paragraph is no longer in its container",
                code="address.not_found",
                hint="read it again from body.paragraphs",
            )
        _warn_if_w15(body, kind, change)
        identifier = next_sdt_id(_root_of(body))
        control = sdt([paragraph.element], kind=kind, id=identifier, form="block")
        container[index] = control
        link_parents(control)
        view = _control(body, control, container)
        change.touched(view.address)
        change.text(after=view.text)
        return view


def insert_content_control_in_range(span: Any, kind: str = "RichText") -> ContentControl:
    """Wrap the runs a span covers in a run-level content control.

    Office JS ``Range.insertContentControl``. The runs are split at the span's
    boundaries first, on grapheme boundaries (CR-003 section 3.12), so the
    control holds exactly the span's text.

    Raises:
        SpanError: the span crosses a run holder; the message names where to
            split.
        BindingError: ``RepeatingSection`` was asked for, which Word has no
            run-level form of.
    """
    from docx4j_py.traversal import run_items_of

    if kind == "RepeatingSection":
        raise BindingError(
            "a repeating section is a block-level control, not a run-level one",
            code="binding.form_mismatch",
            hint="call insert_content_control on a paragraph or on the body",
        )
    span.refuse_branch_text("insert_content_control")
    span.require_one_holder("insert_content_control")
    _require_no_revision_between(span)

    paragraph = span.paragraph
    body = paragraph.parent_body
    with recording(body, "insert_content_control") as change:
        if span.start != span.end:
            span.start = paragraph.split_at(span.start, prefer="back")
            span.end = paragraph.split_at(span.end, prefer="forward")
        runs = span.runs
        holder = runs[0].parent if runs else paragraph.element
        items = run_items_of(holder)
        if items is None:
            items = run_items_of(paragraph.element)
        if items is None:  # pragma: no cover - a paragraph always holds runs
            raise ContentError(
                "this span is not in a run holder",
                code="range.no_holder",
                hint="use get_range() on a paragraph that holds runs",
            )
        identifier = next_sdt_id(_root_of(body))
        if runs:
            # the whole slice from the first run to the last, so that what lies
            # between them --- a bookmark, a comment's range markers --- goes into
            # the control and keeps its place (docx4j-core-ts CR-002 section 43)
            first = _position(items, runs[0])
            last = _position(items, runs[-1])
            control = sdt(list(items[first : last + 1]), kind=kind, id=identifier, form="run")
            items[first : last + 1] = [control]
        else:
            control = sdt([], kind=kind, id=identifier, form="run")
            items.insert(_offset_index(items, paragraph, span.start), control)
        link_parents(control)
        view = _control(body, control, items)
        change.touched(paragraph)
        change.text(after=view.text)
        return view


def wrap_paragraphs(span: Any, kind: str = "RichText") -> ContentControl:
    """A block content control around the paragraphs of a range across paragraphs, as Word puts one.

    docx4j-core-ts's ``wrapParagraphs`` (``Range.mts`` at ``87bbb38``; Word's
    check 54 cases 38 and 43, untracked and tracked): only a range holding its
    paragraphs whole, from the first's start to the last's end, its mark or
    not, in one list of blocks. Tracked, the control is an insertion in Word's
    form (:func:`~docx4j_py.model.content.tracking.mark_holder_inserted`).

    Raises:
        SpanError: ``range.partial_paragraph``, a range starting or ending
            inside a paragraph, across a table, into one or out of one, as Word
            refused it (``GeneralException``, cases 36, 37 and 39 to 42);
            nothing has changed.
    """
    from docx4j_py.model.content.errors import SpanError
    from docx4j_py.model.content.reports import refuse_read_only
    from docx4j_py.model.content.span_edits import extent_of
    from docx4j_py.model.content.tracking import mark_holder_inserted

    def refuse(why: str) -> SpanError:
        return SpanError(
            f"insert_content_control refused: a range across paragraphs takes a control only where it "
            f"holds its paragraphs whole, and this one {why}, which Word refuses too (core-ts check 54)",
            code="range.partial_paragraph",
            hint="take the paragraphs' whole range (from the first's get_range('Start') to the last's "
            "get_range('Whole')), or wrap each paragraph (paragraph.insert_content_control())",
        )

    for paragraph in (span.paragraph, span.end_paragraph):
        refuse_read_only(paragraph.parent_body, "insert_content_control")
    if span.start != 0 or span.end < len(span.end_paragraph.text):
        raise refuse("starts or ends inside a paragraph")
    try:
        extent = extent_of(span, "insert_content_control")
    except SpanError as error:
        raise refuse("starts or ends inside a table or a content control") from error
    items, first, last = extent.items, extent.first, extent.last
    if extent.head is None or extent.tail is None or any(
        element_name(items[index]) == _W_TBL for index in range(first, last + 1)
    ):
        raise refuse("holds a table")
    for paragraph in span.paragraphs:
        refuse_read_only(paragraph.parent_body, "insert_content_control")
    body = span.paragraph.parent_body
    with recording(body, "insert_content_control") as change:
        _warn_if_w15(body, kind, change)
        identifier = next_sdt_id(_root_of(body))
        parent = getattr(items[first], "parent", None)
        control = sdt(list(items[first : last + 1]), kind=kind, id=identifier, form="block")
        items[first : last + 1] = [control]
        link_parents(control)
        control.parent = parent
        tracker = span.paragraph.change_tracker
        if tracker is not None:
            mark_holder_inserted(tracker, control, items)
        view = _control(body, control, items)
        change.touched(view.address)
        change.text(after=view.text)
        return view


#: A table, ``w:tbl``.
_W_TBL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"


def _require_no_revision_between(span: Any) -> None:
    """Refuse, before anything changes, a tracked change between the span's first run and its last.

    A ``w:del`` or ``w:moveFrom`` there has no text in the accepted view, so
    :meth:`Range.require_one_holder` does not see it; wrapping the slice would
    move it into the control. ``Range.hyperlink`` refuses the same span
    (CR-003 section 21.3 item 8), as docx4j-core-ts does since its CR-002
    section 43.
    """
    from docx4j_py.model.content.errors import SpanError
    from docx4j_py.model.content.tracking import REVISION_NAMES
    from docx4j_py.traversal import element_name

    segments = [
        segment
        for segment in span.paragraph.segments()
        if segment.end > span.start and segment.start < span.end
    ]
    if not segments:
        return
    owner = segments[0].run_owner
    first = _position(owner, segments[0].run)
    last = _position(owner, segments[-1].run)
    for item in owner[first : last + 1]:
        name = element_name(item)
        if name in REVISION_NAMES:
            raise SpanError(
                f"the range {span.start}:{span.end} holds a tracked change "
                f"(w:{(name or '').rpartition('}')[2]}) between its runs, "
                "which a content control would have to wrap",
                code="range.crosses_holder",
                hint="accept or reject the tracked change first, or wrap a span on one side of it",
            )


def _position(items: list, element: Any) -> int:
    """Where `element` is in `items`, by identity."""
    for position, item in enumerate(items):
        if item is element:
            return position
    raise ValueError("the element is not in this list")  # pragma: no cover - found through it


def _offset_index(items: list, paragraph: Any, offset: int) -> int:
    """Where an **empty** control goes: after the last run ending at or before `offset`.

    Then past anything already sitting there that carries no text --- an empty
    control inserted a moment ago, a bookmark --- so that two empty controls
    asked for at the same offset come out in the order they were asked for.
    """
    from docx4j_py.traversal import text_of

    index = 0
    for segment in paragraph.segments():
        if segment.end > offset:
            break
        # the run, or the kept element its text is read from (CR-011 Phase F)
        site = segment.site()[0]
        for position, item in enumerate(items):
            if item is site:
                index = position + 1
                break
    while index < len(items) and not text_of(items[index]):
        index += 1
    return index
