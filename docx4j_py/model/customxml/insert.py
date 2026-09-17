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
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.child import link_parents
from docx4j_py.model.content.errors import BindingError, ContentError
from docx4j_py.model.content.reports import recording
from docx4j_py.wml.sdt import next_sdt_id, sdt

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.controls import ContentControl

__all__ = [
    "insert_content_control_in_body",
    "insert_content_control_in_paragraph",
    "insert_content_control_in_range",
]


def _root_of(body: Any) -> Any:
    """The tree :func:`next_sdt_id` scans: the whole part, or the container."""
    part = getattr(body, "part", None)
    contents = getattr(part, "contents", None) if part is not None else None
    return contents if contents is not None else body.container


def _control(body: Any, element: Any, container: list) -> ContentControl:
    from docx4j_py.model.content.controls import ContentControl

    return ContentControl(element, container, body)


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
    span.require_one_holder("insert_content_control")

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
        index = len(items)
        for run in runs:
            for position, item in enumerate(items):
                if item is run:
                    index = min(index, position)
                    break
        if not runs:
            index = _offset_index(items, paragraph, span.start)
        identifier = next_sdt_id(_root_of(body))
        control = sdt(list(runs), kind=kind, id=identifier, form="run")
        for run in runs:
            items.remove(run)
        items.insert(index, control)
        link_parents(control)
        view = _control(body, control, items)
        change.touched(paragraph)
        change.text(after=view.text)
        return view


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
        for position, item in enumerate(items):
            if item is segment.run:
                index = position + 1
                break
    while index < len(items) and not text_of(items[index]):
        index += 1
    return index
