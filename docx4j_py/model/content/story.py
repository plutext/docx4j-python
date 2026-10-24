"""A story's text in Office JS's ``body.text`` form, with each paragraph's place in it.

CR-013 Phase A (section 4.2: the story layout behind ``Body.text``), after
docx4j-core-ts's ``Body.textIn`` at ``9307bdb`` (its CR-012 phase 4) and
``storyLayout`` at ``ef57729`` (``story.mts``, CR-014 phase 1). The form is
Word's, measured in Word 365 (core-ts checks 30, 31, 48 and 51):

* each paragraph's text (``Paragraph.text``) followed by its mark, ``\\r``,
  but a mark tracked as deleted, which Word's default view shows as nothing,
  the paragraphs staying apart (CR-016 Phase B; core-ts ``222cc20``, its CR-012
  section 19, check 58 cases 6 to 10); the original view writes no inserted
  mark instead (:func:`shown_in_original_view`);
* in a table, a row's cells a ``\\t`` apart, each cell's paragraphs a mark apart
  with the cell's last mark not written, and the row ended by ``\\r\\n``; a
  row-level content control's rows the same, since :func:`rows_of` reads
  through it;
* a block-level content control, custom XML and a kept block-level
  ``mc:AlternateContent`` (the branch Word shows) transparent;
* nothing of a text box, which is in a ``w:drawing`` and not the story's
  (core-ts check 42).

``Body.text`` is :func:`story_layout`'s text with the last paragraph's mark
left out, so the two cannot disagree.

CR-013 Phase B (``story.mts`` at ``ef57729``): :func:`story_of` is the
outermost block list above a paragraph, which a ``Range`` across paragraphs
reads a slice of. A text box's content is a story of its own, and so is a
branch of a kept ``mc:AlternateContent`` that Word does not read; Word's
branch is part of the story around it. :func:`top_index` and the `start` and
`stop` of :func:`story_layout` lay out only the top-level blocks a range
touches: each top-level block's text is its own, so a slice's layout is the
story's, shifted.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Callable
from typing import Any

from docx4j_py.child import mce_branch
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.text_model import (
    W_P,
    W_TBL,
    W_TC,
    W_TR,
    block_children_of,
    cells_of,
    is_block_container,
    is_branch,
    rows_of,
)
from docx4j_py.traversal import element_name

__all__ = [
    "Place",
    "StoryLayout",
    "first_and_last_paragraph",
    "mark_text_of",
    "shown_in_default_view",
    "shown_in_original_view",
    "story_layout",
    "story_of",
    "top_index",
]

#: A paragraph's mark in a story's text (Office JS's ``\r``).
MARK = "\r"
#: Between two cells of a row.
CELL_END = "\t"
#: After a row's last cell.
ROW_END = "\r\n"


def _mark(element: Any) -> Any:
    """A paragraph's mark properties, ``w:pPr/w:rPr``, or None."""
    p_pr = getattr(element, "p_pr", None)
    return getattr(p_pr, "r_pr", None) if p_pr is not None else None


def shown_in_default_view(element: Any) -> bool:
    """Whether a paragraph's mark is written in Word's default view: not where it is tracked as deleted.

    CR-016 Phase B, core-ts ``shownInDefaultView`` (``story.mts`` at
    ``222cc20``; its CR-012 section 19): a deleted mark reads as nothing there,
    the paragraphs staying apart (check 58 cases 6 to 8). A cell's last
    paragraph is ended by the cell's ``\\t`` whatever its mark (cases 9 and 10).
    """
    mark = _mark(element)
    return mark is None or getattr(mark, "del_value", None) is None


def shown_in_original_view(element: Any) -> bool:
    """Whether a paragraph's mark is written in the original view: not where it is tracked as inserted.

    The marks Reject All keeps: a deleted mark written, an inserted one not
    (core-ts ``Body.textIn`` at ``222cc20``; CR-016 Phase B).
    """
    mark = _mark(element)
    return mark is None or getattr(mark, "ins", None) is None


#: The mark rule of each view of :func:`story_layout` (``"accepted"`` is Word's default view's).
MARK_SHOWN: dict[str, Callable[[Any], bool]] = {
    "accepted": shown_in_default_view,
    "original": shown_in_original_view,
}


@dataclasses.dataclass(frozen=True, slots=True)
class Place:
    """Where one paragraph's text stands in a story's text."""

    element: Any
    """The ``w:p``."""
    container: list
    """The live list holding it."""
    start: int
    """The offset of its first character in the story's text."""
    end: int
    """One past its last character, where its mark (or the cell's end) begins."""


@dataclasses.dataclass(frozen=True, slots=True)
class StoryLayout:
    """A story's text in ``body.text``'s form, each paragraph's mark written as its view shows it, and each paragraph's place."""

    text: str
    places: list[Place]

    def index_of(self, element: Any) -> int:
        """The index in :attr:`places` of a ``w:p``, by identity; -1 where the walk did not meet it."""
        for index, place in enumerate(self.places):
            if place.element is element:
                return index
        return -1


def story_layout(
    items: list,
    text_of: Callable[[Any, list], str],
    *,
    start: int = 0,
    stop: int | None = None,
    mark_shown: Callable[[Any], bool] = shown_in_default_view,
) -> StoryLayout:
    """The text of a block list in Office JS's ``body.text`` form, with each paragraph's place.

    Args:
        items: a block-level list: a body's, a cell's, a content control's;
            a list of rows (a row-level control's content) or of cells is laid
            out as a row is.
        text_of: ``text_of(p, container)``, a paragraph's text in the view
            asked for (``Paragraph.text``, or ``get_text(view=...)``).
        start: the first of `items` to lay out (CR-013 Phase B).
        stop: one past the last, or None for the rest. Each top-level block's
            text is its own, so ``items[start:stop]``'s layout is the whole
            list's slice, shifted; the places keep `items` as the container of
            a top-level paragraph.
        mark_shown: whether a paragraph's mark is written; Word's default
            view's rule (:func:`shown_in_default_view`, no deleted mark) unless
            another view's is given (:data:`MARK_SHOWN`; CR-016 Phase B).

    Returns:
        The text with every paragraph's mark the view shows written, the last
        one's too; ``Body.text`` drops a final ``\\r``.
    """
    pieces: list[str] = []
    places: list[Place] = []
    position = 0

    def push(piece: str) -> None:
        nonlocal position
        pieces.append(piece)
        position += len(piece)

    def pop_mark() -> None:
        nonlocal position
        if pieces and pieces[-1] == MARK:
            pieces.pop()
            position -= len(MARK)

    def row(element: Any) -> None:
        for index, (cell, _owner) in enumerate(cells_of(element)):
            if index > 0:
                push(CELL_END)
            visit(block_children_of(cell) or [])
            # the cell's last mark is the cell's end, not written
            pop_mark()
        push(ROW_END)

    def visit(blocks: list, low: int = 0, high: int | None = None) -> None:
        for element in blocks if low == 0 and high is None else blocks[low:high]:
            if isinstance(element, str):
                continue
            name = element_name(element)
            if name == W_P:
                text = text_of(element, blocks)
                places.append(Place(element, blocks, position, position + len(text)))
                push(text)
                if mark_shown(element):
                    push(MARK)
                continue
            if name == W_TBL:
                for tr, _owner in rows_of(element):
                    row(tr)
                continue
            if name == W_TR:
                row(element)
                continue
            if name == W_TC:
                visit(block_children_of(element) or [])
                pop_mark()
                continue
            children = block_children_of(element)
            if children is not None:
                visit(children)

    visit(items, start, stop)
    return StoryLayout("".join(pieces), places)


def story_of(element: Any) -> tuple[list, Any]:
    """The story a paragraph stands in: the outermost block list above it, and the object holding that list.

    core-ts ``storyOf`` (``story.mts``, ``ef57729``): up the parent pointers
    through the block containers (a cell, a row, a table, a control's content,
    custom XML), to a body, a header or footer, a note, a comment or a text
    box's content. A paragraph above it ends the climb: the paragraph is in a
    text box, whose content is its story. A branch of a kept
    ``mc:AlternateContent`` that Word does not read is a story of its own;
    Word's branch is part of the story around it, as ``Body.paragraphs``
    reads it (CR-012 Phase A).

    Raises:
        ContentError: ``range.no_story``, a paragraph in no block container
            (one taken out of its document).
    """
    story: tuple[list, Any] | None = None
    node = getattr(element, "parent", None)
    for _ in range(256):
        if node is None:
            break
        if element_name(node) == W_P:
            break
        if is_branch(node):
            ac = getattr(node, "parent", None)
            items = getattr(node, "any_element", None)
            if ac is not None and isinstance(items, list) and mce_branch(ac) is not node:
                return items, node
        elif is_block_container(node):
            items = block_children_of(node)
            if isinstance(items, list):
                story = (items, node)
        node = getattr(node, "parent", None)
    if story is None:
        raise ContentError(
            "this paragraph is in no story: it is not in a body, a header, a footer, a note, "
            "a comment or a text box",
            code="range.no_story",
            hint="take a range from a paragraph of the document",
        )
    return story


def top_index(element: Any, items: list, owner: Any) -> int:
    """The index in `items`, a story's list held by `owner`, of the top-level block holding `element`.

    The block is `element` itself or the ancestor whose parent is `owner`,
    found by identity; -1 where the parent pointers do not reach `owner`.
    """
    node = element
    for _ in range(256):
        if node is None:
            return -1
        parent = getattr(node, "parent", None)
        if parent is owner:
            for index, item in enumerate(items):
                if item is node:
                    return index
            return -1
        node = parent
    return -1


def first_and_last_paragraph(container: list, index: int) -> tuple[Place | None, Place | None]:
    """The first and the last paragraph's :class:`Place` of the block at ``container[index]``, as a reader reads it, or Nones.

    Tables and controls are descended into; the places hold the live lists,
    `container` itself for a block that is a paragraph.
    """
    places = story_layout(container, lambda _element, _container: "", start=index, stop=index + 1).places
    if not places:
        return None, None
    return places[0], places[-1]


def mark_text_of(element: Any) -> str:
    """The character Office JS gives a paragraph's mark at the end of its ``'Whole'`` range.

    core-ts check 48: ``\\t`` for the last paragraph of a table cell, the
    cell's end, whatever its mark; ``""`` for a mark tracked as deleted, which
    Word's default view does not show (CR-016 Phase B; core-ts ``markText`` at
    ``222cc20``, check 58); and ``\\r`` for any other. The paragraph is a cell's last when,
    climbing through the block holders above it (a content control, custom
    XML, a kept branch), the first ancestor is a ``w:tc`` and the cell's last
    paragraph, read through the same holders and not into a nested table, is
    this one (core-ts ``lastInCell``).
    """
    other = MARK if shown_in_default_view(element) else ""
    holder = getattr(element, "parent", None)
    for _ in range(32):
        if holder is None:
            return other
        name = element_name(holder)
        if name == W_TC:
            break
        if name in (W_P, W_TBL, W_TR):
            return other
        holder = getattr(holder, "parent", None)
    else:
        return other
    last: list[Any] = [None]

    def visit(blocks: list) -> None:
        for block in blocks:
            name = element_name(block)
            if name == W_P:
                last[0] = block
            elif name != W_TBL:
                children = block_children_of(block)
                if children is not None:
                    visit(children)

    visit(block_children_of(holder) or [])
    return CELL_END if last[0] is element else other
