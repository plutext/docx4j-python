"""A story's text in Office JS's ``body.text`` form, with each paragraph's place in it.

CR-013 Phase A (section 4.2: the story layout behind ``Body.text``), after
docx4j-core-ts's ``Body.textIn`` at ``9307bdb`` (its CR-012 phase 4) and
``storyLayout`` at ``ef57729`` (``story.mts``, CR-014 phase 1). The form is
Word's, measured in Word 365 (core-ts checks 30, 31, 48 and 51):

* each paragraph's text (``Paragraph.text``) followed by its mark, ``\\r``;
* in a table, a row's cells a ``\\t`` apart, each cell's paragraphs a mark apart
  with the cell's last mark not written, and the row ended by ``\\r\\n``; a
  row-level content control's rows the same, since :func:`rows_of` reads
  through it;
* a block-level content control, custom XML and a kept block-level
  ``mc:AlternateContent`` (the branch Word shows) transparent;
* nothing of a text box, which is in a ``w:drawing`` and not the story's
  (core-ts check 42).

``Body.text`` is :func:`story_layout`'s text with the last paragraph's mark
left out, so the two cannot disagree. ``story_of`` (the outermost block list
above a paragraph) and a ``Range`` across paragraphs are CR-013 Phase B's.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Callable
from typing import Any

from docx4j_py.model.content.text_model import (
    W_P,
    W_TBL,
    W_TC,
    W_TR,
    block_children_of,
    cells_of,
    rows_of,
)
from docx4j_py.traversal import element_name

__all__ = ["Place", "StoryLayout", "mark_text_of", "story_layout"]

#: A paragraph's mark in a story's text (Office JS's ``\r``).
MARK = "\r"
#: Between two cells of a row.
CELL_END = "\t"
#: After a row's last cell.
ROW_END = "\r\n"


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
    """A story's text in ``body.text``'s form, every paragraph's mark written, and each paragraph's place."""

    text: str
    places: list[Place]


def story_layout(items: list, text_of: Callable[[Any, list], str]) -> StoryLayout:
    """The text of a block list in Office JS's ``body.text`` form, with each paragraph's place.

    Args:
        items: a block-level list: a body's, a cell's, a content control's;
            a list of rows (a row-level control's content) or of cells is laid
            out as a row is.
        text_of: ``text_of(p, container)``, a paragraph's text in the view
            asked for (``Paragraph.text``, or ``get_text(view=...)``).

    Returns:
        The text with **every** paragraph's mark written, the last one's too;
        ``Body.text`` drops a final ``\\r``.
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

    def visit(blocks: list) -> None:
        for element in blocks:
            if isinstance(element, str):
                continue
            name = element_name(element)
            if name == W_P:
                text = text_of(element, blocks)
                places.append(Place(element, blocks, position, position + len(text)))
                push(text)
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

    visit(items)
    return StoryLayout("".join(pieces), places)


def mark_text_of(element: Any) -> str:
    """The character Office JS gives a paragraph's mark at the end of its ``'Whole'`` range.

    core-ts check 48: ``\\t`` for the last paragraph of a table cell, the
    cell's end, and ``\\r`` for any other. The paragraph is a cell's last when,
    climbing through the block holders above it (a content control, custom
    XML, a kept branch), the first ancestor is a ``w:tc`` and the cell's last
    paragraph, read through the same holders and not into a nested table, is
    this one (core-ts ``lastInCell``).
    """
    holder = getattr(element, "parent", None)
    for _ in range(32):
        if holder is None:
            return MARK
        name = element_name(holder)
        if name == W_TC:
            break
        if name in (W_P, W_TBL, W_TR):
            return MARK
        holder = getattr(holder, "parent", None)
    else:
        return MARK
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
    return CELL_END if last[0] is element else MARK
