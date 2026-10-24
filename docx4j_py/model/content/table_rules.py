"""Word's rule for what a table cell holds: it ends with a paragraph, as Word reads it.

CR-016 Phase A (section 4.2; Q3's D1 to D4, decided 2026-10-24 as
recommended), the port of docx4j-core-ts's ``tableRules.mts`` at ``222cc20``
(its CR-002 section 55) and docx4j's ``org.docx4j.model.table.TableRules`` at
``18a3a31f4`` (``VERSION_17_3_4``), their names in Python's spelling. The rule
is docx4j's, measured on Word 365 in twelve probes
(``docs/developer/word-table-rules.md`` there):

* a cell holds a block;
* its last block, as Word reads it, is a ``w:p``: looked for past range markup
  (bookmarks, comment and move ranges, permissions), through block-level
  content controls and custom XML, and through the one branch of a kept
  ``mc:AlternateContent`` Word reads (:func:`~docx4j_py.child.mce_branch`, the
  branch the readers take: CR-011 Phase F, CR-012 Phase A); the branches Word
  does not read are not checked;
* a cell that breaks it opens as unreadable content, and Word's repair appends
  an empty ``w:p`` to the cell, after any wrapper. Word's own edits never break
  it.

:func:`ends_with_paragraph` and :func:`end_with_paragraph` are docx4j's two;
the rest is what the content API's edits use to keep the rule
(:class:`~docx4j_py.model.content.body.Body`,
:class:`~docx4j_py.model.content.paragraph.Paragraph`,
:class:`~docx4j_py.model.content.controls.ContentControl`, and
:func:`~docx4j_py.model.content.span_edits.keep_cell_mark_last`). Nothing here
runs on load: a file that already breaks the rule is left as it was read until
an edit touches that cell (D4).
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.child import is_alternate_content, mce_choices, mce_fallback
from docx4j_py.model.content.text_model import block_children_of, is_branch
from docx4j_py.wml import (
    CTCustomXmlBlock,
    CTCustomXmlCell,
    CTCustomXmlRow,
    CTSdtCell,
    CTSdtRow,
    P,
    SdtBlock,
    SdtContentBlock,
    Tbl,
    Tc,
    Tr,
)

__all__ = [
    "CellEnd",
    "end_cells_with_paragraphs",
    "end_with_paragraph",
    "ends_cell",
    "ends_with_paragraph",
    "paragraph_end_of",
    "wrapping_cell_of",
]

#: How far up, or down, a climb through the wrappers goes (core-ts's guard).
_MAX_DEPTH = 32

#: What counts as a block among a cell's contents; anything else (range markup) does not.
_BLOCKS: tuple[type, ...] = (P, Tbl, SdtBlock, CTCustomXmlBlock)

#: What stands between a cell and the blocks Word reads through: a block control, its content, custom XML.
#: A kept ``mc:AlternateContent`` and its branches are wrappers too, found by :func:`_is_wrapper`.
_WRAPPERS: tuple[type, ...] = (SdtBlock, SdtContentBlock, CTCustomXmlBlock)

#: Between a table and its rows, a row and its cells: a row- or cell-level control or custom XML.
_ROW_AND_CELL_WRAPPERS: tuple[type, ...] = (CTSdtRow, CTSdtCell, CTCustomXmlRow, CTCustomXmlCell)


@dataclasses.dataclass(frozen=True, slots=True)
class CellEnd:
    """The paragraph a cell ends with, as Word reads it, and the live list it is in (core-ts's ``CellEnd``)."""

    element: P
    """The ``w:p``."""
    container: list
    """The list holding it: the cell's own, a control's content, custom XML's, a branch's."""


def _is_block(item: Any) -> bool:
    return isinstance(item, _BLOCKS) or is_alternate_content(item)


def _is_wrapper(node: Any) -> bool:
    return isinstance(node, _WRAPPERS) or is_alternate_content(node) or is_branch(node)


def _last_block(items: list | None) -> Any:
    """The last block-level item of a list, or None where it holds none."""
    for item in reversed(items or ()):
        if _is_block(item):
            return item
    return None


def _contents_of(contents: Any) -> list | None:
    """A cell's content list, for a ``w:tc`` or the list itself (docx4j's two overloads)."""
    return getattr(contents, "content", None) if isinstance(contents, Tc) else contents


def paragraph_end_of(contents: Any) -> CellEnd | None:
    """The paragraph these cell contents end with, as Word reads them, and its list; None where there is none.

    The last block past range markup, through block-level controls and custom
    XML, and through the branch of a kept ``mc:AlternateContent`` a reader takes;
    None where the last block, so read, is a table, or there is no block at all.
    `contents` is a cell's list, or the ``w:tc`` itself.
    """
    items = _contents_of(contents)
    for _ in range(_MAX_DEPTH):
        last = _last_block(items)
        if last is None or isinstance(last, Tbl):
            return None
        if isinstance(last, P):
            return CellEnd(last, items)  # type: ignore[arg-type]
        # a control's content, custom XML's, or the branch Word reads of a kept mc:AlternateContent
        items = block_children_of(last)
    return None


def ends_with_paragraph(contents: Any) -> bool:
    """docx4j's ``endsWithParagraph``: whether the last block, as Word reads it, is a paragraph.

    `contents` is a cell's content list, a ``w:tc``, or any list of block-level
    items (what an insert is about to put in). Range markup after the last
    block does not count; a list with no block at all is False.
    """
    return paragraph_end_of(contents) is not None


def end_with_paragraph(tc: Tc) -> P | None:
    """docx4j's ``endWithParagraph``, Word's repair: an empty ``w:p`` appended to the cell where it doesn't end with one.

    The paragraph goes after any wrapper, never inside it, and nothing is done
    where the rule holds, so a cell ending in a wrapped paragraph gets no spare
    one. Returns the paragraph appended (docx4j returns True), so that a caller
    can give it a ``w14:paraId``; None where the rule held.
    """
    if ends_with_paragraph(tc):
        return None
    p = P()
    tc.content.append(p)
    p.parent = tc
    return p


def wrapping_cell_of(owner: Any) -> Tc | None:
    """The cell whose content holds this container: up the parent pointers past the wrappers, or None.

    `owner` is what holds a block list: a ``w:tc`` (itself), a control or its
    ``w:sdtContent``, custom XML, an ``mc:AlternateContent`` branch. A paragraph
    in a text box, a header or the body has none.
    """
    node = owner
    for _ in range(_MAX_DEPTH):
        if node is None:
            return None
        if isinstance(node, Tc):
            return node
        if not _is_wrapper(node):
            return None
        node = getattr(node, "parent", None)
    return None


def ends_cell(paragraph: P) -> bool:
    """Whether this paragraph is the one its cell ends with, as Word reads it, through the wrappers."""
    cell = wrapping_cell_of(getattr(paragraph, "parent", None))
    if cell is None:
        return False
    end = paragraph_end_of(cell)
    return end is not None and end.element is paragraph


def _branch_lists(ac: Any) -> list[list]:
    """Every branch's list of a typed ``mc:AlternateContent``, Word's or not."""
    branches = [*mce_choices(ac), mce_fallback(ac)]
    return [items for branch in branches if isinstance(items := getattr(branch, "any_element", None), list)]


def _every(items: list | None, wanted: type, out: list) -> None:
    """Each `wanted` (a row or a cell) in a table's or a row's list, through the wrappers and every branch."""
    for item in items or ():
        if isinstance(item, wanted):
            out.append(item)
        elif is_alternate_content(item):
            for branch in _branch_lists(item):
                _every(branch, wanted, out)
        elif isinstance(item, _ROW_AND_CELL_WRAPPERS):
            _every(block_children_of(item), wanted, out)


def end_cells_with_paragraphs(blocks: list) -> list[P]:
    """D3: every cell of the tables in these blocks, at any depth, ends with a paragraph (docx4j's repair).

    For content that goes in through an insert, before it goes in, so that
    under change tracking the repair is part of the insertion. Nested tables
    first; every branch of a kept ``mc:AlternateContent``; nothing inside a
    paragraph. Returns the paragraphs appended.
    """
    appended: list[P] = []

    def visit(items: list | None) -> None:
        for item in items or ():
            if isinstance(item, (P, str)):
                continue
            if isinstance(item, Tbl):
                rows: list[Tr] = []
                _every(item.content, Tr, rows)
                for row in rows:
                    cells: list[Tc] = []
                    _every(row.content, Tc, cells)
                    for cell in cells:
                        visit(cell.content)
                        p = end_with_paragraph(cell)
                        if p is not None:
                            appended.append(p)
            elif is_alternate_content(item):
                for branch in _branch_lists(item):
                    visit(branch)
            else:
                visit(block_children_of(item))

    visit(blocks)
    return appended
