"""A tracked change's text, as Office JS gives it. docx4j ``Texts``.

CR-010 section 3.1; core-ts ``TrackedChange.text``, checks 22, 27, 30 and 31. A
paragraph mark is ``\\r``; a row is its cells a tab apart, then ``\\r\\n``; a
section's change is its paragraphs' text, then ``\\f``; a deletion's is the
deleted text; a field's is its result.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.revisions import trees
from docx4j_py.model.revisions.change import (
    CellT,
    FieldT,
    GroupT,
    MarkPropsT,
    MarkT,
    MathRunT,
    ParaPropsT,
    RowT,
    RunPropsT,
    RunT,
    SectionT,
    TableT,
    TrackedChange,
)
from docx4j_py.model.revisions.trees import Revision, View

__all__ = ["row_text", "text_of"]


def text_of(change: TrackedChange) -> str:
    """The text of one change."""
    t = change.target
    if isinstance(t, RunT):
        view = View.ACCEPTED if t.revision in (Revision.INS, Revision.MOVE_TO) else View.ORIGINAL
        return trees.text(t.item, view)
    if isinstance(t, (MarkT, MarkPropsT)):
        return "\r"
    if isinstance(t, GroupT):
        return "".join(piece.text for piece in t.pieces)
    if isinstance(t, RunPropsT):
        return trees.text(t.run, View.ACCEPTED)
    if isinstance(t, FieldT):
        out: list[str] = []
        for item in t.span:
            trees.append_text(item, View.ACCEPTED, out)
        return "".join(out)
    if isinstance(t, ParaPropsT):
        return "" if change.para is None else trees.paragraph_text(change.para.p, View.ACCEPTED)
    if isinstance(t, RowT):
        return row_text(t.tr, View.ACCEPTED if t.ins else View.ORIGINAL)
    if isinstance(t, MathRunT):
        out = []
        for item in t.item.content:
            found = trees.item_text(item)
            if found is not None:
                out.append(found)
        return "".join(out)
    if isinstance(t, TableT):
        from docx4j_py.model.revisions.tables import change_text

        return change_text(t.tbl)
    if isinstance(t, CellT):
        return row_text(t.tr, View.ACCEPTED)
    if isinstance(t, SectionT):
        return t.text
    return ""


def row_text(tr: Any, view: View) -> str:
    """A row's text: its cells a tab apart, then ``\\r\\n``.

    A cell's paragraphs ``\\r`` apart; a nested table's rows each with their own
    ``\\r\\n`` (checks 30 and 31).
    """
    out: list[str] = []
    first = True
    tc_class = trees.cls("Tc")
    for item in tr.content:
        if not isinstance(item, tc_class):
            continue
        if not first:
            out.append("\t")
        first = False
        _cell_text(item.content, view, out, [False])
    out.append("\r\n")
    return "".join(out)


def _cell_text(items: list, view: View, out: list[str], after_paragraph: list[bool]) -> None:
    p_class, tbl_class, tr_class = trees.cls("P"), trees.cls("Tbl"), trees.cls("Tr")
    for item in items:
        if isinstance(item, p_class):
            if after_paragraph[0]:
                out.append("\r")
            out.append(trees.paragraph_text(item, view))
            after_paragraph[0] = True
        elif isinstance(item, tbl_class):
            for row in item.content:
                if isinstance(row, tr_class):
                    out.append(row_text(row, view))
            after_paragraph[0] = False
        else:
            nested = trees.block_items(item)
            if nested is not None:
                _cell_text(nested, view, out, after_paragraph)
