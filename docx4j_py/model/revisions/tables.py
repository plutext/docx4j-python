"""A table's property records, listed. docx4j ``Tables`` (the listing half).

CR-010 section 4 and section 5: Phase A **lists** a table's ``w:tblPrChange``,
``w:tblGridChange``, ``w:tblPrExChange``, ``w:trPrChange`` and ``w:tcPrChange``
as one change per table, as docx4j's phase 1 did, so that the listing never
reports a document clean while Word still shows the records. Resolving them ---
accepted dropped, rejected restored wholesale and a merged grid collapsed --- is
Phase B's (docx4j CR-043 phase 2, ``Tables.accept`` / ``reject``).

The rule here is docx4j phase 1's: a table with any record is listed. Phase 2's
refinement --- records that change nothing are not listed, and swept --- comes
with Phase B, because the sweep that drops them is Phase B's too.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.revisions import trees
from docx4j_py.model.revisions.trees import View

__all__ = ["cells_of", "change_text", "has_records", "record_of", "rows_of"]


def rows_of(tbl: Any) -> list[tuple[Any, list]]:
    """A table's rows and the live list holding each, through row-level controls and custom XML."""
    out: list[tuple[Any, list]] = []
    _rows(tbl.content, out)
    return out


def _rows(items: list, out: list[tuple[Any, list]]) -> None:
    tr_class, tc_class = trees.cls("Tr"), trees.cls("Tc")
    for item in items:
        if isinstance(item, tr_class):
            out.append((item, items))
        elif not isinstance(item, tc_class):
            nested = trees.block_items(item)
            if nested is not None:
                _rows(nested, out)


def cells_of(tr: Any) -> list[tuple[Any, list]]:
    """A row's cells and the live list holding each, through cell-level controls and custom XML."""
    out: list[tuple[Any, list]] = []
    _cells(tr.content, out)
    return out


def _cells(items: list, out: list[tuple[Any, list]]) -> None:
    tc_class = trees.cls("Tc")
    for item in items:
        if isinstance(item, tc_class):
            out.append((item, items))
        else:
            nested = trees.block_items(item)
            if nested is not None:
                _cells(nested, out)


def _row_record(tr: Any) -> Any:
    tr_pr = tr.tr_pr
    if tr_pr is not None and tr_pr.tr_pr_change is not None:
        return tr_pr.tr_pr_change
    ex = tr.tbl_pr_ex
    if ex is not None and ex.tbl_pr_ex_change is not None:
        return ex.tbl_pr_ex_change
    for tc, _owner in cells_of(tr):
        tc_pr = tc.tc_pr
        if tc_pr is not None and tc_pr.tc_pr_change is not None:
            return tc_pr.tc_pr_change
    return None


def record_of(tbl: Any) -> Any:
    """The record a table's change reports: the table's, else the first row's or cell's.

    None for a grid's record alone, which has no author of its own.
    """
    tbl_pr = tbl.tbl_pr
    if tbl_pr is not None and tbl_pr.tbl_pr_change is not None:
        return tbl_pr.tbl_pr_change
    for tr, _owner in rows_of(tbl):
        found = _row_record(tr)
        if found is not None:
            return found
    return None


def has_records(tbl: Any) -> bool:
    """True when a table has any property record (core-ts ``hasTableRecords``)."""
    grid = tbl.tbl_grid
    if grid is not None and grid.tbl_grid_change is not None:
        return True
    return record_of(tbl) is not None


def change_text(tbl: Any) -> str:
    """A table's change's text: its rows', as Office JS gives a table's (checks 27, 28)."""
    from docx4j_py.model.revisions.texts import row_text

    return "".join(row_text(tr, View.ACCEPTED) for tr, _owner in rows_of(tbl))
