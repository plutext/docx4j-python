"""A table's property records, cells inserted and deleted, a section's record. docx4j ``Tables``.

CR-010 section 5 (Phase B), docx4j CR-043 phase 2 (its section 13): a table's
``w:tblPrChange``, ``w:tblGridChange``, ``w:tblPrExChange``, ``w:trPrChange``
and ``w:tcPrChange``, which Word writes as one set for one change and Office JS
lists as one change per table; ``w:cellIns`` / ``w:cellDel`` (check 27) and
``w:cellMerge`` (U8); and a section's ``w:sectPrChange`` (C07, check 18). A port
of docx4j's ``Tables`` (itself core-ts's ``tableRevisions.mts`` with what Word's
saves of C23 add):

* **listed** only when the records change something, compared as canonical XML
  with the markers and the records left out (checks 24 and 26's rule, U20);
  records that change nothing are swept;
* **accepted**, every record goes, and a ``w:trPr`` left empty with it;
* **rejected**, the table, its grid, and each row and cell in the change are
  restored wholesale --- a row in the change with no ``w:trPrChange`` cleared
  (check 28) --- and a merged grid collapsed (C23: the columns no row keeps
  apart merged, those past every row's last cell dropped with the rows'
  ``w:gridAfter`` and ``w:wAfter``);
* either way, a row's ``w:tblPrEx`` that only repeats the table's ``w:tblPr``
  goes (C23).

The recorded properties are not the live classes (CR-043 section 3.4): they go
back through their XML (:func:`~docx4j_py.model.revisions.props.convert`).
"""

from __future__ import annotations

import enum
from typing import Any

from docx4j_py.model.revisions import props, trees
from docx4j_py.model.revisions.trees import View

__all__ = [
    "accept",
    "cells_of",
    "change_text",
    "collapse_grid",
    "has_records",
    "in_change",
    "prune_row",
    "record_of",
    "reject",
    "reject_section",
    "remove_cell",
    "resolve_cell_merge",
    "rows_of",
]

_TABLE_SKIP = frozenset({"tblPrChange"})
_GRID_SKIP = frozenset({"tblGridChange"})
_EX_SKIP = frozenset({"tblPrExChange"})
_ROW_SKIP = frozenset({"ins", "del", "trPrChange"})
_CELL_SKIP = frozenset({"cellIns", "cellDel", "cellMerge", "tcPrChange"})


# ---------------------------------------------------------------------------
# rows and cells
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# what a table's records are
# ---------------------------------------------------------------------------


def in_change(tr: Any) -> bool:
    """A row in the change: one with a ``w:trPrChange`` or a ``w:tblPrExChange``, or a cell with a ``w:tcPrChange``."""
    tr_pr = tr.tr_pr
    if tr_pr is not None and tr_pr.tr_pr_change is not None:
        return True
    ex = tr.tbl_pr_ex
    if ex is not None and ex.tbl_pr_ex_change is not None:
        return True
    return any(tc.tc_pr is not None and tc.tc_pr.tc_pr_change is not None for tc, _owner in cells_of(tr))


def has_records(tbl: Any) -> bool:
    """True when a table has any property record, whether it changes anything or not."""
    if tbl.tbl_pr is not None and tbl.tbl_pr.tbl_pr_change is not None:
        return True
    if tbl.tbl_grid is not None and tbl.tbl_grid.tbl_grid_change is not None:
        return True
    return any(in_change(tr) for tr, _owner in rows_of(tbl))


def _table_differs(tbl: Any) -> bool:
    """True when the table's own properties or its grid are not as recorded."""
    tbl_pr = tbl.tbl_pr
    if (
        tbl_pr is not None
        and tbl_pr.tbl_pr_change is not None
        and props.canonical(tbl_pr, "tblPr", _TABLE_SKIP) != props.canonical(tbl_pr.tbl_pr_change.tbl_pr, "tblPr")
    ):
        return True
    grid = tbl.tbl_grid
    return (
        grid is not None
        and grid.tbl_grid_change is not None
        and props.canonical(grid, "tblGrid", _GRID_SKIP) != props.canonical(grid.tbl_grid_change.tbl_grid, "tblGrid")
    )


def _row_differs(tr: Any) -> bool:
    """True when a row in the change is not as its records have it.

    A row without a ``w:trPrChange`` had no properties (check 28).
    """
    tr_pr = tr.tr_pr
    recorded = tr_pr.tr_pr_change.tr_pr if tr_pr is not None and tr_pr.tr_pr_change is not None else None
    if props.canonical(tr_pr, "trPr", _ROW_SKIP) != props.canonical(recorded, "trPr"):
        return True
    ex = tr.tbl_pr_ex
    if (
        ex is not None
        and ex.tbl_pr_ex_change is not None
        and props.canonical(ex, "tblPrEx", _EX_SKIP) != props.canonical(ex.tbl_pr_ex_change.tbl_pr_ex, "tblPrEx")
    ):
        return True
    for tc, _owner in cells_of(tr):
        tc_pr = tc.tc_pr
        if (
            tc_pr is not None
            and tc_pr.tc_pr_change is not None
            and props.canonical(tc_pr, "tcPr", _CELL_SKIP) != props.canonical(tc_pr.tc_pr_change.tc_pr, "tcPr")
        ):
            return True
    return False


def record_of(tbl: Any) -> Any:
    """The record a table's change reports --- the table's, else the first changed row's or cell's.

    None when its records change nothing (checks 24 and 26's rule, U20) or are
    a grid's alone, which has no author: those the sweep drops.
    """
    if not has_records(tbl):
        return None
    changed = [tr for tr, _owner in rows_of(tbl) if in_change(tr)]
    if not _table_differs(tbl) and not any(_row_differs(tr) for tr in changed):
        return None
    if tbl.tbl_pr is not None and tbl.tbl_pr.tbl_pr_change is not None:
        return tbl.tbl_pr.tbl_pr_change
    for tr in changed:
        if tr.tr_pr is not None and tr.tr_pr.tr_pr_change is not None:
            return tr.tr_pr.tr_pr_change
        if tr.tbl_pr_ex is not None and tr.tbl_pr_ex.tbl_pr_ex_change is not None:
            return tr.tbl_pr_ex.tbl_pr_ex_change
        for tc, _owner in cells_of(tr):
            if tc.tc_pr is not None and tc.tc_pr.tc_pr_change is not None:
                return tc.tc_pr.tc_pr_change
    return None


def change_text(tbl: Any) -> str:
    """A table's change's text: every row's where its own properties or grid changed, else the changed rows' (checks 27, 28)."""
    from docx4j_py.model.revisions.texts import row_text

    whole = _table_differs(tbl)
    return "".join(
        row_text(tr, View.ACCEPTED) for tr, _owner in rows_of(tbl) if whole or (in_change(tr) and _row_differs(tr))
    )


# ---------------------------------------------------------------------------
# accepting and rejecting them
# ---------------------------------------------------------------------------


def accept(tbl: Any) -> None:
    """Accepting keeps the properties as they are and drops every record; a row's ``w:trPr`` left empty goes."""
    changed = [tr for tr, _owner in rows_of(tbl) if in_change(tr)]
    if tbl.tbl_pr is not None:
        tbl.tbl_pr.tbl_pr_change = None
    if tbl.tbl_grid is not None:
        tbl.tbl_grid.tbl_grid_change = None
    for tr, _owner in rows_of(tbl):
        if tr.tr_pr is not None:
            tr.tr_pr.tr_pr_change = None
        if tr.tbl_pr_ex is not None:
            tr.tbl_pr_ex.tbl_pr_ex_change = None
        prune_row(tr)
        for tc, _cell_owner in cells_of(tr):
            if tc.tc_pr is not None:
                tc.tc_pr.tc_pr_change = None
    for tr in changed:
        _drop_redundant_exception(tr, tbl)


def reject(tbl: Any) -> None:
    """Rejecting puts the recorded table back, wholesale.

    Its properties, its grid, and each row and cell in the change as recorded
    --- a row in the change without a ``w:trPrChange`` cleared, having had no
    properties (check 28) --- then a merged grid collapsed (check 27, C23).
    Rows outside the change are left.
    """
    from docx4j_py.wml import CTTblPrEx, TblGrid, TblPr, TcPr, TrPr

    changed = [tr for tr, _owner in rows_of(tbl) if in_change(tr)]
    tbl_pr = tbl.tbl_pr
    if tbl_pr is not None and tbl_pr.tbl_pr_change is not None:
        restored = props.convert(tbl_pr.tbl_pr_change.tbl_pr, "tblPr", TblPr) or TblPr()
        _set(tbl, "tbl_pr", restored)
    grid = tbl.tbl_grid
    grid_changed = grid is not None and grid.tbl_grid_change is not None
    if grid_changed:
        restored = props.convert(grid.tbl_grid_change.tbl_grid, "tblGrid", TblGrid) or TblGrid()
        _set(tbl, "tbl_grid", restored)
    for tr in changed:
        ex = tr.tbl_pr_ex
        if ex is not None and ex.tbl_pr_ex_change is not None:
            _set(tr, "tbl_pr_ex", props.convert(ex.tbl_pr_ex_change.tbl_pr_ex, "tblPrEx", CTTblPrEx))
        tr_pr = tr.tr_pr
        if tr_pr is not None:
            recorded = tr_pr.tr_pr_change.tr_pr if tr_pr.tr_pr_change is not None else None
            restored = (props.convert(recorded, "trPr", TrPr) if recorded is not None else None) or TrPr()
            restored.ins = tr_pr.ins
            restored.del_value = tr_pr.del_value
            restored.tr_pr_change = None
            _relink(restored, "ins", "del_value")
            _set(tr, "tr_pr", restored)
        for tc, _owner in cells_of(tr):
            tc_pr = tc.tc_pr
            if tc_pr is None or tc_pr.tc_pr_change is None:
                continue
            restored = props.convert(tc_pr.tc_pr_change.tc_pr, "tcPr", TcPr) or TcPr()
            restored.cell_ins_or_cell_del_or_cell_merge = tc_pr.cell_ins_or_cell_del_or_cell_merge
            restored.tc_pr_change = None
            _relink(restored, "cell_ins_or_cell_del_or_cell_merge")
            _set(tc, "tc_pr", None if trees.is_empty(restored) else restored)
    if grid_changed:
        collapse_grid(tbl)
    for tr, _owner in rows_of(tbl):
        prune_row(tr)
    for tr in changed:
        _drop_redundant_exception(tr, tbl)


def prune_row(tr: Any) -> None:
    """A row's ``w:trPr`` with nothing left in it goes, as Word writes none (check 18)."""
    if tr.tr_pr is not None and trees.is_empty(tr.tr_pr):
        tr.tr_pr = None


def _drop_redundant_exception(tr: Any, tbl: Any) -> None:
    """A row's ``w:tblPrEx`` that says only what the table's ``w:tblPr`` says goes.

    Word wrote none for C23's indented row once the change was accepted or
    rejected.
    """
    ex = tr.tbl_pr_ex
    if ex is None or ex.tbl_pr_ex_change is not None:
        return
    if trees.is_empty(ex):
        tr.tbl_pr_ex = None
        return
    ex_xml = props.element_of(ex, "tblPrEx")
    tbl_xml = props.element_of(tbl.tbl_pr, "tblPr")
    for child in ex_xml:
        if not isinstance(child.tag, str):
            continue
        local = child.tag.rpartition("}")[2]
        same = None
        for theirs in tbl_xml if tbl_xml is not None else ():
            if isinstance(theirs.tag, str) and theirs.tag.rpartition("}")[2] == local:
                same = theirs
        if same is None:
            return
        a: list[str] = []
        b: list[str] = []
        props.inner(child, None, a)
        props.inner(same, None, b)
        if a != b:
            return
    tr.tbl_pr_ex = None


def _set(owner: Any, field: str, value: Any) -> None:
    """A single-valued field set, and the value's parent pointer with it (assignment sets none, CLAUDE.md)."""
    setattr(owner, field, value)
    if value is not None:
        value.parent = owner


def _relink(owner: Any, *fields: str) -> None:
    """The parents of members moved over from the properties a restored one replaces."""
    for field in fields:
        value = getattr(owner, field, None)
        if value is not None:
            value.parent = owner


# ---------------------------------------------------------------------------
# a merged grid collapsed
# ---------------------------------------------------------------------------


def _row_member(tr: Any, local: str) -> Any:
    tr_pr = tr.tr_pr
    if tr_pr is None:
        return None
    for member in tr_pr.content:
        if trees.name(member) == local:
            return member
    return None


def _grid_before(tr: Any) -> int:
    member = _row_member(tr, "gridBefore")
    value = getattr(member, "val", None) if member is not None else None
    return int(value) if value is not None else 0


def _set_row_member(tr: Any, count: str, width: str, value: int) -> None:
    member = _row_member(tr, count)
    if value <= 0:
        if tr.tr_pr is not None:
            for gone in (member, _row_member(tr, width)):
                if gone is not None:
                    index = trees.index_of(tr.tr_pr.content, gone)
                    if index >= 0:
                        del tr.tr_pr.content[index]
        return
    if member is None:
        return  # a row that skipped nothing skips nothing still
    member.val = value


def _span(tc: Any) -> int:
    tc_pr = tc.tc_pr
    grid_span = tc_pr.grid_span if tc_pr is not None else None
    value = grid_span.val if grid_span is not None else None
    if value is None:
        return 1
    return max(int(value), 1)


def collapse_grid(tbl: Any) -> None:
    """Grid columns that no row keeps apart become one.

    The edges a row's cells (or its ``w:gridBefore``) fall on are kept, the
    others merged, and the columns past every row's last cell dropped (Word's
    reject of C23 left two columns where the merged grid had five, the first
    row's ``w:gridAfter`` gone); each cell's ``w:gridSpan`` and each row's
    skipped columns are counted again over the columns kept.
    """
    from docx4j_py.wml import el

    grid = tbl.tbl_grid
    if grid is None:
        return
    cols = grid.grid_col
    if len(cols) < 2:
        return
    rows = [tr for tr, _owner in rows_of(tbl)]
    edges = {0}
    for tr in rows:
        at = _grid_before(tr)
        edges.add(at)
        for tc, _owner in cells_of(tr):
            at += _span(tc)
            edges.add(at)
    last = min(max(edges), len(cols))
    kept = sorted(edge for edge in edges if edge <= last)
    if len(kept) - 1 == len(cols):
        return

    def index(edge: int) -> int:
        return kept.index(edge) if edge in kept else -1

    merged = []
    for i in range(1, len(kept)):
        width = 0
        for c in range(kept[i - 1], min(kept[i], len(cols))):
            if cols[c].w is not None:
                width += int(cols[c].w)
        merged.append(el.gridCol(w=width))
    cols.clear()
    cols.extend(merged)
    for tr in rows:
        at = _grid_before(tr)
        _set_row_member(tr, "gridBefore", "wBefore", index(at))
        for tc, _owner in cells_of(tr):
            following = at + _span(tc)
            count = index(min(following, last)) - index(min(at, last))
            tc_pr = tc.tc_pr
            if tc_pr is not None:
                if count > 1:
                    grid_span = tc_pr.grid_span if tc_pr.grid_span is not None else el.gridSpan()
                    grid_span.val = count
                    _set(tc_pr, "grid_span", grid_span)
                else:
                    tc_pr.grid_span = None
            at = following
        _set_row_member(tr, "gridAfter", "wAfter", len(kept) - 1 - index(min(at, last)))


# ---------------------------------------------------------------------------
# cells inserted and deleted, cells merged
# ---------------------------------------------------------------------------


def remove_cell(tc: Any, tr: Any, removals: Any, story: Any) -> None:
    """A cell taken out of its row, the cell before it (else after it) widened over its grid columns.

    As Word did rejecting an inserted cell and accepting a deleted one (check
    27: ``w:gridSpan`` added, and the widths added when both are ``dxa``).
    """
    from docx4j_py.wml import el

    cells = cells_of(tr)
    at = next((i for i, (cell, _owner) in enumerate(cells) if cell is tc), -1)
    if at < 0:
        return
    owner = cells[at][1]
    if at > 0:
        neighbour = cells[at - 1][0]
    elif at + 1 < len(cells):
        neighbour = cells[at + 1][0]
    else:
        neighbour = None
    i = trees.index_of(owner, tc)
    if i < 0:
        return
    removals.discard(story, tc)
    del owner[i]
    if neighbour is None:
        return
    kept = neighbour.tc_pr
    if kept is None:
        kept = el.tcPr()
        _set(neighbour, "tc_pr", kept)
    _set(kept, "grid_span", el.gridSpan(val=_span(neighbour) + _span(tc)))
    mine = tc.tc_pr.tc_w if tc.tc_pr is not None else None
    theirs = kept.tc_w
    if (
        mine is not None
        and theirs is not None
        and _value(mine.type_value) == "dxa"
        and _value(theirs.type_value) == "dxa"
        and mine.w is not None
        and theirs.w is not None
    ):
        theirs.w = int(theirs.w) + int(mine.w)


def cell_marker(tc_pr: Any, local: str) -> Any:
    """A cell's ``w:cellIns``, ``w:cellDel`` or ``w:cellMerge`` (one choice in the schema), or None."""
    if tc_pr is None:
        return None
    marker = tc_pr.cell_ins_or_cell_del_or_cell_merge
    return marker if marker is not None and trees.name(marker) == local else None


def resolve_cell_merge(tc_pr: Any, accept: bool) -> None:
    """A ``w:cellMerge`` resolved (never listed: Office JS cannot list one).

    Rejected, the merge it recorded comes back (docx4j's reading, U8): its
    ``w:vMergeOrig`` ``rest`` a ``w:vMerge w:val="restart"``, ``cont`` a plain
    ``w:vMerge``, anything else none.
    """
    from docx4j_py.wml import el

    change = cell_marker(tc_pr, "cellMerge")
    if change is None:
        return
    tc_pr.cell_ins_or_cell_del_or_cell_merge = None
    if accept:
        return
    orig = _value(change.v_merge_orig)
    if orig == "rest":
        _set(tc_pr, "v_merge", el.vMerge(val="restart"))
    elif orig == "cont":
        _set(tc_pr, "v_merge", el.vMerge())
    else:
        tc_pr.v_merge = None


# ---------------------------------------------------------------------------
# sections
# ---------------------------------------------------------------------------

#: What a section's record never holds: the header and footer references (U13),
#: the record itself, the rsids.
_SECTION_KEEP = frozenset({"header_reference_or_footer_reference", "sect_pr_change"})


def reject_section(sect_pr: Any) -> None:
    """A section's properties put back as its ``w:sectPrChange`` recorded them.

    Word records only what changed, and rejecting laid those over the section,
    the rest left as it stood (check 18; a paragraph's section, C07); a
    recorded ``w:docGrid`` of pitch 0 stands for none, and Word wrote none; the
    header and footer references are not in the record and stay (U13).
    """
    import dataclasses

    from docx4j_py.wml import SectPr

    change = sect_pr.sect_pr_change
    recorded = change.sect_pr if change is not None else None
    sect_pr.sect_pr_change = None
    if recorded is None:
        return
    restored = props.convert(recorded, "sectPr", SectPr)
    if restored is None:
        return
    for field in dataclasses.fields(SectPr):
        name = field.name
        if name in _SECTION_KEEP or name.startswith("rsid"):
            continue
        value = getattr(restored, name, None)
        if value is None or isinstance(value, list):
            continue
        setattr(restored, name, None)
        _set(sect_pr, name, value)
    grid = sect_pr.doc_grid
    if (
        grid is not None
        and (grid.line_pitch is None or int(grid.line_pitch) == 0)
        and grid.type_value is None
        and grid.char_space is None
    ):
        sect_pr.doc_grid = None


def _value(found: Any) -> Any:
    """An attribute's value as it reads: an enum's value, else the value itself."""
    return found.value if isinstance(found, enum.Enum) else found
