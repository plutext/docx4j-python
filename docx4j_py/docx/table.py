"""``Table``, ``_Row``, ``_Column``, ``_Cell`` and their collections.

CR-003 section 3.6, Phase J. python-docx's table model is a **layout grid**:
``table.cell(r, c)`` is a position in the grid, and a horizontally or
vertically merged cell appears at every position it covers, the same ``_Cell``
object each time. The content API's :class:`~docx4j_py.model.content.Table`
indexes the ``w:tc`` elements a row actually holds, which is the other useful
answer and not the same one; this module implements python-docx's grid walk
(``w:gridSpan`` and ``w:vMerge``, and a row that starts late or ends early) and
delegates every *mutation* to the content API, so a facade edit records a
:class:`~docx4j_py.model.content.reports.ChangeReport` like any other.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any, overload

from docx4j_py.docx._base import BlockItemContainer, Story, recording
from docx4j_py.docx.enums import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_STYLE_TYPE,
    WD_TABLE_ALIGNMENT,
    WD_TABLE_DIRECTION,
)
from docx4j_py.docx.shared import Inches, Length, Twips, twips_of
from docx4j_py.model.content.table import Table as ContentTable
from docx4j_py.wml import Tbl, Tc, Tr, el

if TYPE_CHECKING:  # pragma: no cover - annotations only
    from docx4j_py.docx.styles import BaseStyle

__all__ = ["Table", "_Cell", "_Column", "_Columns", "_Row", "_Rows"]


def _rows_of(tbl: Tbl) -> list[Tr]:
    """The ``w:tr`` children of a table, in order.

    Direct children only, as python-docx's ``tr_lst`` is: a row inside a
    row-level ``w:sdt`` is not one of these, which is where the content API's
    ``Table.rows`` (CR-003 section 14) and python-docx part company.
    """
    return [item for item in getattr(tbl, "content", None) or () if isinstance(item, Tr)]


def _cells_of(tr: Tr) -> list[Tc]:
    """The ``w:tc`` children of a row, in order."""
    return [item for item in getattr(tr, "content", None) or () if isinstance(item, Tc)]


def _grid_span(tc: Tc) -> int:
    """``w:gridSpan/@val``, or 1."""
    tc_pr = getattr(tc, "tc_pr", None)
    value = getattr(getattr(tc_pr, "grid_span", None), "val", None) if tc_pr else None
    return max(1, int(value)) if value is not None else 1


def _v_merge(tc: Tc) -> str | None:
    """``w:vMerge/@val``, defaulted to ``"continue"`` when the element is present."""
    tc_pr = getattr(tc, "tc_pr", None)
    merge = getattr(tc_pr, "v_merge", None) if tc_pr else None
    if merge is None:
        return None
    value = getattr(merge, "val", None)
    value = getattr(value, "value", value)
    return "continue" if value is None else str(value)


def _grid_before(tr: Tr) -> int:
    """``w:trPr/w:gridBefore/@val``: layout columns the row skips at the start."""
    return _tr_pr_count(tr, "grid_before")


def _grid_after(tr: Tr) -> int:
    """``w:trPr/w:gridAfter/@val``: layout columns the row skips at the end."""
    return _tr_pr_count(tr, "grid_after")


def _tr_pr_count(tr: Tr, field: str) -> int:
    found = _tr_pr_child(tr, "gridBefore" if field == "grid_before" else "gridAfter")
    value = getattr(found, "val", None) if found is not None else None
    return int(value) if value is not None else 0


def _tr_pr_items(tr: Tr) -> list[Any]:
    """The ``w:trPr`` choice list, which the generated model keeps in ``content``."""
    tr_pr = getattr(tr, "tr_pr", None)
    return list(getattr(tr_pr, "content", None) or ()) if tr_pr is not None else []


def _tr_pr_child(tr: Tr, name: str) -> Any | None:
    """The one ``w:trPr`` child of an element name, or None."""
    from docx4j_py.traversal import element_name

    wanted = f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{name}"
    for item in _tr_pr_items(tr):
        if element_name(item) == wanted:
            return item
    return None


def _set_tr_pr_child(tr: Tr, name: str, item: Any | None) -> None:
    """Replace, add or remove the one ``w:trPr`` child of an element name."""
    from docx4j_py.child import ChildList
    from docx4j_py.traversal import element_name

    tr_pr = getattr(tr, "tr_pr", None)
    if tr_pr is None:
        if item is None:
            return
        tr_pr = el.trPr()
        tr.tr_pr = tr_pr
        tr_pr.parent = tr
    items = tr_pr.content
    if not isinstance(items, ChildList):
        items = ChildList(list(items or ()), owner=tr_pr)
        tr_pr.content = items
    wanted = f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{name}"
    for index, existing in enumerate(list(items)):
        if element_name(existing) == wanted:
            if item is None:
                del items[index]
            else:
                items[index] = item
            return
    if item is not None:
        items.append(item)


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------


class Table:
    """A ``w:tbl``: python-docx's ``Table``."""

    __slots__ = ("_story", "_tbl")

    def __init__(self, tbl: Tbl, story: Story) -> None:
        """Wrap the ``w:tbl``."""
        self._tbl = tbl
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:tbl`` are equal."""
        return isinstance(other, Table) and other._tbl is self._tbl

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._tbl))

    def __repr__(self) -> str:
        """``<docx4j_py.docx Table 3x2>``."""
        return f"<docx4j_py.docx Table {len(_rows_of(self._tbl))}x{self._column_count}>"

    # -- what is underneath ------------------------------------------------

    @property
    def element(self) -> Tbl:
        """The typed ``w:tbl``."""
        return self._tbl

    @property
    def _element(self) -> Tbl:
        """python-docx's private name for :attr:`element`."""
        return self._tbl

    @property
    def part(self) -> Any:
        """The part this table is in."""
        return self._story.part

    @property
    def table(self) -> Table:
        """This table; the terminus of the ``parent.table`` chain."""
        return self

    @property
    def view(self) -> ContentTable:
        """The content API's :class:`~docx4j_py.model.content.Table`. An extension."""
        body = self._story.body
        return ContentTable(self._tbl, body._container_of(self._tbl), body)

    # -- the grid ----------------------------------------------------------

    @property
    def _grid_cols(self) -> list[Any]:
        grid = getattr(self._tbl, "tbl_grid", None)
        return list(getattr(grid, "grid_col", None) or ()) if grid is not None else []

    @property
    def _column_count(self) -> int:
        """The number of layout-grid columns, as python-docx computes it."""
        columns = len(self._grid_cols)
        if columns:
            return columns
        return max(
            (_grid_before(tr) + sum(_grid_span(tc) for tc in _cells_of(tr)) + _grid_after(tr))
            for tr in _rows_of(self._tbl)
        ) if _rows_of(self._tbl) else 0

    @property
    def _cells(self) -> list[_Cell]:
        """One ``_Cell`` per layout-grid position; a merged cell repeats.

        python-docx's ``Table._cells``, with its ``iter_tcs`` inlined: a
        ``w:vMerge="continue"`` position answers with the cell one row up, and
        a ``w:gridSpan`` position with the cell to its left.
        """
        col_count = self._column_count
        cells: list[_Cell] = []
        for tr in _rows_of(self._tbl):
            for tc in _cells_of(tr):
                for span_index in range(_grid_span(tc)):
                    if _v_merge(tc) == "continue" and len(cells) >= col_count:
                        cells.append(cells[-col_count])
                    elif span_index > 0 and cells:
                        cells.append(cells[-1])
                    else:
                        cells.append(_Cell(tc, self))
        return cells

    def cell(self, row_idx: int, col_idx: int) -> _Cell:
        """The cell at the layout-grid position (`row_idx`, `col_idx`)."""
        cells = self._cells
        index = col_idx + row_idx * self._column_count
        if not 0 <= index < len(cells) or not 0 <= col_idx < self._column_count:
            raise IndexError(f"no cell at ({row_idx}, {col_idx}) in this table")
        return cells[index]

    def row_cells(self, row_idx: int) -> list[_Cell]:
        """The cells of the row at `row_idx`. Deprecated in python-docx."""
        count = self._column_count
        start = row_idx * count
        return self._cells[start : start + count]

    def column_cells(self, column_idx: int) -> list[_Cell]:
        """The cells of the column at `column_idx`."""
        cells = self._cells
        count = self._column_count
        return [cells[index] for index in range(column_idx, len(cells), count)]

    @property
    def rows(self) -> _Rows:
        """The table's rows."""
        return _Rows(self._tbl, self)

    @property
    def columns(self) -> _Columns:
        """The table's layout-grid columns."""
        return _Columns(self._tbl, self)

    # -- growing -----------------------------------------------------------

    def add_row(self) -> _Row:
        """A new row at the bottom, its cells sized from the grid."""
        with recording(self._story.body, "docx.add_row"):
            self.view.add_rows(1)
            return _Row(_rows_of(self._tbl)[-1], self)

    def add_column(self, width: int | Length) -> _Column:
        """A new rightmost column of `width`, with a cell in every row."""
        with recording(self._story.body, "docx.add_column"):
            twips = twips_of(width) or 0
            grid = getattr(self._tbl, "tbl_grid", None)
            if grid is None:
                grid = el.tblGrid()
                self._tbl.tbl_grid = grid
                grid.parent = self._tbl
            column = el.gridCol(w=twips)
            grid.grid_col.append(column)
            for tr in _rows_of(self._tbl):
                tc = el.tc()
                tc_pr = el.tcPr()
                tc_pr.tc_w = el.tcW(w=twips, type_value="dxa")
                tc.tc_pr = tc_pr
                tc.content.append(el.p())
                tr.content.append(tc)
            from docx4j_py.child import link_parents

            link_parents(self._tbl)
            return _Column(column, self)

    def _set_column_widths(self, twips: list[int]) -> None:
        """Write `twips` on ``w:tblGrid`` and on every row's ``w:tcW``.

        What python-docx's ``CT_Tbl.new_tbl`` does when it builds a table of a
        given width, and what ``_Cell.add_table`` needs. Not a content API
        member: CR-003 section 19.6 asks for a ``Table.column_widths`` setter
        there, and Phase J did **not** take it (section 20), because
        python-docx's own ``column.width`` and ``cell.width`` each write one
        half and neither writes both.
        """
        for column, width in zip(self._grid_cols, twips):
            column.w = width
        for tr in _rows_of(self._tbl):
            index = 0
            for tc in _cells_of(tr):
                span = _grid_span(tc)
                total = sum(twips[index : index + span]) if index < len(twips) else 0
                tc_pr = getattr(tc, "tc_pr", None)
                if tc_pr is None:
                    tc_pr = el.tcPr()
                    tc.tc_pr = tc_pr
                    tc_pr.parent = tc
                tc_pr.tc_w = el.tcW(w=total, type_value="dxa")
                tc_pr.tc_w.parent = tc_pr
                index += span

    # -- properties --------------------------------------------------------

    @property
    def _tbl_pr(self) -> Any:
        tbl_pr = getattr(self._tbl, "tbl_pr", None)
        if tbl_pr is None:
            tbl_pr = el.tblPr()
            self._tbl.tbl_pr = tbl_pr
            tbl_pr.parent = self._tbl
        return tbl_pr

    @property
    def alignment(self) -> WD_TABLE_ALIGNMENT | None:
        """``w:tblPr/w:jc``, as a ``WD_TABLE_ALIGNMENT`` member or None."""
        tbl_pr = getattr(self._tbl, "tbl_pr", None)
        value = getattr(getattr(tbl_pr, "jc", None), "val", None) if tbl_pr else None
        value = getattr(value, "value", value)
        if value is None:
            return None
        try:
            return WD_TABLE_ALIGNMENT.from_xml(str(value))
        except ValueError:
            return None

    @alignment.setter
    def alignment(self, value: WD_TABLE_ALIGNMENT | None) -> None:
        with recording(self._story.body, "docx.table.alignment"):
            tbl_pr = self._tbl_pr
            if value is None:
                tbl_pr.jc = None
                return
            tbl_pr.jc = el.jc(val=WD_TABLE_ALIGNMENT(value).xml_value)
            tbl_pr.jc.parent = tbl_pr

    @property
    def autofit(self) -> bool:
        """True when ``w:tblLayout`` is not ``fixed``, as python-docx reads it."""
        tbl_pr = getattr(self._tbl, "tbl_pr", None)
        layout = getattr(tbl_pr, "tbl_layout", None) if tbl_pr else None
        value = getattr(layout, "type_value", None)
        value = getattr(value, "value", value)
        return value is None or str(value) != "fixed"

    @autofit.setter
    def autofit(self, value: bool) -> None:
        with recording(self._story.body, "docx.table.autofit"):
            tbl_pr = self._tbl_pr
            layout = tbl_pr.tbl_layout
            if layout is None:
                layout = el.tblLayout()
                tbl_pr.tbl_layout = layout
                layout.parent = tbl_pr
            layout.type_value = "autofit" if value else "fixed"

    @property
    def table_direction(self) -> WD_TABLE_DIRECTION | None:
        """``w:bidiVisual``: which side the first column is on."""
        tbl_pr = getattr(self._tbl, "tbl_pr", None)
        toggle = getattr(tbl_pr, "bidi_visual", None) if tbl_pr else None
        if toggle is None:
            return None
        value = getattr(toggle, "val", None)
        on = True if value is None else bool(value)
        return WD_TABLE_DIRECTION.RTL if on else WD_TABLE_DIRECTION.LTR

    @table_direction.setter
    def table_direction(self, value: WD_TABLE_DIRECTION | None) -> None:
        with recording(self._story.body, "docx.table.table_direction"):
            tbl_pr = self._tbl_pr
            if value is None:
                tbl_pr.bidi_visual = None
                return
            rtl = int(value) == int(WD_TABLE_DIRECTION.RTL)
            tbl_pr.bidi_visual = el.bidiVisual() if rtl else el.bidiVisual(val=False)
            tbl_pr.bidi_visual.parent = tbl_pr

    @property
    def style(self) -> BaseStyle | None:
        """The table's style object, or the document's default table style."""
        from docx4j_py.docx.styles import style_for

        tbl_pr = getattr(self._tbl, "tbl_pr", None)
        style_id = getattr(getattr(tbl_pr, "tbl_style", None), "val", None) if tbl_pr else None
        return style_for(self._story.package, style_id, WD_STYLE_TYPE.TABLE)

    @style.setter
    def style(self, style_or_name: str | BaseStyle | None) -> None:
        from docx4j_py.docx.styles import style_id_for

        style_id = style_id_for(
            self._story.package, style_or_name, WD_STYLE_TYPE.TABLE, define=True
        )
        with recording(self._story.body, "docx.table.style"):
            if style_id is None:
                tbl_pr = getattr(self._tbl, "tbl_pr", None)
                if tbl_pr is not None:
                    tbl_pr.tbl_style = None
                return
            self.view.style_id = style_id


# ---------------------------------------------------------------------------
# rows, columns and cells
# ---------------------------------------------------------------------------


class _Row:
    """A ``w:tr``: python-docx's ``_Row``."""

    __slots__ = ("_table", "_tr")

    def __init__(self, tr: Tr, table: Table) -> None:
        """Wrap the ``w:tr``."""
        self._tr = tr
        self._table = table

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:tr`` are equal."""
        return isinstance(other, _Row) and other._tr is self._tr

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._tr))

    def __repr__(self) -> str:
        """``<docx4j_py.docx _Row, 3 cells>``."""
        return f"<docx4j_py.docx _Row, {len(self.cells)} cells>"

    @property
    def element(self) -> Tr:
        """The typed ``w:tr``."""
        return self._tr

    @property
    def _element(self) -> Tr:
        """python-docx's private name for :attr:`element`."""
        return self._tr

    @property
    def table(self) -> Table:
        """The table this row is in."""
        return self._table

    @property
    def part(self) -> Any:
        """The part this row is in."""
        return self._table.part

    @property
    def cells(self) -> tuple[_Cell, ...]:
        """One cell per populated layout-grid position, a merged cell repeated."""
        out: list[_Cell] = []
        for tc in _cells_of(self._tr):
            source = tc
            for _ in range(64):
                if _v_merge(source) != "continue":
                    break
                above = self._tc_above(source)
                if above is None:
                    break
                source = above
            cell = _Cell(source, self._table)
            out.extend([cell] * _grid_span(tc))
        return tuple(out)

    def _tc_above(self, tc: Tc) -> Tc | None:
        """The ``w:tc`` at the same grid offset in the row before this one."""
        rows = _rows_of(self._table.element)
        try:
            index = next(i for i, row in enumerate(rows) if row is self._tr)
        except StopIteration:
            return None
        if index == 0:
            return None
        offset = 0
        for candidate in _cells_of(self._tr):
            if candidate is tc:
                break
            offset += _grid_span(candidate)
        above = 0
        for candidate in _cells_of(rows[index - 1]):
            if above == offset:
                return candidate
            above += _grid_span(candidate)
        return None

    @property
    def grid_cols_before(self) -> int:
        """Unpopulated layout columns before the first cell (``w:gridBefore``)."""
        return _grid_before(self._tr)

    @property
    def grid_cols_after(self) -> int:
        """Unpopulated layout columns after the last cell (``w:gridAfter``)."""
        return _grid_after(self._tr)

    @property
    def height(self) -> Length | None:
        """``w:trPr/w:trHeight/@val``, or None."""
        found = _tr_pr_child(self._tr, "trHeight")
        value = getattr(found, "val", None) if found is not None else None
        return None if value is None else Twips(int(value))

    @height.setter
    def height(self, value: int | Length | None) -> None:
        with recording(self._table._story.body, "docx.row.height"):
            found = _tr_pr_child(self._tr, "trHeight")
            if value is None:
                if found is None:
                    return
                rule = getattr(found, "h_rule", None)
                if rule is None:
                    _set_tr_pr_child(self._tr, "trHeight", None)
                else:
                    found.val = None
                return
            item = found if found is not None else el.trHeight()
            item.val = twips_of(value)
            _set_tr_pr_child(self._tr, "trHeight", item)

    @property
    def height_rule(self) -> WD_ROW_HEIGHT_RULE | None:
        """``w:trHeight/@hRule``, as a ``WD_ROW_HEIGHT_RULE`` member or None."""
        found = _tr_pr_child(self._tr, "trHeight")
        value = getattr(found, "h_rule", None) if found is not None else None
        value = getattr(value, "value", value)
        if value is None:
            return None
        try:
            return WD_ROW_HEIGHT_RULE.from_xml(str(value))
        except ValueError:
            return None

    @height_rule.setter
    def height_rule(self, value: WD_ROW_HEIGHT_RULE | None) -> None:
        with recording(self._table._story.body, "docx.row.height_rule"):
            found = _tr_pr_child(self._tr, "trHeight")
            if value is None:
                if found is None:
                    return
                if found.val is None:
                    _set_tr_pr_child(self._tr, "trHeight", None)
                else:
                    found.h_rule = None
                return
            item = found if found is not None else el.trHeight()
            item.h_rule = WD_ROW_HEIGHT_RULE(value).xml_value
            _set_tr_pr_child(self._tr, "trHeight", item)


class _Rows:
    """``table.rows``: a sequence of :class:`_Row`."""

    __slots__ = ("_table", "_tbl")

    def __init__(self, tbl: Tbl, table: Table) -> None:
        """Wrap the table whose rows these are."""
        self._tbl = tbl
        self._table = table

    def __repr__(self) -> str:
        """``<docx4j_py.docx _Rows, 3 rows>``."""
        return f"<docx4j_py.docx _Rows, {len(self)} rows>"

    @property
    def table(self) -> Table:
        """The table these rows are in."""
        return self._table

    def __len__(self) -> int:
        """How many ``w:tr`` the table holds."""
        return len(_rows_of(self._tbl))

    def __iter__(self) -> Iterator[_Row]:
        """Each row, in document order."""
        return (_Row(tr, self._table) for tr in _rows_of(self._tbl))

    @overload
    def __getitem__(self, idx: int) -> _Row: ...

    @overload
    def __getitem__(self, idx: slice) -> list[_Row]: ...

    def __getitem__(self, idx: int | slice) -> _Row | list[_Row]:
        """Indexed and sliced access, as python-docx's is."""
        return list(self)[idx]


class _Column:
    """A ``w:gridCol``: python-docx's ``_Column``."""

    __slots__ = ("_grid_col", "_table")

    def __init__(self, grid_col: Any, table: Table) -> None:
        """Wrap the ``w:gridCol``."""
        self._grid_col = grid_col
        self._table = table

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:gridCol`` are equal."""
        return isinstance(other, _Column) and other._grid_col is self._grid_col

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._grid_col))

    def __repr__(self) -> str:
        """``<docx4j_py.docx _Column, 3119 twips>``."""
        return f"<docx4j_py.docx _Column, {self._grid_col.w} twips>"

    @property
    def table(self) -> Table:
        """The table this column is in."""
        return self._table

    @property
    def _index(self) -> int:
        for index, column in enumerate(self._table._grid_cols):
            if column is self._grid_col:
                return index
        return -1

    @property
    def cells(self) -> tuple[_Cell, ...]:
        """The cells at this column's layout-grid position, top to bottom."""
        return tuple(self._table.column_cells(self._index))

    @property
    def width(self) -> Length | None:
        """``w:gridCol/@w``, or None."""
        value = getattr(self._grid_col, "w", None)
        return None if value is None else Twips(int(value))

    @width.setter
    def width(self, value: int | Length | None) -> None:
        with recording(self._table._story.body, "docx.column.width"):
            self._grid_col.w = twips_of(value)


class _Columns:
    """``table.columns``: a sequence of :class:`_Column`."""

    __slots__ = ("_table", "_tbl")

    def __init__(self, tbl: Tbl, table: Table) -> None:
        """Wrap the table whose columns these are."""
        self._tbl = tbl
        self._table = table

    def __repr__(self) -> str:
        """``<docx4j_py.docx _Columns, 2 columns>``."""
        return f"<docx4j_py.docx _Columns, {len(self)} columns>"

    @property
    def table(self) -> Table:
        """The table these columns are in."""
        return self._table

    def __len__(self) -> int:
        """How many ``w:gridCol`` the grid holds."""
        return len(self._table._grid_cols)

    def __iter__(self) -> Iterator[_Column]:
        """Each column, left to right."""
        return (_Column(column, self._table) for column in self._table._grid_cols)

    def __getitem__(self, idx: int) -> _Column:
        """The column at `idx`."""
        try:
            return _Column(self._table._grid_cols[idx], self._table)
        except IndexError:
            raise IndexError(f"column index [{idx}] is out of range") from None


class _Cell(BlockItemContainer):
    """A ``w:tc``: python-docx's ``_Cell``, a block container of its own."""

    __slots__ = ("_table", "_tc")

    def __init__(self, tc: Tc, table: Table) -> None:
        """Wrap the ``w:tc``."""
        self._tc = tc
        self._table = table

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:tc`` are equal."""
        return isinstance(other, _Cell) and other._tc is self._tc

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._tc))

    def __repr__(self) -> str:
        """``<docx4j_py.docx _Cell 'Total'>``."""
        text = self.text.replace("\n", " ")
        preview = text[:24] + "…" if len(text) > 24 else text
        return f"<docx4j_py.docx _Cell {preview!r}>"

    @property
    def _story(self) -> Story:
        """The part this cell is in."""
        return self._table._story

    @property
    def _container(self) -> Tc:
        """The ``w:tc`` whose block children this cell holds."""
        return self._tc

    @property
    def element(self) -> Tc:
        """The typed ``w:tc``."""
        return self._tc

    @property
    def _element(self) -> Tc:
        """python-docx's private name for :attr:`element`."""
        return self._tc

    @property
    def part(self) -> Any:
        """The part this cell is in."""
        return self._table.part

    @property
    def table(self) -> Table:
        """The table this cell is in."""
        return self._table

    # -- text --------------------------------------------------------------

    @property
    def text(self) -> str:
        """The cell's text, a line per paragraph."""
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, text: str) -> None:
        with recording(self._story.body, "docx.cell.text"):
            body = self._content_body
            body.content.clear()
            body.insert_paragraph("")
            self.paragraphs[0].add_run(text)

    def add_table(self, rows: int, cols: int) -> Table:
        """A table in this cell, followed by the empty paragraph Word requires."""
        width = self.width if self.width is not None else Inches(1)
        with recording(self._story.body, "docx.cell.add_table"):
            table = super().add_table(rows, cols, width)
            self.add_paragraph()
            return table

    # -- properties --------------------------------------------------------

    @property
    def _tc_pr(self) -> Any:
        tc_pr = getattr(self._tc, "tc_pr", None)
        if tc_pr is None:
            tc_pr = el.tcPr()
            self._tc.tc_pr = tc_pr
            tc_pr.parent = self._tc
        return tc_pr

    @property
    def grid_span(self) -> int:
        """How many layout-grid columns this cell spans; 1 when it is not merged."""
        return _grid_span(self._tc)

    @property
    def width(self) -> Length | None:
        """``w:tcW/@w``, or None when no explicit width is set."""
        tc_pr = getattr(self._tc, "tc_pr", None)
        tc_w = getattr(tc_pr, "tc_w", None) if tc_pr else None
        value = getattr(tc_w, "w", None) if tc_w is not None else None
        return None if value is None else Twips(int(value))

    @width.setter
    def width(self, value: int | Length | None) -> None:
        with recording(self._story.body, "docx.cell.width"):
            if value is None:
                tc_pr = getattr(self._tc, "tc_pr", None)
                if tc_pr is not None:
                    tc_pr.tc_w = None
                return
            tc_pr = self._tc_pr
            tc_pr.tc_w = el.tcW(w=twips_of(value), type_value="dxa")
            tc_pr.tc_w.parent = tc_pr

    @property
    def vertical_alignment(self) -> WD_CELL_VERTICAL_ALIGNMENT | None:
        """``w:tcPr/w:vAlign``, as a ``WD_ALIGN_VERTICAL`` member or None."""
        tc_pr = getattr(self._tc, "tc_pr", None)
        value = getattr(getattr(tc_pr, "v_align", None), "val", None) if tc_pr else None
        value = getattr(value, "value", value)
        if value is None:
            return None
        try:
            return WD_CELL_VERTICAL_ALIGNMENT.from_xml(str(value))
        except ValueError:
            return None

    @vertical_alignment.setter
    def vertical_alignment(self, value: WD_CELL_VERTICAL_ALIGNMENT | None) -> None:
        with recording(self._story.body, "docx.cell.vertical_alignment"):
            if value is None:
                tc_pr = getattr(self._tc, "tc_pr", None)
                if tc_pr is not None:
                    tc_pr.v_align = None
                return
            tc_pr = self._tc_pr
            tc_pr.v_align = el.vAlign(val=WD_CELL_VERTICAL_ALIGNMENT(value).xml_value)
            tc_pr.v_align.parent = tc_pr

    def merge(self, other_cell: _Cell) -> _Cell:
        """The cell made by merging the rectangle this cell and `other_cell` corner.

        Not implemented: python-docx's ``CT_Tc.merge`` grows the span of the
        top-left ``w:tc``, moves the content of every cell in the rectangle
        into it and marks the rest ``w:vMerge="continue"`` with their
        ``w:gridSpan`` adjusted --- about 150 lines of grid surgery whose
        failure mode in Word is a repair prompt. Phase J left it out rather
        than half-implement it (CR-003 section 20).
        """
        raise NotImplementedError(
            "docx4j_py.docx does not implement _Cell.merge; build the merged cell "
            "with the typed tree (cell.element.tc_pr.grid_span / v_merge) or insert "
            "the table's XML through body.insert_xml (CR-003 section 20)"
        )
