"""``Table``, ``TableRow`` and ``TableCell``: Office JS's over ``w:tbl``.

CR-003 section 3.2, Phase C. Views, as :class:`~docx4j_py.model.content.Paragraph`
is: each holds its element and the live list that contains it, nothing is
cached, and every write mutates the tree in place and links the parents of what
it adds.

Two rules of CR-003 section 4 land here:

* **rows and cells descend into row- and cell-level content controls.** An
  OpenDoPE repeat wraps its ``w:tr`` in a ``w:sdt``, so
  :func:`~docx4j_py.model.content.text_model.rows_of` reports the rows Word
  shows and the list each one really lives in, while :meth:`Table.add_rows`
  puts new rows in the table's **own** content;
* **``header_row_count`` counts leading rows** carrying ``w:tblHeader``, and
  :meth:`~docx4j_py.model.content.body.Body.insert_table` sets no style unless
  one is asked for, so a new table is borderless until
  ``table.style_built_in = "TableGrid"``.

A cell's :attr:`TableCell.body` is ``Body.sub(tc, prefix)``, so the addresses of
CR-003 section 3.4 carry straight on into it: ``body/4/0/1/0`` is the first
block of the second cell of the first row of the fifth block.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.child import link_parents
from docx4j_py.model.content.addresses import ordinal_of
from docx4j_py.model.content.enums import BodyLocation, ParagraphLocation, TextLocation
from docx4j_py.model.content.errors import ContentError, InvalidTargetError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.styles import (
    built_in_of,
    define_built_in,
    id_of_built_in,
    style_id_of,
    style_name_of,
)
from docx4j_py.model.content.text_model import (
    W_TBL,
    W_TC,
    W_TR,
    block_children_of,
    cells_of,
    rows_of,
)
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, text_of
from docx4j_py.wml import P, Tbl, Tc, Tr, el, tbl, to_xml, tr

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "DEFAULT_TABLE_WIDTH",
    "TWIPS_PER_POINT",
    "Table",
    "TableCell",
    "TableRow",
    "cell_of",
    "insert_table_into",
    "writable_width",
]

#: Twips per point, the unit Office JS reports a width in.
TWIPS_PER_POINT = 20

#: A4 less 2.54 cm margins, in twips: what the ``tbl`` builder falls back to and
#: what a body with no ``w:sectPr`` gets.
DEFAULT_TABLE_WIDTH = 9026

#: How far up the parent chain :func:`cell_of` looks. A cell is at most a
#: ``w:sdt`` or two above a paragraph.
_MAX_DEPTH = 64

#: How much of the text a ``repr`` shows.
_PREVIEW = 30

#: ``w:trPr/w:tblHeader``, the element :attr:`TableRow.is_header` is over.
W_TBL_HEADER = f"{{{WML_NS}}}tblHeader"


def writable_width(body: Any) -> int:
    """The section's page width less its margins, in twips; A4 when unstated.

    The number :meth:`Body.insert_table` sizes a grid from and CR-003 section
    10.4 asks for: ``pgSz/@w`` minus ``pgMar/@left`` and ``@right``. It lives
    here because :func:`insert_table_into` is the caller that defines it;
    :mod:`docx4j_py.model.markdown.importer` imports it rather than keeping a
    copy (CR-003 section 13.5).
    """
    container = getattr(body, "container", body)
    sect_pr = getattr(container, "sect_pr", None)
    pg_sz = getattr(sect_pr, "pg_sz", None) if sect_pr is not None else None
    width = getattr(pg_sz, "w", None) if pg_sz is not None else None
    if width is None:
        return DEFAULT_TABLE_WIDTH
    pg_mar = getattr(sect_pr, "pg_mar", None)
    left = getattr(pg_mar, "left", None) if pg_mar is not None else None
    right = getattr(pg_mar, "right", None) if pg_mar is not None else None
    try:
        writable = int(width) - int(left if left is not None else 1440)
        writable -= int(right if right is not None else 1440)
    except (TypeError, ValueError):  # pragma: no cover - a malformed w:sectPr
        return DEFAULT_TABLE_WIDTH
    return writable if writable > 0 else DEFAULT_TABLE_WIDTH


def _index_of(items: list, element: Any) -> int:
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def _owner_of(container: list, fallback: Any) -> Any:
    """The typed object whose list this is, for the parent pointers."""
    owner = getattr(container, "owner", None)
    return owner if owner is not None else fallback


def _row_element(widths: list[int], values: list[str] | None = None) -> Tr:
    """A ``w:tr`` of cells of these widths, one paragraph each."""
    return tr(
        [values[i] if values is not None and i < len(values) else "" for i in range(len(widths))],
        widths=widths,
    )


def _delete_row(tracker: Any, row: Tr, body: Any) -> bool:
    """Mark a row deleted, **content and all**; True when it was taken back.

    CR-003 section 16.10: Word writes a deleted row as ``w:trPr/w:del`` plus
    every run in every cell in a ``w:del`` with ``w:delText`` and every cell
    paragraph's mark marked deleted. A row this author inserted is taken back,
    and then the caller removes it.
    """
    from docx4j_py.model.content.tracking import track_deleted_row

    return track_deleted_row(tracker, row, body)


def _mark_rows_inserted(tracker: Any, rows: list[Tr]) -> None:
    """``w:trPr/w:ins`` on each new row, and a ``w:ins`` on each paragraph in it.

    The one place rows are built is :func:`_row_element`, and the one place they
    are marked is here, so ``add_rows`` and ``insert_rows`` track alike
    (CR-003 section 14.7).
    """
    if tracker is None:
        return
    from docx4j_py.model.content.tracking import track_inserted_paragraph

    for row in rows:
        tracker.mark_row_inserted(row)
        for cell, _owner in cells_of(row):
            for block in getattr(cell, "content", None) or []:
                if isinstance(block, P):
                    track_inserted_paragraph(tracker, block)


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------


class Table:
    """A subset of Office JS ``Word.Table`` over a ``w:tbl``."""

    __slots__ = ("container", "element", "parent_body")

    def __init__(self, element: Tbl, container: list, parent_body: Body) -> None:
        """Build the view over a ``w:tbl``, the list holding it and its body."""
        #: The ``w:tbl``.
        self.element = element
        #: The live list holding the element.
        self.container = container
        #: The nearest :class:`~docx4j_py.model.content.Body`.
        self.parent_body = parent_body

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:tbl`` are equal (CR-003 section 3.1)."""
        return isinstance(other, Table) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity, so views can go in a set."""
        return hash(id(self.element))

    def __str__(self) -> str:
        """The table's text."""
        return self.text

    def __repr__(self) -> str:
        """``<Table body/4 3x2 'Region | Total'>``."""
        rows = self.rows
        preview = " | ".join(rows[0].values) if rows else ""
        if len(preview) > _PREVIEW:
            preview = preview[:_PREVIEW] + "…"
        where = self.ordinal
        columns = rows[0].cell_count if rows else 0
        return f"<Table{' ' + where if where else ''} {len(rows)}x{columns} {preview!r}>"

    @property
    def tbl(self) -> Tbl:
        """The ``w:tbl``; docx4j's name for :attr:`element`."""
        return self.element

    @property
    def change_tracker(self) -> Any:
        """The package's tracker while ``change_tracking_mode`` is on, else None."""
        return self.parent_body.change_tracker

    @property
    def name(self) -> str:
        """``"w:tbl"``, as :class:`~docx4j_py.model.content.body.Block` reports it."""
        return "w:tbl"

    @property
    def index(self) -> int:
        """This table's position in its container, or -1 if it is gone."""
        return _index_of(self.container, self.element)

    # -- addresses ---------------------------------------------------------

    @property
    def address(self) -> str:
        """The ordinal address (``"body/4"``); CR-003 section 3.4.

        A table has no ``w14:paraId``, so its address is always the ordinal,
        which is what :class:`~docx4j_py.model.content.body.Block` reported
        before this view existed.
        """
        return self.ordinal

    @property
    def ordinal(self) -> str:
        """The ordinal address, ``""`` when the table is no longer in its body."""
        return ordinal_of(self.parent_body, self.element) or ""

    # -- reading -----------------------------------------------------------

    @property
    def row_count(self) -> int:
        """The number of rows, row-level content controls descended into."""
        return len(rows_of(self.element))

    @property
    def rows(self) -> list[TableRow]:
        """Every row, in order (Office JS ``Table.rows``)."""
        return [TableRow(element, container, self) for element, container in rows_of(self.element)]

    @property
    def values(self) -> list[list[str]]:
        """The text of every cell, row by row (Office JS ``Table.values``)."""
        return [row.values for row in self.rows]

    @values.setter
    def values(self, values: list[list[str]]) -> None:
        with recording(self.parent_body, "set_values") as change:
            change.text(before=self.text)
            rows = self.rows
            for index, row in enumerate(values):
                if index < len(rows):
                    rows[index].values = row
            change.touched(self.address)
            change.text(after=self.text)

    @property
    def text(self) -> str:
        """The table's text, as docx4j's ``TextUtils`` reads it. Extension."""
        return text_of(self.element)

    @property
    def parent_table_cell(self) -> TableCell | None:
        """The cell this table is nested in, or None. Extension."""
        parent = getattr(self.element, "parent", None)
        return None if parent is None else cell_of(parent, self.parent_body)

    def cell(self, row_index: int, cell_index: int) -> TableCell:
        """The cell at a row and a column (Office JS ``getCell``, python-docx's name).

        Raises:
            ContentError: there is no such row or cell; the message says how
                many there are.
        """
        rows = self.rows
        if not 0 <= row_index < len(rows):
            raise ContentError(
                f"no row {row_index} in this table; it has {len(rows)}",
                code="table.no_row",
                hint=f"row_index is 0 to {len(rows) - 1}" if rows else "the table has no rows",
            )
        row = rows[row_index]
        cells = row.cells
        if not 0 <= cell_index < len(cells):
            raise ContentError(
                f"no cell {cell_index} in row {row_index}; it has {len(cells)}",
                code="table.no_cell",
                hint=f"cell_index is 0 to {len(cells) - 1}" if cells else "the row has no cells",
            )
        return cells[cell_index]

    def column_widths(self) -> list[float]:
        """The column widths in **points**, from ``w:tblGrid`` or the first row.

        Points, because that is the unit every other measurement in this API is
        in (:attr:`TableCell.width`, ``Paragraph.left_indent``); the TypeScript
        engine reports twips here and :meth:`column_widths_twips` is that.
        """
        return [width / TWIPS_PER_POINT for width in self.column_widths_twips()]

    def column_widths_twips(self) -> list[int]:
        """The column widths in twips, which is what a new row is built from."""
        grid = getattr(self.element.tbl_grid, "grid_col", None) if self.element.tbl_grid else None
        if grid:
            each = DEFAULT_TABLE_WIDTH // len(grid)
            return [int(column.w) if column.w is not None else each for column in grid]
        rows = rows_of(self.element)
        cells = cells_of(rows[0][0]) if rows else []
        count = max(1, len(cells))
        each = DEFAULT_TABLE_WIDTH // count
        out: list[int] = []
        for index in range(count):
            tc_pr = getattr(cells[index][0], "tc_pr", None) if index < len(cells) else None
            tc_w = getattr(tc_pr, "tc_w", None) if tc_pr is not None else None
            value = getattr(tc_w, "w", None) if tc_w is not None else None
            out.append(int(value) if value is not None else each)
        return out

    # -- style -------------------------------------------------------------

    @property
    def style_id(self) -> str:
        """``w:tblStyle``, docx4j's view of a style; ``""`` when there is none."""
        tbl_pr = self.element.tbl_pr
        style = getattr(tbl_pr, "tbl_style", None) if tbl_pr is not None else None
        value = getattr(style, "val", None) if style is not None else None
        value = getattr(value, "value", value)
        return str(value) if value else ""

    @style_id.setter
    def style_id(self, value: str) -> None:
        with recording(self.parent_body, "format") as change:
            change.text(before=f"style_id={self.style_id}")
            tbl_pr = self._tbl_pr()
            if not value:
                tbl_pr.tbl_style = None
            else:
                tbl_pr.tbl_style = el.tblStyle(val=value)
                tbl_pr.tbl_style.parent = tbl_pr
            change.touched(self.address)
            change.text(after=f"style_id={self.style_id}")

    @property
    def style(self) -> str:
        """The style's display name (``"Table Grid"``), as Office JS reports it.

        Read from the styles part's ``w:name`` when that part is already
        unmarshalled and derived from the id otherwise; **a read never
        unmarshals it** (CR-003 section 4), exactly as ``Paragraph.style``.
        Setting a built-in style the document does not define **adds** the
        definition, as on a paragraph (CR-003 section 14.9): a dangling
        ``w:tblStyle`` renders as Table Normal.
        """
        style_id = self.style_id
        return style_name_of(self.parent_body.package, style_id) if style_id else ""

    @style.setter
    def style(self, value: str) -> None:
        # the recording is opened here so that the definition a built-in style
        # needs is added inside it, and ``/word/styles.xml`` reaches the report
        with recording(self.parent_body, "format"):
            resolved = style_id_of(self.parent_body.package, value, define=True) if value else ""
            self.style_id = resolved

    @property
    def style_built_in(self) -> str:
        """The ``Word.Style`` value (``"TableGrid"``), or ``"Other"``."""
        return built_in_of(self.style_id)

    @style_built_in.setter
    def style_built_in(self, value: str) -> None:
        with recording(self.parent_body, "format"):
            self.style_id = define_built_in(self.parent_body.package, id_of_built_in(value))

    @property
    def header_row_count(self) -> int:
        """How many **leading** rows repeat at the top of a page (``w:tblHeader``)."""
        count = 0
        for row in self.rows:
            if not row.is_header:
                break
            count += 1
        return count

    @header_row_count.setter
    def header_row_count(self, count: int) -> None:
        with recording(self.parent_body, "format") as change:
            change.text(before=f"header_row_count={self.header_row_count}")
            for index, row in enumerate(self.rows):
                row.is_header = index < count
            change.touched(self.address)
            change.text(after=f"header_row_count={self.header_row_count}")

    # -- editing -----------------------------------------------------------

    def add_rows(
        self,
        row_count: int = 1,
        *,
        location: BodyLocation = "End",
        values: list[list[str]] | None = None,
    ) -> list[TableRow]:
        """Rows at the start or the end, with this table's columns and widths.

        New rows go into the table's **own** content, never into a row-level
        content control (CR-003 section 4).

        Args:
            row_count: how many.
            location: ``"Start"`` or ``"End"`` (the default).
            values: the text of the new rows' cells, row by row.

        Returns:
            The new rows' views, in order.
        """
        if row_count < 1:
            raise ContentError(
                f"add_rows needs at least one row, not {row_count}",
                code="table.no_rows",
                hint="pass row_count=1 or more",
            )
        with recording(self.parent_body, "add_rows") as change:
            widths = self.column_widths_twips()
            added = [
                _row_element(widths, values[i] if values is not None and i < len(values) else None)
                for i in range(row_count)
            ]
            container = self.element.content
            if location == "Start":
                existing = rows_of(self.element)
                first = next((e for e, c in existing if c is container), None)
                index = _index_of(container, first) if first is not None else 0
                index = max(index, 0)
            else:
                index = len(container)
            for offset, row in enumerate(added):
                container.insert(index + offset, row)
                link_parents(row)
                row.parent = self.element
            _mark_rows_inserted(self.change_tracker, added)
            change.touched(self.address)
            change.text(after="\n".join("\t".join(r) for r in (values or [])))
            return [TableRow(row, container, self) for row in added]

    def delete_rows(self, row_index: int, row_count: int = 1) -> None:
        """Remove `row_count` rows from `row_index` (Office JS ``deleteRows``).

        **While the package tracks changes** the rows stay in the tree and take
        a ``w:trPr/w:del`` instead, so :attr:`row_count` and :attr:`values`
        still report them until the change is accepted --- exactly as Word shows
        them, and the one place a tracked call's answer differs from an
        untracked one (CR-003 section 4).
        """
        with recording(self.parent_body, "delete_rows") as change:
            tracker = self.change_tracker
            rows = rows_of(self.element)[row_index : row_index + row_count]
            for element, container in reversed(rows):
                if tracker is not None and not _delete_row(tracker, element, self.parent_body):
                    continue
                index = _index_of(container, element)
                if index >= 0:
                    del container[index]
            change.touched(self.address)
            change.text(after="")

    def delete(self) -> None:
        """Remove the table from its container.

        **While the package tracks changes** the table stays and every row is
        marked deleted (``w:trPr/w:del``), which is what Word shows until the
        change is accepted (CR-003 section 14.7).
        """
        from docx4j_py.model.content.reports import moved_by_delete

        with recording(self.parent_body, "delete") as change:
            index = self.index
            if index < 0:
                return
            change.touched(self.address)
            change.text(before=self.text, after="")
            tracker = self.change_tracker
            if tracker is not None:
                for element, container in reversed(rows_of(self.element)):
                    if getattr(getattr(element, "tr_pr", None), "del_value", None) is not None:
                        continue
                    if _delete_row(tracker, element, self.parent_body):
                        at = _index_of(container, element)
                        if at >= 0:
                            del container[at]
                return
            change.shifted(moved_by_delete(self.parent_body, self.container, index))
            del self.container[index]

    # -- output ------------------------------------------------------------

    def to_markdown(self, **options: Any) -> str:
        """This table as a GFM pipe table (CR-003 section 3.5).

        The renderer is :func:`docx4j_py.model.markdown.table_markdown`, which
        already unwraps row- and cell-level content controls, pads a
        ``w:gridSpan`` and leaves a ``w:vMerge`` continuation empty.
        """
        from docx4j_py.model.markdown import table_markdown

        return table_markdown(self.element, self.parent_body, **options)

    def get_xml(self) -> str:
        """The table as XML, with docx4j's prefixes. Extension."""
        return to_xml(self.element)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a table."""
        rows = self.rows
        return {
            "kind": "table",
            "address": self.address,
            "ordinal": self.ordinal,
            "rows": len(rows),
            "cols": rows[0].cell_count if rows else 0,
            "header_row_count": self.header_row_count,
            "style": self.style,
            "style_id": self.style_id,
            "style_built_in": self.style_built_in,
            "values": self.values,
        }

    # -- helpers -----------------------------------------------------------

    def _tbl_pr(self) -> Any:
        """The ``w:tblPr``, created and linked when absent."""
        tbl_pr = self.element.tbl_pr
        if tbl_pr is None:
            tbl_pr = el.tblPr()
            self.element.tbl_pr = tbl_pr
            tbl_pr.parent = self.element
        return tbl_pr


# ---------------------------------------------------------------------------
# TableRow
# ---------------------------------------------------------------------------


class TableRow:
    """A subset of Office JS ``Word.TableRow`` over a ``w:tr``."""

    __slots__ = ("container", "element", "parent_table")

    def __init__(self, element: Tr, container: list, parent_table: Table) -> None:
        """Build the view over a ``w:tr``, the list holding it and its table."""
        #: The ``w:tr``.
        self.element = element
        #: The live list holding it: the table's content, or a row control's.
        self.container = container
        #: The :class:`Table` this row belongs to.
        self.parent_table = parent_table

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:tr`` are equal."""
        return isinstance(other, TableRow) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self.element))

    def __repr__(self) -> str:
        """``<TableRow 1 of body/4, 3 cells>``."""
        return f"<TableRow {self.row_index} of {self.parent_table.ordinal}, {self.cell_count} cells>"

    @property
    def tr(self) -> Tr:
        """The ``w:tr``; docx4j's name for :attr:`element`."""
        return self.element

    @property
    def parent_body(self) -> Body:
        """The body the table is in."""
        return self.parent_table.parent_body

    @property
    def row_index(self) -> int:
        """This row's index among the table's rows, or -1 if it is gone."""
        for index, (element, _container) in enumerate(rows_of(self.parent_table.element)):
            if element is self.element:
                return index
        return -1

    @property
    def address(self) -> str:
        """The ordinal address of the row (``"body/4/0"``). Extension."""
        return ordinal_of(self.parent_body, self.element) or ""

    @property
    def cell_count(self) -> int:
        """The number of cells."""
        return len(cells_of(self.element))

    @property
    def cells(self) -> list[TableCell]:
        """Every cell, in order, cell-level content controls descended into."""
        return [TableCell(element, container, self) for element, container in cells_of(self.element)]

    @property
    def values(self) -> list[str]:
        """The text of the row's cells (Office JS ``TableRow.values``)."""
        return [cell.text for cell in self.cells]

    @values.setter
    def values(self, values: list[str]) -> None:
        with recording(self.parent_body, "set_values") as change:
            cells = self.cells
            for index, value in enumerate(values):
                if index < len(cells):
                    cells[index].value = value
            change.touched(self.parent_table.address)

    @property
    def is_header(self) -> bool:
        """``w:trPr/w:tblHeader``: the row repeats at the top of every page."""
        item = self._tbl_header()
        return item is not None and getattr(item, "val", None) is not False

    @is_header.setter
    def is_header(self, header: bool) -> None:
        item = self._tbl_header()
        if not header:
            tr_pr = self.element.tr_pr
            if item is not None and tr_pr is not None:
                index = _index_of(tr_pr.content, item)
                if index >= 0:
                    del tr_pr.content[index]
            return
        if item is not None:
            return
        tr_pr = self.element.tr_pr
        if tr_pr is None:
            tr_pr = el.trPr()
            self.element.tr_pr = tr_pr
            tr_pr.parent = self.element
        tr_pr.content.append(el.tblHeader())

    def insert_rows(
        self,
        row_count: int = 1,
        *,
        location: ParagraphLocation = "After",
        values: list[list[str]] | None = None,
    ) -> list[TableRow]:
        """Rows before or after this one, with the table's columns and widths."""
        if row_count < 1:
            raise ContentError(
                f"insert_rows needs at least one row, not {row_count}",
                code="table.no_rows",
                hint="pass row_count=1 or more",
            )
        with recording(self.parent_body, "insert_rows") as change:
            widths = self.parent_table.column_widths_twips()
            added = [
                _row_element(widths, values[i] if values is not None and i < len(values) else None)
                for i in range(row_count)
            ]
            index = _index_of(self.container, self.element)
            if index < 0:
                index = len(self.container)
            elif location == "After":
                index += 1
            owner = _owner_of(self.container, self.parent_table.element)
            for offset, row in enumerate(added):
                self.container.insert(index + offset, row)
                link_parents(row)
                row.parent = owner
            _mark_rows_inserted(self.parent_table.change_tracker, added)
            change.touched(self.parent_table.address)
            return [TableRow(row, self.container, self.parent_table) for row in added]

    def delete(self) -> None:
        """Remove the row from its container.

        **While the package tracks changes** the row stays and takes a
        ``w:trPr/w:del`` (CR-003 section 4).
        """
        with recording(self.parent_body, "delete") as change:
            index = _index_of(self.container, self.element)
            if index < 0:
                return
            change.touched(self.parent_table.address)
            change.text(before=text_of(self.element), after="")
            tracker = self.parent_table.change_tracker
            if tracker is not None and not _delete_row(tracker, self.element, self.parent_body):
                return
            del self.container[index]

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: the index, the values and the header flag."""
        return {
            "kind": "row",
            "address": self.address,
            "row_index": self.row_index,
            "cells": self.cell_count,
            "is_header": self.is_header,
            "values": self.values,
        }

    def _tbl_header(self) -> Any:
        tr_pr = self.element.tr_pr
        if tr_pr is None:
            return None
        for item in tr_pr.content:
            if element_name(item) == W_TBL_HEADER:
                return item
        return None


# ---------------------------------------------------------------------------
# TableCell
# ---------------------------------------------------------------------------


class TableCell:
    """A subset of Office JS ``Word.TableCell`` over a ``w:tc``: a body of its own."""

    __slots__ = ("container", "element", "parent_row")

    def __init__(self, element: Tc, container: list, parent_row: TableRow) -> None:
        """Build the view over a ``w:tc``, the list holding it and its row."""
        #: The ``w:tc``.
        self.element = element
        #: The live list holding it: the row's content, or a cell control's.
        self.container = container
        #: The :class:`TableRow` this cell is in.
        self.parent_row = parent_row

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:tc`` are equal."""
        return isinstance(other, TableCell) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self.element))

    def __str__(self) -> str:
        """The cell's text."""
        return self.text

    def __repr__(self) -> str:
        """``<TableCell body/4/0/1 'Total'>``."""
        text = self.text.replace("\n", " ")
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        return f"<TableCell {self.address} {preview!r}>"

    @property
    def tc(self) -> Tc:
        """The ``w:tc``; docx4j's name for :attr:`element`."""
        return self.element

    @property
    def parent_table(self) -> Table:
        """The table this cell is in."""
        return self.parent_row.parent_table

    @property
    def row_index(self) -> int:
        """The index of the row this cell is in."""
        return self.parent_row.row_index

    @property
    def cell_index(self) -> int:
        """This cell's index in its row, or -1 if it is gone."""
        for index, (element, _container) in enumerate(cells_of(self.parent_row.element)):
            if element is self.element:
                return index
        return -1

    @property
    def address(self) -> str:
        """The ordinal address of the cell (``"body/4/0/1"``). Extension."""
        return ordinal_of(self.parent_table.parent_body, self.element) or ""

    @property
    def body(self) -> Body:
        """The cell's content as a :class:`Body` (Office JS ``TableCell.body``).

        The same part and package as the table's body, addressed from the cell,
        so a paragraph in it reports ``body/4/0/1/0``.
        """
        parent = self.parent_table.parent_body
        return parent.sub(self.element, self.address or parent.prefix)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Every paragraph in the cell, nested tables descended into."""
        return self.body.paragraphs

    @property
    def tables(self) -> list[Table]:
        """The tables directly in this cell."""
        return self.body.tables

    @property
    def text(self) -> str:
        """The cell's text, a line per paragraph."""
        return text_of(self.element)

    @property
    def value(self) -> str:
        """Office JS ``TableCell.value``: the text; setting replaces the content."""
        return self.text

    @value.setter
    def value(self, text: str) -> None:
        with recording(self.parent_table.parent_body, "set_value") as change:
            change.text(before=self.text)
            body = self.body
            body.content.clear()
            body.insert_paragraph(text)
            change.touched(self.address)
            change.text(after=text)

    def insert_paragraph(self, text: str = "", *, location: BodyLocation = "End", **options: Any):
        """A paragraph at the start or the end of the cell."""
        return self.body.insert_paragraph(text, location=location, **options)

    def insert_text(self, text: str, *, location: TextLocation = "End") -> Range:
        """Text at the start, the end, or in place of the cell's content."""
        return self.body.insert_text(text, location=location)

    @property
    def width(self) -> float:
        """The cell width in points (``w:tcW`` of type ``dxa``); 0 when it is not fixed."""
        tc_pr = self.element.tc_pr
        tc_w = getattr(tc_pr, "tc_w", None) if tc_pr is not None else None
        if tc_w is None:
            return 0.0
        kind = getattr(tc_w, "type_value", None)
        kind = getattr(kind, "value", kind)
        if kind != "dxa" or tc_w.w is None:
            return 0.0
        return float(tc_w.w) / TWIPS_PER_POINT

    @width.setter
    def width(self, points: float) -> None:
        tc_pr = self.element.tc_pr
        if tc_pr is None:
            tc_pr = el.tcPr()
            self.element.tc_pr = tc_pr
            tc_pr.parent = self.element
        if tc_pr.tc_w is None:
            tc_pr.tc_w = el.tcW()
            tc_pr.tc_w.parent = tc_pr
        tc_pr.tc_w.w = round(points * TWIPS_PER_POINT)
        tc_pr.tc_w.type_value = "dxa"

    @property
    def column_width(self) -> float:
        """Office JS ``TableCell.columnWidth``: the same number as :attr:`width`."""
        return self.width

    @column_width.setter
    def column_width(self, points: float) -> None:
        self.width = points

    def get_xml(self) -> str:
        """The cell as XML, with docx4j's prefixes. Extension."""
        return to_xml(self.element)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: the position, the text and the width."""
        return {
            "kind": "cell",
            "address": self.address,
            "row_index": self.row_index,
            "cell_index": self.cell_index,
            "text": self.text,
            "width": self.width,
        }


# ---------------------------------------------------------------------------
# the cell an element is in, and inserting a table
# ---------------------------------------------------------------------------


def cell_of(value: Any, parent_body: Body) -> TableCell | None:
    """The cell a paragraph or a table is in, from the parent pointers, or None.

    Shared by ``Paragraph.parent_table_cell`` and :attr:`Table.parent_table_cell`.
    The walk goes up through the four ``w:sdt`` forms without naming them,
    because a cell-level control sits between a cell and its row.
    """
    cell = _ancestor(value, W_TC, include_self=True)
    if cell is None:
        return None
    row = _ancestor(cell, W_TR)
    if row is None:
        return None
    table = _ancestor(row, W_TBL)
    if table is None:
        return None
    container = block_children_of(getattr(table, "parent", None) or ())
    if container is None or _index_of(container, table) < 0:
        container = getattr(parent_body, "content", [])
    view = Table(table, container, parent_body)
    row_view = next((r for r in view.rows if r.element is row), None)
    if row_view is None:
        return None
    return next((c for c in row_view.cells if c.element is cell), None)


def _ancestor(value: Any, qname: str, *, include_self: bool = False) -> Any:
    """The nearest ancestor of an element name, or None."""
    current = value if include_self else getattr(value, "parent", None)
    for _ in range(_MAX_DEPTH):
        if current is None:
            return None
        if element_name(current) == qname:
            return current
        current = getattr(current, "parent", None)
    return None


def insert_table_into(
    body: Body,
    row_count: int,
    column_count: int,
    *,
    location: BodyLocation = "End",
    values: list[list[str]] | None = None,
    style: str | None = None,
) -> Table:
    """What :meth:`Body.insert_table` is. See its docstring."""
    if row_count < 1 or column_count < 1:
        raise InvalidTargetError(
            f"a table needs at least one row and one column, not {row_count} by {column_count}",
            code="table.empty",
            hint="pass insert_table(rows, columns) with both at least 1",
        )
    with recording(body, "insert_table") as change:
        rows = [
            [
                (values[r][c] if values is not None and r < len(values) and c < len(values[r]) else "")
                for c in range(column_count)
            ]
            for r in range(row_count)
        ]
        element = tbl(rows, width=writable_width(body))
        if style is not None:
            resolved = style_id_of(body.package, style, define=True)
            tbl_pr = element.tbl_pr
            tbl_pr.tbl_style = el.tblStyle(val=resolved)
            tbl_pr.tbl_style.parent = tbl_pr
        body.insert_element(element, location=location)
        view = Table(element, body.content, body)
        change.touched(view.address)
        change.text(after="\n".join("\t".join(row) for row in rows))
        return view


# ``Tc`` and ``Tr`` are imported for the annotations above to resolve.
_ = (Tc, Tr)
