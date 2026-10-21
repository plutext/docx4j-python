"""Plain text from the block model (WD section 7.1). CR-011 section 3.2 item 4.

docx4j's ``org.docx4j.model.text.PlainText`` at ``198d015a3``.

Each paragraph is one line, a break one more (REQ-603), the text exactly as
the document holds it (REQ-602: tabs, both hyphens, zero-width and tag
characters kept), U+000A between lines and after the last. The characters are
section 6.1's: ``w:tab`` and ``w:ptab`` a tab, ``w:br`` and ``w:cr`` a line
feed, ``w:noBreakHyphen`` U+2011, ``w:softHyphen`` U+00AD, ``w:sym`` the
character its ``w:char`` names.

**Tables** (REQ-605): a simple table (no merged or split cells, no nested
table, one paragraph a cell, every row the same cell edges) a row a line, its
cells a tab apart; any other in the listed format, after the note line
``[table: listed format]``, a row a line as ``Header: value; Header: value``,
the headers from the ``w:tblHeader`` rows else the first row, a cell's
paragraphs (a nested table's included) ``; `` apart; a row whose one cell
spans the grid is its text alone, and a table with no labels has its cells a
tab apart. ``tables: tsv`` writes every table in the simple form, a cell's
paragraphs a space apart; ``tables: grid`` as a pandoc grid table, its spans
drawn.

What the extractor adds follows the WD's bracket convention (REQ-707):
``[n]`` at a note's reference and at the start of the note; ``[image: ...]``
(or ``[image, generated: ...]``) under ``alt_text: true``; ``[page n]`` under
``page_markers: true``; and under ``metadata: true``, before the body, a line
a kind of property that has any (REQ-515): ``[core properties: title: ...;
subject: ...]``, ``[application properties: Company: ...]``, ``[custom
properties: name: value]``. In the ``markup`` view insertions are
``{++text++}``, deletions ``{--text--}`` and comments ``{>>text<<}`` at their
references, replies after them (REQ-702, REQ-504), each run's marker its own,
flat, never one inside another; under ``markup_detail`` a marker's opening is
followed by ``author:`` or ``author, date:``.

Under ``layout: indent`` (REQ-709) a paragraph outside a table begins with
four spaces for each tab stop (720 twips) of its first line's indent, its left
indent plus its first-line indent less its hanging indent, rounded to the
nearest stop; its tabs are kept.

A grid table's widths are counted in UTF-16 code units, as Java's
``String.length`` counts them, so that a grid is docx4j's character for
character.
"""

from __future__ import annotations

import math
from typing import Any

from docx4j_py.model.text.result import (
    TextBlock,
    TextCell,
    TextComment,
    TextResult,
    TextRevision,
    TextRow,
    TextRun,
)

__all__ = ["INDENT_SPACES", "LISTED_FORMAT_NOTE", "is_simple", "plain_text"]

#: The spaces for each 720-twip tab stop of indent under ``layout: indent`` (REQ-709).
INDENT_SPACES = 4
#: The note line before a table in the listed format (REQ-605, REQ-707).
LISTED_FORMAT_NOTE = "[table: listed format]"


def plain_text(result: TextResult) -> str:
    """The plain text of a result (WD sections 7.1 and 7.5)."""
    writer = _Writer(result)
    if result.options.metadata:
        writer.metadata(result.metadata)
    for block in result.blocks:
        writer.top(block)
        writer.out.append("\n")
    return "".join(writer.out)


def _width(text: str) -> int:
    """Java's ``String.length``: UTF-16 code units."""
    return len(text.encode("utf-16-le")) // 2


def _edges(row: TextRow) -> frozenset[int]:
    """A row's cell edges in grid columns, by ``gridSpan``: 0, and each cell's end."""
    edges = {0}
    col = 0
    for cell in row.cells:
        col += max(1, cell.grid_span)
        edges.add(col)
    return frozenset(edges)


def _starts(row: TextRow) -> list[int]:
    """The grid columns each cell of a row starts at, by ``gridSpan``."""
    out = []
    col = 0
    for cell in row.cells:
        out.append(col)
        col += max(1, cell.grid_span)
    return out


def _header_rows(rows: list[TextRow]) -> int:
    """The header rows: those marked ``w:tblHeader``, else the first."""
    n = 0
    while n < len(rows) and rows[n].header:
        n += 1
    return n if n > 0 else min(1, len(rows))


def is_simple(table: TextBlock) -> bool:
    """No merged or split cells, no nested table, one paragraph a cell (REQ-605).

    A cell is merged or split when it crosses a column edge another row has,
    or is in a vertical merge; a cell spanning grid columns that every row
    spans alike merges nothing (Word leaves such columns behind). A cell's
    empty paragraphs don't count: Word ends every cell with a paragraph, so a
    block-level control in a cell comes with an empty one after it.
    """
    edges: frozenset[int] | None = None
    for row in table.rows:
        own = _edges(row)
        if edges is not None and edges != own:
            return False
        edges = own
        for cell in row.cells:
            if cell.merge_continue or cell.merge_restart:
                return False
            paragraphs = 0
            for b in cell.blocks:
                if b.kind != "PARAGRAPH":
                    return False
                if b.runs or b.list_label is not None:
                    paragraphs += 1
            if paragraphs > 1:
                return False
    return True


def _width_of_grid(rows: list[TextRow]) -> int:
    """The grid's width: the most grid columns a row spans."""
    return max((sum(max(1, c.grid_span) for c in row.cells) for row in rows), default=0)


def _full_width(row: TextRow, width: int) -> bool:
    """A row of one cell spanning every grid column: a title, or a divider such as "Part A"."""
    return len(row.cells) == 1 and max(1, row.cells[0].grid_span) == width


def _header(headers: dict[int, str], start: int, span: int) -> str | None:
    """The first header with text over a cell's grid columns, or None."""
    for g in range(start, start + span):
        h = headers.get(g)
        if h:
            return h
    return None


def _label(header: str | None, value: str) -> str:
    return value if not header else f"{header}: {value}"


class _Placed:
    """A cell placed on the grid."""

    __slots__ = ("col", "lines", "row", "row_span", "span")

    def __init__(self, row: int, col: int, span: int, lines: list[str]) -> None:
        self.row = row
        self.col = col
        self.span = span
        self.row_span = 1
        self.lines = lines


class _Writer:
    __slots__ = ("detail", "options", "out", "result")

    def __init__(self, result: TextResult) -> None:
        self.result = result
        self.options = result.options
        self.detail = result.options.markup_detail
        self.out: list[str] = []

    # -- the metadata block (REQ-515) ----------------------------------------

    def metadata(self, metadata: dict) -> None:
        """A line a kind of property that has any, before the body."""
        self._group("core properties", metadata.get(TextResult.CORE))
        self._group("application properties", metadata.get(TextResult.APP))
        self._group("custom properties", metadata.get(TextResult.CUSTOM))

    def _group(self, label: str, properties: object) -> None:
        if not isinstance(properties, dict) or not properties:
            return
        from docx4j_py.model.text.metadata import shown

        self.out.append(f"[{label}: ")
        self.out.append("; ".join(f"{name}: {shown(value)}" for name, value in properties.items()))
        self.out.append("]\n")

    # -- blocks ------------------------------------------------------------------

    def top(self, block: TextBlock) -> None:
        """A block of the story: a paragraph with its indent, a table, a chart."""
        if block.kind == "PARAGRAPH" and self.options.layout == "indent":
            self._indent(block)
        self.block(block)

    def _indent(self, block: TextBlock) -> None:
        twips = (block.indent_left or 0) + (block.indent_first_line or 0) - (block.indent_hanging or 0)
        stops = math.floor(twips / 720.0 + 0.5)  # Java's Math.round
        if stops > 0:
            self.out.append(" " * (stops * INDENT_SPACES))

    def block(self, block: TextBlock) -> None:
        if block.kind == "TABLE":
            self.table(block)
        elif block.kind == "CHART":
            self.chart(block)
        else:
            self.paragraph(block)

    def chart(self, block: TextBlock) -> None:
        """A chart: its note (docx4j CR-047's description), its title, then its table's rows a tab apart (REQ-511)."""
        lines: list[str] = []
        if block.chart is not None:
            lines.append(block.chart.note)
        if block.chart_title:
            lines.append(block.chart_title)
        if block.chart_header:
            lines.append("\t".join(block.chart_header))
        lines.extend("\t".join(row) for row in block.chart_rows)
        self.out.append("\n".join(lines))

    def paragraph(self, block: TextBlock) -> None:
        if block.list_label is not None:
            self.out.append(block.list_label)
        runs = block.runs
        i = 0
        while i < len(runs):
            revision = runs[i].revision
            j = i
            while j < len(runs) and self._same_marker(revision, runs[j].revision):
                j += 1
            if revision is not None:
                self.out.append("{++" if revision.type == "INSERTION" else "{--")
                self._detail(revision.author, revision.date)
            for k in range(i, j):
                self.run(runs[k])
            if revision is not None:
                self.out.append("++}" if revision.type == "INSERTION" else "--}")
            i = j

    def _same_marker(self, a: TextRevision | None, b: TextRevision | None) -> bool:
        """Runs under one marker: the same type, and the same author and date where the detail shows them."""
        if a is None or b is None:
            return a is b
        if a.type != b.type:
            return False
        if self.detail == "author":
            return a.author == b.author
        if self.detail == "author_date":
            return a.author == b.author and a.date == b.date
        return True

    def _detail(self, author: str | None, date: str | None) -> None:
        if self.detail == "author":
            self.out.append(f"{author or ''}: ")
        elif self.detail == "author_date":
            self.out.append(f"{author or ''}, {date or ''}: ")

    def run(self, run: TextRun) -> None:
        kind = run.kind
        if kind in ("TEXT", "FIELD_CODE", "LIST_LABEL"):
            self.out.append(run.text)
        elif kind == "NOTE_REFERENCE":
            if self.options.note_references and self.options.notes and run.note_number > 0:
                self.out.append(f"[{run.note_number}]")
        elif kind == "NOTE_MARK":
            if run.note_number > 0:
                self.out.append(f"[{run.note_number}]")
        elif kind == "IMAGE":
            description = run.description
            if self.options.alt_text and description is not None:
                opening = "[image, generated: " if description.provenance == "GENERATED" else "[image: "
                self.out.append(f"{opening}{description.text}]")
        elif kind == "PAGE_MARKER":
            if self.options.page_markers:
                self.out.append(f"[page {run.page_number}]")
        elif kind == "COMMENT_REFERENCE":
            comments = self.result.comments
            comment = comments.get(run.comment_id or "")
            if comment is None or (comment.parent_id is not None and comment.parent_id in comments):
                return
            self.comment(comment)
            for reply in comment.replies:
                self.comment(reply)

    def comment(self, comment: TextComment) -> None:
        self.out.append("{>>")
        self._detail(comment.author, comment.date)
        self.out.append(f"{comment.text}<<}}")

    def _text_of(self, block: TextBlock) -> str:
        """A block written on its own, for a table cell's value."""
        saved = self.out
        self.out = []
        try:
            self.block(block)
            return "".join(self.out)
        finally:
            self.out = saved

    # -- tables (REQ-605) ----------------------------------------------------------

    def table(self, table: TextBlock) -> None:
        form = "tsv" if table.simple_form else self.options.tables
        if form == "grid":
            self._grid(table)
        elif form == "tsv" or is_simple(table):
            self._tab_separated(table)
        else:
            self._listed(table)

    def _tab_separated(self, table: TextBlock) -> None:
        """A row a line, its cells a tab apart, a cell's non-empty paragraphs a space apart."""
        for r, row in enumerate(table.rows):
            if r > 0:
                self.out.append("\n")
            for c, cell in enumerate(row.cells):
                if c > 0:
                    self.out.append("\t")
                self.out.append(" ".join(self._values(cell)))

    def _values(self, cell: TextCell, out: list[str] | None = None) -> list[str]:
        """A cell's paragraphs, a nested table's included, empty ones left out."""
        if out is None:
            out = []
        for b in cell.blocks:
            if b.kind == "TABLE":
                for row in b.rows:
                    for inner in row.cells:
                        self._values(inner, out)
            else:
                text = self._text_of(b)
                if text:
                    out.append(text)
        return out

    def _value(self, cell: TextCell) -> str:
        """A cell's paragraphs, a nested table's included, ``; `` apart, empty ones left out."""
        return "; ".join(self._values(cell))

    def _listed(self, table: TextBlock) -> None:
        """The listed format (REQ-605, as ruled on 18 October 2026).

        Every cell is labelled by the header over its first grid column, or the
        first of its columns' headers that has text; several header rows' texts
        are joined `` / ``; a cell under a header with no text is its value
        alone. A row whose one cell spans every grid column, a header row or a
        body row, is that cell's text alone, and labels nothing. A cell with no
        value is left out; a vertical merge's value is on every row it covers.
        Labelled cells are ``; `` apart; in a table with no labels (a table of
        header rows only, or whose headers span the grid or have no text) cells
        are a tab apart, as in the simple form.
        """
        self.out.append(LISTED_FORMAT_NOTE)
        rows = table.rows
        width = _width_of_grid(rows)
        header_rows = _header_rows(rows)
        # a table of header rows only is written unlabelled, its rows as body rows
        headers_only = header_rows >= len(rows)
        # each grid column's header; a full-width header row is written once, as its text
        headers: dict[int, str] = {}
        if not headers_only:
            for r in range(header_rows):
                row = rows[r]
                if _full_width(row, width):
                    text = self._value(row.cells[0])
                    if text:
                        self.out.append(f"\n{text}")
                    continue
                starts = _starts(row)
                for c, cell in enumerate(row.cells):
                    text = self._value(cell)
                    if not text:
                        continue
                    for g in range(starts[c], starts[c] + max(1, cell.grid_span)):
                        h = headers.get(g)
                        headers[g] = text if h is None else f"{h} / {text}"
        labelled = bool(headers)
        body = rows if headers_only else rows[header_rows:]
        restarts: dict[int, str] = {}
        for row in body:
            self.out.append("\n")
            if _full_width(row, width):
                self.out.append(self._value(row.cells[0]))
                continue
            starts = _starts(row)
            parts = []
            for c, cell in enumerate(row.cells):
                if cell.merge_continue:
                    value = restarts.get(starts[c], "")
                else:
                    value = self._value(cell)
                    if cell.merge_restart:
                        restarts[starts[c]] = value
                if not value:
                    continue
                parts.append(_label(_header(headers, starts[c], max(1, cell.grid_span)), value) if labelled else value)
            self.out.append(("; " if labelled else "\t").join(parts))

    # -- pandoc grid tables ----------------------------------------------------------

    def _grid(self, table: TextBlock) -> None:
        for line in self._grid_lines(table):
            text = "".join(self.out)
            if text and not text.endswith("\n"):
                self.out.append("\n")
            self.out.append(line)

    def _grid_lines(self, table: TextBlock) -> list[str]:
        rows = table.rows
        placed: list[_Placed] = []
        # the grid's columns are those some row's cell edges mark, so a span every row shares draws as one column
        union: set[int] = set()
        for row in rows:
            union |= _edges(row)
        edge_list = sorted(union)
        columns = len(edge_list) - 1
        if columns <= 0:
            return []
        at: list[list[_Placed | None]] = []
        open_: dict[int, _Placed] = {}
        for r, row in enumerate(rows):
            at.append([None] * columns)
            starts = _starts(row)
            for c, cell in enumerate(row.cells):
                first = edge_list.index(starts[c])
                span = edge_list.index(starts[c] + max(1, cell.grid_span)) - first
                found = open_.get(first)
                if cell.merge_continue and found is not None and found.span == span:
                    p = found
                    p.row_span += 1
                else:
                    p = _Placed(r, first, span, self._cell_lines(cell))
                    placed.append(p)
                    if cell.merge_restart:
                        open_[first] = p
                    else:
                        open_.pop(first, None)
                for g in range(first, min(first + span, columns)):
                    at[r][g] = p
        # column widths: single-column cells first, then spans widen their last column
        widths = [3] * columns
        for p in placed:
            if p.span == 1:
                for line in p.lines:
                    widths[p.col] = max(widths[p.col], _width(line))
        for p in placed:
            if p.span == 1:
                continue
            last = min(columns, p.col + p.span) - 1
            have = sum(widths[p.col : last + 1]) + 3 * (last - p.col)
            for line in p.lines:
                if _width(line) > have:
                    widths[last] += _width(line) - have
                    have = _width(line)
        # row heights: a cell's lines over the rows it spans (the borders it crosses hold none)
        heights = [1] * len(rows)
        for p in placed:
            if p.row_span == 1:
                heights[p.row] = max(heights[p.row], len(p.lines))
        for p in placed:
            if p.row_span == 1:
                continue
            have = sum(heights[p.row : p.row + p.row_span])
            if len(p.lines) > have:
                heights[p.row + p.row_span - 1] += len(p.lines) - have
        header = _header_rows(rows) if len(rows) > 1 else 0
        out = [_border(None, at[0] if at else None, widths, "-")]
        written: dict[int, int] = {}
        for r in range(len(rows)):
            for _line in range(heights[r]):
                parts = ["|"]
                g = 0
                while g < columns:
                    p = at[r][g]
                    span = 1 if p is None else min(p.span, columns - g)
                    width = sum(widths[g : g + span]) + 3 * (span - 1)
                    text = ""
                    if p is not None:
                        n = written.get(id(p), 0)
                        if n < len(p.lines):
                            text = p.lines[n]
                        written[id(p)] = n + 1
                    parts.append(" " + text + " " * max(0, width - _width(text)) + " |")
                    g += span
                out.append("".join(parts))
            if r < len(rows) - 1:
                out.append(_border(at[r], at[r + 1], widths, "=" if r == header - 1 else "-"))
        out.append(_border(at[len(rows) - 1], None, widths, "-"))
        return out

    def _cell_lines(self, cell: TextCell) -> list[str]:
        """A cell's lines: each paragraph's, a blank line between paragraphs, a nested table drawn as a grid."""
        lines: list[str] = []
        for b in cell.blocks:
            if b.kind == "TABLE":
                own = self._grid_lines(b)
            else:
                text = self._text_of(b)
                if not text:
                    continue
                own = text.split("\n")
            if lines:
                lines.append("")
            lines.extend(own)
        return lines


def _border(above: list[Any] | None, below: list[Any] | None, widths: list[int], fill: str) -> str:
    """The border between two rows: open where a cell spans both."""
    out = []
    for g, width in enumerate(widths):
        through = above is not None and below is not None and above[g] is not None and above[g] is below[g]
        join_left = (
            g > 0
            and above is not None
            and below is not None
            and above[g] is not None
            and above[g] is below[g]
            and above[g - 1] is above[g]
        )
        out.append(("|" if through else "+") if g == 0 else (" " if join_left else "+"))
        out.append((" " if through else fill) * (width + 2))
    last = len(widths) - 1
    through_last = above is not None and below is not None and above[last] is not None and above[last] is below[last]
    out.append("|" if through_last else "+")
    return "".join(out)
