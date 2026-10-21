"""A chart described as a table, for a reader without the picture. CR-011 section 6.

docx4j's ``org.docx4j.model.chart.ChartTable`` at ``198d015a3`` (its CR-035):
the chart's title as Office shows it, and its data, categories down and the
series Office shows across, each number formatted by its format code as Excel
formats it (in the US locale, :mod:`~docx4j_py.model.chart.formats`).
Multi-level categories take a column a level, the outermost first. A chart
with no categories (scatter, bubble) lists each series' x and y values (and
bubble sizes), point by point. A gap is an empty cell; an error (``#N/A``) is
written as it is.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from docx4j_py.model.chart.formats import format_raw_cell_contents

__all__ = ["ChartTable", "format_number"]


def format_number(text: str | None, code: str | None, date1904: bool = False) -> str:
    """A number formatted by its format code as Excel does; anything else as it is."""
    if text is None:
        return ""
    try:
        value = float(text.strip())
    except ValueError:
        return text
    code = code or "General"
    try:
        return format_raw_cell_contents(value, 0 if code.lower() == "general" else -1, code, date1904)
    except Exception:  # noqa: BLE001 - docx4j writes the text where its formatter throws
        return text


def _cell(data: Any, index: int, date1904: bool) -> str:
    if data is None:
        return ""
    text = data.text(index)
    if not data.numeric:
        return text or ""
    return format_number(text, data.format_code_of(index), date1904)


def _size(data: Any) -> int:
    return 0 if data is None else data.size()


@dataclass(slots=True, eq=False)
class ChartTable:
    """A chart's title, column headings and rows. docx4j ``ChartTable``."""

    #: The chart's title as Office shows it; None for none.
    title: str | None
    #: The column headings: the categories' columns (empty), then the series shown.
    header: list[str] = field(default_factory=list)
    #: The rows, each as long as the header.
    rows: list[list[str]] = field(default_factory=list)

    @classmethod
    def of(cls, model: Any) -> ChartTable:
        """A chart's table (a :class:`~docx4j_py.model.chart.model.ChartModel`)."""
        date1904 = model.date1904 is True
        shown = [s for g in model.groups for s in g.shown_series]
        title = model.title.shown_text if model.title is not None else None
        header: list[str] = []
        rows: list[list[str]] = []
        categories = next((s.categories for s in shown if s.categories is not None), None)
        xy = categories is None or any(s.x_values is not None for s in shown)
        if xy:
            # x and y: one x column where every series has the same x values, the usual case;
            # else each series' own x beside its y (and bubble size)
            shared = True
            first_x = None
            for s in shown:
                x = s.x_values if s.x_values is not None else s.categories
                if first_x is None:
                    first_x = x
                elif x is None or x.points != first_x.points:
                    shared = False
                if s.bubble_sizes is not None:
                    shared = False
            n = max((_size(s.y_values if s.y_values is not None else s.values) for s in shown), default=0)
            if shared:
                header.append("")
                header.extend(s.shown_name for s in shown)
            else:
                for s in shown:
                    header.append(f"{s.shown_name} (x)")
                    header.append(s.shown_name)
                    if s.bubble_sizes is not None:
                        header.append(f"{s.shown_name} (size)")
            for i in range(n):
                row: list[str] = []
                for k, s in enumerate(shown):
                    y = s.y_values if s.y_values is not None else s.values
                    x = s.x_values if s.x_values is not None else s.categories
                    if not shared or k == 0:
                        row.append(str(i + 1) if x is None else _cell(x, i, date1904))
                    row.append(_cell(y, i, date1904))
                    if not shared and s.bubble_sizes is not None:
                        row.append(_cell(s.bubble_sizes, i, date1904))
                rows.append(row)
            return cls(title, header, rows)
        levels = categories.levels
        header.extend("" for _ in levels)
        header.extend(s.shown_name for s in shown)
        n = max([categories.size(), *(_size(s.values) for s in shown)])
        for i in range(n):
            row = []
            for level in range(len(levels) - 1, -1, -1):
                if categories.numeric and len(levels) == 1:
                    row.append(_cell(categories, i, date1904))
                else:
                    row.append(levels[level].get(i) or "")
            row.extend(_cell(s.values, i, date1904) for s in shown)
            rows.append(row)
        return cls(title, header, rows)

    @classmethod
    def of_chartex(cls, model: Any) -> ChartTable:
        """A chartex chart's table: its shown series' categories and values from the caches."""
        shown = [s for s in model.series if not s.hidden]
        category = None
        for s in shown:
            for d in s.data:
                if category is None and not d.numeric and d.levels:
                    category = d
        levels = len(category.levels) if category is not None else 0
        header = [""] * levels + [s.name or "" for s in shown]
        n = max(category.counts, default=0) if category is not None else 0
        for s in shown:
            values = _values(s)
            if values is not None and values.counts:
                n = max(n, values.counts[0])
        rows = []
        for i in range(n):
            row = [category.levels[level].get(i) or "" for level in range(levels - 1, -1, -1)] if category else []
            for s in shown:
                values = _values(s)
                text = values.levels[0].get(i) if values is not None and values.levels else None
                code = values.format_codes[0] if values is not None and values.format_codes else None
                row.append(format_number(text, code))
            rows.append(row)
        return cls(model.shown_title, header, rows)


def _values(series: Any) -> Any:
    return next((d for d in series.data if d.numeric), None)
