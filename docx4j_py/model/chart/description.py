"""A chart put into words beside its table (docx4j CR-047). CR-011 section 6.

docx4j's ``org.docx4j.model.chart.ChartDescription`` at ``198d015a3``: what
kind of chart it is, the titles of the axes Office shows, and which series are
on secondary axes. The text model (WD REQ-511) writes its :attr:`note` before
the chart's title, so that a reader without the picture knows how to take the
numbers.

* **The kind** is Excel's English name for each group's chart type, in group
  order, joined by "and": clustered column, stacked bar, 100% stacked area,
  3-D column, line, pie, bar of pie, doughnut, scatter, bubble, filled radar,
  stock, 3-D surface, wireframe contour; for chartex, the first series' layout
  (waterfall, funnel, histogram, box and whisker, sunburst, treemap, map), and
  pareto for a histogram with its Pareto line.
* **The axes** are the ones Office shows, or shows the title of, in reading
  order, each with its role: category axis, value axis, secondary value axis,
  secondary category axis, depth axis; for scatter and bubble charts, x axis
  and y axis, then the secondary ones. An axis's title is the one Office
  shows, its default "Axis Title" included. A deleted axis (for chartex, a
  hidden one) is left out unless it has a title. Pie, doughnut and of-pie
  charts have no axes.
* **A group is on the secondary axes** when its value axis is not the first
  group's value axis; its shown series are the secondary series.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any

from docx4j_py.model.chart.table import ChartTable

__all__ = ["ChartDescription"]


@dataclass(frozen=True, slots=True)
class Group:
    """A chart group put into words."""

    #: Excel's name for the group's kind: "clustered column", "line", "waterfall" ...
    kind: str
    #: The element: ``barChart``, ``line3DChart`` ...; for chartex, the first series' ``layoutId``.
    element: str
    #: ``c:barDir`` (bar or col) for a bar chart; else None.
    bar_direction: str | None
    #: ``c:grouping`` as stated; None where the kind has none.
    grouping: str | None
    #: The shown series' names, as Office shows them.
    series: tuple[str, ...]
    #: Whether the group is on the secondary axes.
    secondary: bool


@dataclass(frozen=True, slots=True)
class Axis:
    """An axis Office shows."""

    #: category axis, value axis, secondary value axis, secondary category axis, depth axis; x axis, y axis ...
    role: str
    #: ``c:axPos``: b, l, r or t; None for chartex.
    position: str | None
    #: The title Office shows, its default "Axis Title" included; None for an axis with none.
    title: str | None


def _name(group: Any) -> str:
    """Excel's English name for a group's kind (CR-047 section 2.2)."""
    grouping = group.grouping
    kind = group.kind
    if kind in ("bar", "bar3D"):
        base = "bar" if group.bar_direction == "bar" else "column"
        stacking = (
            "clustered " if grouping in (None, "clustered")
            else "stacked " if grouping == "stacked"
            else "100% stacked " if grouping == "percentStacked"
            else ""
        )  # fmt: skip
        return ("3-D " if kind == "bar3D" else "") + stacking + base
    if kind == "line":
        return _stacked(grouping) + "line"
    if kind == "line3D":
        return "3-D line"
    if kind == "area":
        return _stacked(grouping) + "area"
    if kind == "area3D":
        return "3-D " + _stacked(grouping) + "area"
    if kind == "pie":
        return "pie"
    if kind == "pie3D":
        return "3-D pie"
    if kind == "ofPie":
        return "bar of pie" if group.settings.get("ofPieType") == "bar" else "pie of pie"
    if kind in ("doughnut", "scatter", "bubble", "stock"):
        return kind
    if kind == "radar":
        return "filled radar" if group.settings.get("radarStyle") == "filled" else "radar"
    if kind == "surface":
        return "wireframe contour" if group.settings.get("wireframe") == "true" else "contour"
    if kind == "surface3D":
        return "wireframe 3-D surface" if group.settings.get("wireframe") == "true" else "3-D surface"
    return kind


def _stacked(grouping: str | None) -> str:
    return "stacked " if grouping == "stacked" else "100% stacked " if grouping == "percentStacked" else ""


def _axes_of(model: Any, group: Any) -> list[Any]:
    """The group's axes, in its ``c:axId`` order."""
    out: list[Any] = []
    for axis_id in group.axis_ids:
        axis = model.axis(axis_id)
        if axis is not None and axis not in out:
            out.append(axis)
    return out


def _of_kind(axes: list[Any], kind: str) -> Any:
    return next((a for a in axes if a.kind == kind), None)


def _category(axes: list[Any]) -> Any:
    return next((a for a in axes if a.kind in ("CATEGORY", "DATE")), None)


def _x_axis(axes: list[Any]) -> Any:
    """Of a scatter or bubble chart's two value axes, the one at the bottom or top; else the first."""
    first = None
    for a in axes:
        if a.kind != "VALUE":
            continue
        if a.position in ("b", "t"):
            return a
        if first is None:
            first = a
    return first


def _other(axes: list[Any], not_: Any, kind: str) -> Any:
    return next((a for a in axes if a is not not_ and a.kind == kind), None)


def _value_axis(model: Any, group: Any, xy: bool) -> Any:
    """A group's value axis: its first ``c:valAx``; for scatter and bubble, the one at the left or right (y)."""
    axes = _axes_of(model, group)
    if xy:
        return _other(axes, _x_axis(axes), "VALUE")
    return _of_kind(axes, "VALUE")


def _add(axes: list[Axis], role: str, axis: Any) -> None:
    """An axis Office shows, or shows the title of: a deleted axis's title still shows."""
    if axis is None or (axis.deleted and axis.title is None):
        return
    axes.append(Axis(role, axis.position, None if axis.title is None else axis.title.shown_text))


_LAYOUTS = {"clusteredColumn": "histogram", "paretoLine": "pareto", "boxWhisker": "box and whisker", "regionMap": "map"}


def layout_name(layout: str | None) -> str:
    """Excel's English name for a chartex series layout."""
    if layout is None:
        return "chart"
    return _LAYOUTS.get(layout, layout)  # waterfall, funnel, sunburst, treemap


@dataclass(frozen=True, slots=True)
class ChartDescription:
    """A chart's kind, its axes and its secondary series, beside its table. docx4j ``ChartDescription``."""

    #: Each group's kind, in group order, joined by "and".
    kind: str
    groups: tuple[Group, ...]
    #: The axes Office shows, in reading order.
    axes: tuple[Axis, ...]
    #: The series on secondary axes, by their shown names; empty for most charts.
    secondary_series: tuple[str, ...]
    #: The chart's table.
    table: ChartTable = field(repr=False)

    @classmethod
    def of(cls, model: Any) -> ChartDescription:
        """A chart's description: a :class:`~docx4j_py.model.chart.model.ChartModel`."""
        groups_in = model.groups
        first = groups_in[0] if groups_in else None
        xy = first is not None and first.kind in ("scatter", "bubble")
        primary_value = None if first is None else _value_axis(model, first, xy)
        groups: list[Group] = []
        secondary_series: list[str] = []
        kinds: list[str] = []
        secondary_group = None
        for g in groups_in:
            value = _value_axis(model, g, xy)
            secondary = g is not first and value is not None and value is not primary_value
            if secondary and secondary_group is None:
                secondary_group = g
            names = tuple(s.shown_name for s in g.shown_series)
            if secondary:
                secondary_series.extend(names)
            name = _name(g)
            kinds.append(name)
            groups.append(Group(name, g.kind + "Chart", g.bar_direction, g.grouping, names, secondary))
        axes: list[Axis] = []
        if first is not None:
            primary = _axes_of(model, first)
            secondary_axes = (
                [a for a in _axes_of(model, secondary_group) if a not in primary] if secondary_group is not None else []
            )
            if xy:
                x = _x_axis(primary)
                _add(axes, "x axis", x)
                _add(axes, "y axis", _other(primary, x, "VALUE"))
                x2 = _x_axis(secondary_axes)
                _add(axes, "secondary x axis", x2)
                _add(axes, "secondary y axis", _other(secondary_axes, x2, "VALUE"))
            else:
                _add(axes, "category axis", _category(primary))
                _add(axes, "value axis", _of_kind(primary, "VALUE"))
                _add(axes, "secondary value axis", _of_kind(secondary_axes, "VALUE"))
                _add(axes, "secondary category axis", _category(secondary_axes))
                depth = _of_kind(primary, "SERIES")
                _add(axes, "depth axis", depth if depth is not None else _of_kind(secondary_axes, "SERIES"))
        return cls(" and ".join(kinds), tuple(groups), tuple(axes), tuple(secondary_series), ChartTable.of(model))

    @classmethod
    def of_chartex(cls, model: Any) -> ChartDescription:
        """A chartex chart's description: a :class:`~docx4j_py.model.chart.chartex.ChartExModel`."""
        names: list[str] = []
        first_layout = None
        histogram = pareto_line = False
        for s in model.series:
            if first_layout is None:
                first_layout = s.layout
            if s.layout == "clusteredColumn":
                histogram = True
            if s.layout == "paretoLine":
                pareto_line = True
            if not s.hidden:
                names.append(s.name)  # type: ignore[arg-type]
        kind = "pareto" if histogram and pareto_line else layout_name(first_layout)
        groups = (Group(kind, first_layout, None, None, tuple(names), False),) if first_layout is not None else ()
        axes: list[Axis] = []
        values = 0
        for a in model.axes:
            if a.kind == "category":
                role = "category axis"
            elif a.kind == "value":
                role = "value axis" if values == 0 else "secondary value axis"
                values += 1
            else:
                continue
            if a.hidden and not a.has_title:
                continue
            axes.append(Axis(role, None, a.shown_title))
        return cls(kind, groups, tuple(axes), (), ChartTable.of_chartex(model))

    @property
    def note(self) -> str:
        """The bracket note a reader writes before the chart's title (WD REQ-511, REQ-707).

        ``[chart: <kind>; <role>: <title>; ...; secondary axis: <series>, ...]``,
        an axis only where it has a title; the titles and series names are
        document text inside the extractor's brackets.
        """
        parts = [f"[chart: {self.kind}"]
        parts.extend(f"; {a.role}: {a.title}" for a in self.axes if a.title is not None)
        if self.secondary_series:
            parts.append("; secondary axis: " + ", ".join(self.secondary_series))
        return "".join(parts) + "]"

    def to_dict(self) -> dict[str, Any]:
        """The description as docx4j's JSON has it, its members in its order (camelCase, as docx4j's)."""
        return {
            "kind": self.kind,
            "groups": [
                {
                    "kind": g.kind,
                    "element": g.element,
                    "barDirection": g.bar_direction,
                    "grouping": g.grouping,
                    "series": list(g.series),
                    "secondary": g.secondary,
                }
                for g in self.groups
            ],
            "axes": [{"role": a.role, "position": a.position, "title": a.title} for a in self.axes],
            "secondarySeries": list(self.secondary_series),
            "note": self.note,
        }

    def to_json(self) -> str:
        """docx4j ``ChartDescription.toJson``: two-space indents, one member a line, and a final line feed."""
        return json.dumps(self.to_dict(), indent=2, ensure_ascii=False) + "\n"

    def __str__(self) -> str:
        """The note."""
        return self.note
