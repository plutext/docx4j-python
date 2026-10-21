"""A chartex chart (waterfall, funnel, histogram, ...), for its table and description. CR-011 section 6.

docx4j's ``org.docx4j.model.chart.ChartExModel`` at ``198d015a3``. A chartex
part is an lxml tree in this engine (``ChartExSpacePart``, a
``DefaultXmlPart``: CR-001 section 14.12), so it is read by element name; the
tree is parsed without the part being modified.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

__all__ = ["ChartExModel", "java_double"]

CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _q(local: str) -> str:
    return f"{{{CX}}}{local}"


def _true(value: str | None) -> bool:
    return value in ("1", "true")


def java_double(d: float) -> str:
    """Java's ``String.valueOf(double)``: ``0.5``, ``12345.0``, ``1.2345E7``, ``1.0E-5``."""
    if math.isnan(d):
        return "NaN"
    if math.isinf(d):
        return "Infinity" if d > 0 else "-Infinity"
    if d == 0:
        return "-0.0" if math.copysign(1, d) < 0 else "0.0"
    if 1e-3 <= abs(d) < 1e7:
        text = f"{Decimal(repr(d)):f}"
        return text if "." in text else text + ".0"
    sign, digits, exponent = Decimal(repr(d)).normalize().as_tuple()
    figures = "".join(map(str, digits))
    scientific = len(figures) - 1 + int(exponent)
    return f"{'-' if sign else ''}{figures[0]}.{figures[1:] or '0'}E{scientific}"


def _rich(rich: Any) -> str:
    """Rich text's paragraphs, joined by line feeds; a run's line breaks are line feeds too."""
    lines = []
    for p in rich.iter(f"{{{_A}}}p"):
        out = []
        for child in p:
            if child.tag in (f"{{{_A}}}r", f"{{{_A}}}fld"):
                t = child.find(f"{{{_A}}}t")
                if t is not None and t.text is not None:
                    out.append(t.text)
            elif child.tag == f"{{{_A}}}br":
                out.append("\n")
        lines.append("".join(out))
    return "\n".join(lines)


def _text(tx: Any) -> tuple[str | None, str | None]:
    """The text and formula of a ``cx:tx``."""
    if tx is None:
        return None, None
    rich = tx.find(_q("rich"))
    if rich is not None:
        return _rich(rich), None
    value = formula = None
    data = tx.find(_q("txData"))
    if data is not None:
        for child in data:
            if child.tag == _q("f"):
                formula = child.text or ""
            elif child.tag == _q("v"):
                value = child.text or ""
    return value, formula


@dataclass(slots=True, eq=False)
class Dimension:
    """One dimension of a series' data: ``cx:strDim`` or ``cx:numDim``."""

    type: str | None
    numeric: bool
    formula: str | None = None
    counts: list[int] = field(default_factory=list)
    format_codes: list[str | None] = field(default_factory=list)
    levels: list[dict[int, str]] = field(default_factory=list)


@dataclass(slots=True, eq=False)
class Series:
    """A series (``cx:series``)."""

    layout: str | None
    hidden: bool
    name: str | None
    data: list[Dimension] = field(default_factory=list)


@dataclass(slots=True, eq=False)
class Axis:
    """An axis (``cx:axis``): category or value, hidden or not, and its title."""

    id: int
    kind: str | None
    hidden: bool
    has_title: bool
    title_text: str | None

    @property
    def shown_title(self) -> str | None:
        """The title Office shows: its text, or "Axis Title" where it states none; None for no title."""
        if not self.has_title:
            return None
        return self.title_text if self.title_text is not None else "Axis Title"


@dataclass(slots=True, eq=False)
class ChartExModel:
    """A chartex chart: its title, series and axes. docx4j ``ChartExModel``."""

    has_title: bool = False
    title_text: str | None = None
    series: list[Series] = field(default_factory=list)
    axes: list[Axis] = field(default_factory=list)

    @property
    def shown_title(self) -> str | None:
        """The title Office shows: its text, or "Chart Title" where it states none; None for no title."""
        if not self.has_title:
            return None
        return self.title_text if self.title_text is not None else "Chart Title"

    @classmethod
    def read(cls, source: Any) -> ChartExModel:
        """A chartex part (its tree, parsed without modifying it), or a ``cx:chartSpace`` element."""
        root = source.tree if hasattr(source, "tree") else source
        if root is None or root.tag != _q("chartSpace"):
            raise ValueError("not a chartex chart space")
        model = cls()
        data: dict[int, list[Dimension]] = {}
        chart_data = root.find(_q("chartData"))
        if chart_data is not None:
            for d in chart_data.findall(_q("data")):
                data[int(d.get("id", "0"))] = _dimensions(d)
        chart = root.find(_q("chart"))
        if chart is None:
            return model
        title = chart.find(_q("title"))
        if title is not None:
            model.has_title = True
            model.title_text = _text(title.find(_q("tx")))[0]
        plot_area = chart.find(_q("plotArea"))
        if plot_area is not None:
            region = plot_area.find(_q("plotAreaRegion"))
            for s in region.findall(_q("series")) if region is not None else ():
                series = Series(s.get("layoutId"), _true(s.get("hidden")), _text(s.find(_q("tx")))[0])
                data_id = s.find(_q("dataId"))
                if data_id is not None and data_id.get("val") is not None:
                    series.data = data.get(int(data_id.get("val")), [])
                model.series.append(series)
            for a in plot_area.findall(_q("axis")):
                kind = "category" if a.find(_q("catScaling")) is not None else "value" if a.find(_q("valScaling")) is not None else None
                axis_title = a.find(_q("title"))
                model.axes.append(
                    Axis(
                        int(a.get("id", "0")),
                        kind,
                        _true(a.get("hidden")),
                        axis_title is not None,
                        _text(axis_title.find(_q("tx")))[0] if axis_title is not None else None,
                    )
                )
        return model


def _dimensions(d: Any) -> list[Dimension]:
    out = []
    for child in d:
        if child.tag == _q("strDim"):
            dim = Dimension(child.get("type"), False)
        elif child.tag == _q("numDim"):
            dim = Dimension(child.get("type"), True)
        else:
            continue
        for item in child:
            if item.tag == _q("f"):
                dim.formula = item.text or ""
            elif item.tag == _q("lvl"):
                points: dict[int, str] = {}
                for pt in item.findall(_q("pt")):
                    text = pt.text or ""
                    if dim.numeric:
                        try:
                            x = float(text)
                        except ValueError:
                            continue
                        text = str(int(x)) if x == math.floor(x) and abs(x) < 1e15 else java_double(x)
                    points[int(pt.get("idx", "0"))] = text
                dim.levels.append(points)
                dim.counts.append(int(item.get("ptCount", "0")))
                dim.format_codes.append(item.get("formatCode") if dim.numeric else None)
        out.append(dim)
    return out
