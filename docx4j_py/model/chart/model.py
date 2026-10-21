"""A chart part read as Office shows it, for its table and its description. CR-011 section 6.

docx4j's ``org.docx4j.model.chart`` reading classes at ``198d015a3`` ---
``ChartModel``, ``ChartGroup``, ``ChartSeries``, ``ChartData``,
``ChartAxis``, ``ChartTitle`` --- less what only drawing a chart needs (fills,
data labels, the legend, blanks, the theme). Read from a typed
``c:chartSpace`` (``docx4j_py.dml.chart.ChartSpace``), which a chart part
holds; :meth:`ChartModel.read` takes the part or its content.

Office's defaults are applied where the file states none, as docx4j applies
them: a series with no ``c:tx`` is "Series" and its index plus one ([MS-OI29500]
2.1.1565); a chart title with no text "Chart Title", or the single shown
series' name (2.1.1561); an axis title with no text "Axis Title"; a pie shows
its first series only; a series hidden with Excel's chart filter
(``c15:filtered...Series``) is not shown.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

__all__ = ["ChartAxis", "ChartData", "ChartGroup", "ChartModel", "ChartSeries", "ChartTitle", "root_tag"]

#: A chart part's root.
CHART_SPACE = "{http://schemas.openxmlformats.org/drawingml/2006/chart}chartSpace"


def root_tag(part: Any) -> str | None:
    """The qualified name of an XML part's root, read from its bytes as they would be saved; None for none."""
    import io

    from lxml import etree

    try:
        for _event, element in etree.iterparse(io.BytesIO(part.xml), events=("start",), resolve_entities=False):
            return element.tag
    except Exception:  # noqa: BLE001 - bytes that will not parse have no root
        return None
    return None


def _val(element: Any) -> Any:
    """An element's ``val``, an enumeration's value; None for no element."""
    if element is None:
        return None
    value = getattr(element, "val", None)
    return getattr(value, "value", value)


def _bool(element: Any) -> bool | None:
    """A ``CT_Boolean``: None for no element, true where its ``val`` is absent (the schema's default)."""
    if element is None:
        return None
    value = getattr(element, "val", None)
    return True if value is None else bool(value) and value not in ("0", "false")


def _str(element: Any) -> str | None:
    value = _val(element)
    if value is None:
        return None
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


# ---------------------------------------------------------------------------
# the data
# ---------------------------------------------------------------------------


@dataclass(slots=True, eq=False)
class ChartData:
    """A chart's data as the file holds it: a cache of points, by index. docx4j ``ChartData``."""

    #: ``NUMBER_REFERENCE``, ``NUMBER_LITERAL``, ``STRING_REFERENCE``,
    #: ``STRING_LITERAL``, ``MULTI_LEVEL_REFERENCE`` or ``TEXT`` (a series name as ``c:v``).
    kind: str
    #: ``c:f``: the cells, as written; None for a literal.
    reference: str | None = None
    #: ``c:ptCount``; None where the file has none.
    count: int | None = None
    #: The points' number format (``c:formatCode``); None for text.
    format_code: str | None = None
    #: Every level of labels, the innermost first: one for anything but multi-level categories.
    levels: list[dict[int, str]] = field(default_factory=list)
    point_formats: dict[int, str] = field(default_factory=dict)

    @property
    def numeric(self) -> bool:
        """Whether its points are numbers."""
        return self.kind in ("NUMBER_REFERENCE", "NUMBER_LITERAL")

    @property
    def points(self) -> dict[int, str]:
        """The points of the innermost level, by index; an index without a point is a gap."""
        return self.levels[0] if self.levels else {}

    def text(self, index: int) -> str | None:
        """A point's text; None for a gap."""
        return self.points.get(index)

    def format_code_of(self, index: int) -> str | None:
        """A point's own number format, where its ``c:pt`` gives one; else the data's."""
        return self.point_formats.get(index, self.format_code)

    def size(self) -> int:
        """How many points to show: ``c:ptCount``, or one past the highest index."""
        if self.count is not None:
            return self.count
        return max((i + 1 for level in self.levels for i in level), default=0)

    # -- reading ------------------------------------------------------------------

    @classmethod
    def of_axis(cls, source: Any) -> ChartData | None:
        """Categories or x values (``c:cat``, ``c:xVal``); None for none."""
        if source is None:
            return None
        return cls._of(getattr(source, "content", None))

    @classmethod
    def of_numbers(cls, source: Any) -> ChartData | None:
        """Values, y values or bubble sizes (``c:val``, ``c:yVal``, ``c:bubbleSize``); None for none."""
        if source is None:
            return None
        return cls._of(getattr(source, "num_ref_or_num_lit", None))

    @classmethod
    def of_name(cls, tx: Any) -> ChartData | None:
        """A series' name (``c:tx``); None for none."""
        if tx is None:
            return None
        value = getattr(tx, "str_ref_or_v", None)
        if value is None:
            return None
        if isinstance(value, str):
            return cls("TEXT", None, 1, None, [{0: value}])
        return cls._of(value)

    @classmethod
    def _of(cls, item: Any) -> ChartData | None:
        if item is None:
            return None
        name = type(item).__name__
        if name == "CTMultiLvlStrRef":
            cache = item.multi_lvl_str_cache
            levels: list[dict[int, str]] = []
            count = None
            if cache is not None:
                count = _val(cache.pt_count)
                levels = [_points(lvl.pt) for lvl in cache.lvl or ()]
            return cls("MULTI_LEVEL_REFERENCE", item.f, count, None, levels)
        if name == "CTNumRef":
            return cls._numbers("NUMBER_REFERENCE", item.f, item.num_cache)
        if name == "CTNumData":
            return cls._numbers("NUMBER_LITERAL", None, item)
        if name == "CTStrRef":
            return cls._strings("STRING_REFERENCE", item.f, item.str_cache)
        if name == "CTStrData":
            return cls._strings("STRING_LITERAL", None, item)
        return None

    @classmethod
    def _strings(cls, kind: str, f: str | None, data: Any) -> ChartData:
        count = _val(data.pt_count) if data is not None else None
        points = _points(data.pt) if data is not None else {}
        return cls(kind, f, count, None, [points])

    @classmethod
    def _numbers(cls, kind: str, f: str | None, data: Any) -> ChartData:
        count = None
        code = None
        points: dict[int, str] = {}
        formats: dict[int, str] = {}
        if data is not None:
            count = _val(data.pt_count)
            code = data.format_code
            for pt in data.pt or ():
                points[int(pt.idx)] = pt.v
                if getattr(pt, "format_code", None) is not None:
                    formats[int(pt.idx)] = pt.format_code
        return cls(kind, f, count, code, [points], formats)


def _points(pts: Any) -> dict[int, str]:
    return {int(pt.idx): pt.v for pt in pts or ()}


@dataclass(slots=True, eq=False)
class ChartTitle:
    """A chart's or an axis's title. docx4j ``ChartTitle``."""

    #: The title's text as the file states it (``c:tx``); None where it states none.
    text: str | None
    #: ``c:tx/c:strRef``: the cell the title comes from; None for rich text or none.
    reference: ChartData | None
    #: ``c:overlay``.
    overlay: bool
    #: The text Office shows: the title's own, or its default where it states none.
    shown_text: str

    @classmethod
    def of(cls, title: Any, fallback: str) -> ChartTitle | None:
        """A ``c:title``; `fallback` is what Office shows where it states no text."""
        if title is None:
            return None
        text = None
        reference = None
        tx = getattr(title, "tx", None)
        content = getattr(tx, "str_ref_or_rich", None) if tx is not None else None
        if content is not None:
            if type(content).__name__ == "CTStrRef":
                reference = ChartData._of(content)
                text = reference.text(0) if reference is not None else None
            else:
                text = rich_text(content)
        overlay = _bool(getattr(title, "overlay", None)) is True
        return cls(text, reference, overlay, text if text is not None else fallback)


def rich_text(body: Any) -> str:
    """Rich text's paragraphs, joined by line feeds; a run's line breaks are line feeds too."""
    from docx4j_py.model.text.drawings import text

    return "\n".join(text(p) for p in getattr(body, "p", None) or ())


@dataclass(slots=True, eq=False)
class ChartAxis:
    """An axis. docx4j ``ChartAxis``."""

    #: ``CATEGORY`` (``c:catAx``), ``VALUE`` (``c:valAx``), ``DATE`` (``c:dateAx``) or ``SERIES`` (``c:serAx``).
    kind: str
    #: ``c:axId``.
    id: int
    #: ``c:axPos``: b, l, r or t.
    position: str | None
    #: ``c:delete``: the axis is not drawn; false where not stated.
    deleted: bool
    #: The axis' title; None for none.
    title: ChartTitle | None
    #: ``c:crossAx``.
    cross_axis_id: int | None = None

    _KINDS = {"CTCatAx": "CATEGORY", "CTValAx": "VALUE", "CTDateAx": "DATE", "CTSerAx": "SERIES"}

    @classmethod
    def of(cls, axis: Any) -> ChartAxis | None:
        kind = cls._KINDS.get(type(axis).__name__)
        if kind is None:
            return None
        ax_id = _val(axis.ax_id)
        cross = _val(getattr(axis, "cross_ax", None))
        return cls(
            kind,
            0 if ax_id is None else int(ax_id),
            _str(axis.ax_pos),
            _bool(axis.delete) is True,
            ChartTitle.of(axis.title, "Axis Title"),
            None if cross is None else int(cross),
        )


@dataclass(slots=True, eq=False)
class ChartSeries:
    """A series. docx4j ``ChartSeries``."""

    #: ``c:idx``: numbers its automatic colour and default name.
    idx: int
    #: ``c:order``: its place in the plot.
    order: int
    #: ``c:tx``: the name as the file has it; None for none.
    name: ChartData | None
    #: The name Office shows: its own, or "Series" and its index plus one.
    shown_name: str
    categories: ChartData | None = None
    values: ChartData | None = None
    x_values: ChartData | None = None
    y_values: ChartData | None = None
    bubble_sizes: ChartData | None = None
    #: Hidden with Excel's chart filter (``c15:filtered...Series``).
    hidden: bool = False

    @classmethod
    def of(cls, ser: Any, hidden: bool) -> ChartSeries:
        idx = _val(getattr(ser, "idx", None))
        order = _val(getattr(ser, "order", None))
        index = 0 if idx is None else int(idx)
        name = ChartData.of_name(getattr(ser, "tx", None))
        named = name.text(0) if name is not None else None
        return cls(
            index,
            index if order is None else int(order),
            name,
            named if named is not None else f"Series{index + 1}",
            ChartData.of_axis(getattr(ser, "cat", None)),
            ChartData.of_numbers(getattr(ser, "val", None)),
            ChartData.of_axis(getattr(ser, "x_val", None)),
            ChartData.of_numbers(getattr(ser, "y_val", None)),
            ChartData.of_numbers(getattr(ser, "bubble_size", None)),
            hidden,
        )


#: A group's kind by its class (docx4j ``ChartGroup.Kind``): the element's name less "Chart".
_GROUP_KINDS = {
    "CTAreaChart": "area",
    "CTArea3DChart": "area3D",
    "CTBarChart": "bar",
    "CTBar3DChart": "bar3D",
    "CTBubbleChart": "bubble",
    "CTDoughnutChart": "doughnut",
    "CTLineChart": "line",
    "CTLine3DChart": "line3D",
    "CTOfPieChart": "ofPie",
    "CTPieChart": "pie",
    "CTPie3DChart": "pie3D",
    "CTRadarChart": "radar",
    "CTScatterChart": "scatter",
    "CTStockChart": "stock",
    "CTSurfaceChart": "surface",
    "CTSurface3DChart": "surface3D",
}


@dataclass(slots=True, eq=False)
class ChartGroup:
    """A chart group (``c:barChart``, ``c:lineChart``, ...). docx4j ``ChartGroup``."""

    #: The element's name less "Chart": bar, line3D, ofPie ...
    kind: str
    #: ``c:barDir``: bar or col; bar charts only.
    bar_direction: str | None
    #: ``c:grouping`` as stated.
    grouping: str | None
    #: The kind's settings the description needs: ``ofPieType``, ``radarStyle``, ``wireframe``.
    settings: dict[str, str]
    #: ``c:axId``: the axes it is drawn against.
    axis_ids: list[int]
    #: Every series, shown and filtered out, in their order.
    series: list[ChartSeries]

    @classmethod
    def of(cls, group: Any) -> ChartGroup | None:
        kind = _GROUP_KINDS.get(type(group).__name__)
        if kind is None:
            return None
        settings: dict[str, str] = {}
        for name, attr in (("ofPieType", "of_pie_type"), ("radarStyle", "radar_style"), ("wireframe", "wireframe")):
            element = getattr(group, attr, None)
            if element is None:
                continue
            value = ("true" if _bool(element) else "false") if attr == "wireframe" else _str(element)
            if value is not None:
                settings[name] = value
        series = [ChartSeries.of(s, False) for s in getattr(group, "ser", None) or ()]
        # the series filtered out with Excel's chart filter, in the group's extLst
        ext_lst = getattr(group, "ext_lst", None)
        for ext in getattr(ext_lst, "ext", None) or ():
            content = ext.any_element if isinstance(ext.any_element, list) else [ext.any_element]
            for item in content:
                name = type(item).__name__
                if item is None or not (name.startswith("Filtered") and name.endswith("Series")):
                    continue
                ser = getattr(item, "ser", None)
                for s in ser if isinstance(ser, list) else ([] if ser is None else [ser]):
                    series.append(ChartSeries.of(s, True))
        series.sort(key=lambda s: s.order)
        return cls(
            kind,
            _str(getattr(group, "bar_dir", None)),
            _str(getattr(group, "grouping", None)),
            settings,
            [int(_val(a)) for a in getattr(group, "ax_id", None) or () if _val(a) is not None],
            series,
        )

    @property
    def shown_series(self) -> list[ChartSeries]:
        """The series Office draws: not filtered out, and for a pie chart the first only."""
        out = []
        for s in self.series:
            if s.hidden:
                continue
            out.append(s)
            if self.kind in ("pie", "pie3D"):
                break
        return out


@dataclass(slots=True, eq=False)
class ChartModel:
    """A chart as Office shows it, for its table and description. docx4j ``ChartModel``."""

    #: ``c:date1904``.
    date1904: bool | None
    #: The chart's title; None where it has no ``c:title``.
    title: ChartTitle | None
    #: The chart groups, in document order: more than one in a combination chart.
    groups: list[ChartGroup]
    #: The axes, in document order.
    axes: list[ChartAxis]

    @classmethod
    def read(cls, source: Any) -> ChartModel:
        """A chart part, or a ``c:chartSpace``, read. The part is read with ``read_contents``.

        Raises:
            ValueError: the part's root is not a ``c:chartSpace`` (docx4j's
                unmarshal refuses it; this engine's lenient parser would not).
        """
        if hasattr(source, "read_contents"):
            if root_tag(source) != CHART_SPACE:
                raise ValueError(f"{source.part_name} is not a chart space")
            chart_space = source.read_contents()
        else:
            chart_space = source
        if chart_space is None:
            raise ValueError("no chart space")
        chart = getattr(chart_space, "chart", None)
        groups: list[ChartGroup] = []
        axes: list[ChartAxis] = []
        title = None
        if chart is not None:
            plot_area = chart.plot_area
            if plot_area is not None:
                for item in plot_area.content or ():
                    group = ChartGroup.of(item)
                    if group is not None:
                        groups.append(group)
                for item in plot_area.content_1 or ():
                    axis = ChartAxis.of(item)
                    if axis is not None:
                        axes.append(axis)
        model = cls(_bool(getattr(chart_space, "date1904", None)), None, groups, axes)
        if chart is not None:
            model.title = ChartTitle.of(chart.title, model._default_title())
        return model

    def _default_title(self) -> str:
        """[MS-OI29500] 2.1.1561 (a): "Chart Title", or the single shown series' name where it has one."""
        shown = [s for g in self.groups for s in g.shown_series]
        if len(shown) == 1 and shown[0].name is not None and shown[0].name.text(0) is not None:
            return shown[0].name.text(0)  # type: ignore[return-value]
        return "Chart Title"

    def axis(self, axis_id: int) -> ChartAxis | None:
        """The axis with an id; None for none."""
        return next((a for a in self.axes if a.id == axis_id), None)
