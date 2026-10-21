"""Charts put into words, ``docx4j_py.model.chart``. CR-011 Phase C (section 13).

docx4j's ``ChartDescriptionTest`` and ``ChartTableTest`` at ``198d015a3``,
their strings kept, over docx4j's chart fixtures (``tests/fixtures/charts/``):
each chart's description against docx4j's JSON golden (33 charts, chartex
among them), the kinds and axis roles, and the table's cells as Excel formats
them. A chart part is read straight from its package's zip, so a workbook's or
a presentation's chart is read without this package loading either.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from conformance import ROOT

from docx4j_py.model.chart import ChartDescription, ChartExModel, ChartModel, ChartTable
from docx4j_py.model.chart.description import layout_name
from docx4j_py.model.chart.formats import format_raw_cell_contents

CHARTS = ROOT / "tests" / "fixtures" / "charts"
GOLDENS = CHARTS / "description"


def _chart_parts(fixture: Path) -> list[str]:
    with zipfile.ZipFile(fixture) as z:
        return sorted(n for n in z.namelist() if "/charts/chart" in n and n.endswith(".xml"))


def _read(fixture: Path, part: str):
    """A chart part's model: a ChartModel, or a ChartExModel for chartex."""
    with zipfile.ZipFile(fixture) as z:
        data = z.read(part)
    if "chartEx" in part:
        from lxml import etree

        return ChartExModel.read(etree.fromstring(data))
    from docx4j_py.dml.chart import ChartSpace
    from docx4j_py.runtime import parser

    return ChartModel.read(parser(lenient=True, skipped_report=True).from_bytes(data, ChartSpace))


def describe(fixture: str, part: str | None = None) -> ChartDescription:
    path = CHARTS / fixture
    part = part or next(p for p in _chart_parts(path) if "chartEx" not in p)
    model = _read(path, part)
    return ChartDescription.of_chartex(model) if isinstance(model, ChartExModel) else ChartDescription.of(model)


def describe_ex(fixture: str) -> ChartDescription:
    path = CHARTS / fixture
    return ChartDescription.of_chartex(_read(path, next(p for p in _chart_parts(path) if "chartEx" in p)))


def table(fixture: str) -> ChartTable:
    return describe(fixture).table


def roles(d: ChartDescription) -> list[str]:
    return [f"{a.role}@{'null' if a.position is None else a.position}={'null' if a.title is None else a.title}" for a in d.axes]


def _goldens() -> list[tuple[str, str, Path]]:
    out = []
    for golden in sorted(GOLDENS.glob("*.json")):
        name = golden.name[: -len(".json")]
        for extension in (".docx_", ".xlsx_", ".pptx_"):
            at = name.find(extension)
            if at > 0:
                fixture = name[: at + len(extension) - 1]
                part = name[at + len(extension) :].replace("_", "/")
                folder = "chartex/" if fixture == "manual-bincount-6.xlsx" else ""
                out.append((folder + fixture, part, golden))
                break
    return out


def test_every_golden_is_read():
    assert len(_goldens()) == 33


@pytest.mark.parametrize(("fixture", "part", "golden"), _goldens(), ids=[g[2].name for g in _goldens()])
def test_the_description_is_docx4js(fixture, part, golden):
    """docx4j's ``jsonGoldens``: every chart part of the chart fixtures, chartex included."""
    assert describe(fixture, part).to_json() == golden.read_text(encoding="utf-8")


def test_bar_with_default_axis_titles():
    d = describe("bar-chart.xlsx")
    assert d.kind == "clustered bar"
    assert roles(d) == ["category axis@l=Axis Title", "value axis@b=Axis Title"]
    assert d.note == "[chart: clustered bar; category axis: Axis Title; value axis: Axis Title]"
    assert d.groups[0].element == "barChart" and d.groups[0].bar_direction == "bar"
    assert describe("bar-chart.docx").note == "[chart: clustered bar]"  # no c:title: none


def test_combo_on_secondary_axes():
    d = describe("p2-combo-secondary.xlsx")
    assert d.kind == "clustered column and line"
    assert not d.groups[0].secondary and d.groups[1].secondary
    assert d.secondary_series == ("Share",)
    assert roles(d) == ["category axis@b=null", "value axis@l=null", "secondary value axis@r=null"]
    assert d.note == "[chart: clustered column and line; secondary axis: Share]"


def test_axis_roles():
    assert roles(describe("p2-3d-column.xlsx")) == ["category axis@b=null", "value axis@l=null", "depth axis@b=null"]
    assert describe("p2-3d-column.xlsx").kind == "3-D column"
    assert describe("pie-chart.xlsx").axes == () and describe("pie-chart.xlsx").note == "[chart: pie]"
    assert roles(describe("p2-scatter.xlsx")) == ["x axis@b=null", "y axis@l=null"]
    assert roles(describe("p2-bubble.xlsx")) == ["x axis@b=null", "y axis@l=null"]
    assert roles(describe("p2-date-axis.xlsx"))[0] == "category axis@b=null"  # a date axis is a category axis


def test_chartex():
    d = describe_ex("waterfall.docx")
    assert d.note == "[chart: waterfall]"
    assert roles(d) == ["category axis@null=null", "value axis@null=null"]
    x = describe_ex("waterfall.xlsx")
    assert x.groups[0].series == ("Series 1",) and x.groups[0].element == "waterfall"
    assert [layout_name(n) for n in ("clusteredColumn", "boxWhisker", "regionMap", "funnel")] == [
        "histogram",
        "box and whisker",
        "map",
        "funnel",
    ]


def _pareto(axes: str) -> ChartExModel:
    from lxml import etree

    return ChartExModel.read(
        etree.fromstring(
            '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea>'
            '<cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:tx><cx:txData><cx:v>Scores</cx:v></cx:txData>'
            f'</cx:tx></cx:series><cx:series layoutId="paretoLine"/></cx:plotAreaRegion>{axes}'
            "</cx:plotArea></cx:chart></cx:chartSpace>"
        )
    )


def test_chartex_axis_titles_and_pareto():
    m = _pareto(
        '<cx:axis id="0"><cx:catScaling/><cx:title><cx:tx><cx:txData><cx:v>Bins</cx:v></cx:txData></cx:tx></cx:title></cx:axis>'
        '<cx:axis id="1"><cx:valScaling/><cx:title/></cx:axis>'
        '<cx:axis id="2"><cx:valScaling/></cx:axis>'
        '<cx:axis id="3" hidden="1"><cx:valScaling/></cx:axis>'
    )
    assert len(m.axes) == 4 and m.axes[0].title_text == "Bins"
    assert m.axes[1].shown_title == "Axis Title" and m.axes[2].shown_title is None and m.axes[3].hidden
    d = ChartDescription.of_chartex(m)
    assert d.kind == "pareto"
    assert roles(d) == ["category axis@null=Bins", "value axis@null=Axis Title", "secondary value axis@null=null"]
    assert d.note == "[chart: pareto; category axis: Bins; value axis: Axis Title]"
    d = ChartDescription.of_chartex(
        _pareto(
            '<cx:axis id="0"><cx:catScaling/></cx:axis><cx:axis id="1" hidden="1"><cx:valScaling/></cx:axis>'
            '<cx:axis id="2" hidden="1"><cx:valScaling/><cx:title><cx:tx><cx:txData><cx:v>Percent</cx:v></cx:txData>'
            "</cx:tx></cx:title></cx:axis>"
        )
    )
    assert roles(d) == ["category axis@null=null", "secondary value axis@null=Percent"]
    assert d.note == "[chart: pareto; secondary value axis: Percent]"


def test_a_deleted_axis_title_shows():
    d = describe("cr047-deleted-axis-titled.xlsx")
    assert roles(d) == ["category axis@l=CR047_SHOWN_AXIS_TITLE", "value axis@b=CR047_DELETED_AXIS_TITLE"]
    assert describe("cr047-deleted-axis-default-title.xlsx").note == (
        "[chart: clustered bar; category axis: CR047_SHOWN_AXIS_TITLE; value axis: Axis Title]"
    )


def test_a_hidden_chartex_axis_title_shows():
    d = describe_ex("cr047-pareto-hidden-axis-titled.xlsx")
    assert roles(d) == [
        "category axis@null=CR047_SHOWN_AXIS_TITLE",
        "value axis@null=CR047_HIDDEN_AXIS_TITLE",
        "secondary value axis@null=null",
    ]
    assert describe_ex("cr047-pareto-hidden-axis-default-title.xlsx").note == (
        "[chart: pareto; category axis: CR047_SHOWN_AXIS_TITLE; value axis: Axis Title; secondary value axis: Axis Title]"
    )
    d = describe_ex("cr047-pareto-excel-deleted.xlsx")
    assert roles(d) == ["category axis@null=null"] and d.note == "[chart: pareto]"


@pytest.mark.parametrize(
    ("fixture", "kind"),
    [
        ("clustered-column.xlsx", "clustered column"),
        ("stacked-column.xlsx", "stacked column"),
        ("line-chart.xlsx", "line"),
        ("p2-area.xlsx", "area"),
        ("p2-bar-of-pie.xlsx", "bar of pie"),
        ("p2-doughnut.xlsx", "doughnut"),
        ("p2-radar.xlsx", "radar"),
        ("p2-stock.xlsx", "stock"),
        ("p2-surface.xlsx", "3-D surface"),
        ("p2-scatter.xlsx", "scatter"),
        ("p2-bubble.xlsx", "bubble"),
    ],
)
def test_the_kinds_of_the_fixtures(fixture, kind):
    assert describe(fixture).kind == kind


@pytest.mark.parametrize(
    ("element", "kind"),
    [
        ('<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/></c:barChart>', "100% stacked bar"),
        ('<c:barChart><c:barDir val="col"/><c:grouping val="standard"/></c:barChart>', "column"),
        ('<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/></c:bar3DChart>', "3-D clustered column"),
        ('<c:lineChart><c:grouping val="stacked"/></c:lineChart>', "stacked line"),
        ('<c:line3DChart><c:grouping val="standard"/></c:line3DChart>', "3-D line"),
        ('<c:area3DChart><c:grouping val="percentStacked"/></c:area3DChart>', "3-D 100% stacked area"),
        ("<c:pie3DChart/>", "3-D pie"),
        ('<c:ofPieChart><c:ofPieType val="pie"/></c:ofPieChart>', "pie of pie"),
        ('<c:radarChart><c:radarStyle val="filled"/></c:radarChart>', "filled radar"),
        ('<c:surfaceChart><c:wireframe val="1"/></c:surfaceChart>', "wireframe contour"),
        ('<c:surface3DChart><c:wireframe val="1"/></c:surface3DChart>', "wireframe 3-D surface"),
        ('<c:surfaceChart><c:wireframe val="0"/></c:surfaceChart>', "contour"),
    ],
)
def test_the_kinds_of_the_elements(element, kind):
    from docx4j_py.dml.chart import ChartSpace
    from docx4j_py.runtime import parser

    xml = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>'
        f"{element}</c:plotArea></c:chart></c:chartSpace>"
    )
    model = ChartModel.read(parser(lenient=True, skipped_report=True).from_bytes(xml.encode(), ChartSpace))
    assert ChartDescription.of(model).kind == kind


# ---------------------------------------------------------------------------
# the table (docx4j's ChartTableTest)
# ---------------------------------------------------------------------------


def test_categories_down_series_across():
    t = table("clustered-column.xlsx")
    assert t.title == "Chart Title"
    assert t.header == ["", "Series 1", "Series 2", "Series 3"]
    assert t.rows[0] == ["Category 1", "4.3", "2.4", "2"]
    assert t.rows[1] == ["Category 2", "2.5", "4.4", "2"]  # 4.4000000000000004 in the cache
    assert len(t.rows) == 4


def test_gaps_and_errors():
    t = table("p2-gaps.xlsx")
    assert t.rows[1] == ["Category 2", "", "4.4", "2"]
    assert t.rows[2] == ["Category 3", "3.5", "#N/A", "3"]


def test_dates_formatted_by_their_format_codes():
    assert table("p2-date-axis.xlsx").rows[0] == ["1-Jan", "120"]
    stock = table("p2-stock.xlsx")
    assert stock.header == ["", "High", "Low", "Close"]
    assert stock.rows[0] == ["2026-01-01", "55", "11", "25"]


def test_multi_level_categories_a_column_each():
    t = table("p2-multilevel.xlsx")
    assert t.header == ["", "", "Sales"]
    assert t.rows[:3] == [["North", "Q1", "4.3"], ["", "Q2", "2.5"], ["South", "Q1", "3.5"]]


def test_scatter_shares_its_x():
    t = table("p2-scatter.xlsx")
    assert t.header == ["", "Y"] and t.rows[0] == ["0.7", "2.7"] and len(t.rows) == 5


def test_only_what_office_shows():
    f = table("p2-filtered.xlsx")
    assert f.header == ["", "Series 1", "Series 3"] and len(f.rows) == 3
    assert f.rows[2] == ["Category 4", "4.5", "5"]
    assert len(table("pie-chart.xlsx").header) == 2


def test_a_combination_chart_all_its_series():
    t = table("p2-combo-secondary.xlsx")
    assert t.header == ["", "Series 1", "Series 2", "Share"]
    assert t.rows[3] == ["Category 4", "4.5", "2.8", "0.62"]


def test_a_chartex_table():
    t = describe_ex("waterfall.docx").table
    assert t.title == "Chart Title" and t.header == ["", "Series1"]
    assert t.rows[3] == ["Category 4", "-40"] and len(t.rows) == 8


# ---------------------------------------------------------------------------
# the formats (a measured subset of POI's DataFormatter, CR-011 risk 5)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("value", "code", "text"),
    [
        (4.4000000000000004, "General", "4.4"),
        (3.5, "General", "3.5"),
        (0.62, "General", "0.62"),
        (1.5e11, "General", "1.5E+11"),
        (-2.5, "General", "-2.5"),
        (45000, "m/d/yy", "3/15/23"),  # docx4j's workbookCells: built-in 14
        (0.256, "0.00%", "25.60%"),
        (1234.5, "#,##0.00", "1,234.50"),
        (-1234.567, "$#,##0.00", "-$1,234.57"),
        (-5, "#,##0.00;(#,##0.00)", "(5.00)"),
        (12345, "0.00E+00", "1.23E+04"),
        (5, '#,##0.00 "EUR"', "5.00 EUR"),
        (45000, "d\\-mmm", "15-Mar"),
        (45000, "mmmmm", "M"),
        (0.5, "h:mm AM/PM", "12:00 PM"),
        (45000, "[$-409]mmmm d, yyyy;@", "March 15, 2023"),
        (2.5, "0", "3"),
        (2.25, "0.0", "2.3"),
    ],
)
def test_a_number_as_excel_shows_it(value, code, text):
    assert format_raw_cell_contents(value, 0 if code == "General" else -1, code) == text


def test_a_code_outside_the_subset_is_the_text():
    from docx4j_py.model.chart.table import format_number

    assert format_number("0.5", "# ?/?") == "0.5"
    assert format_number("1.5", "[h]:mm") == "1.5"
    assert format_number("#N/A", "General") == "#N/A"


def test_reading_a_workbooks_chart_changes_nothing():
    data = (CHARTS / "p2-combo-secondary.xlsx").read_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert "xl/charts/chart1.xml" in z.namelist()
    describe("p2-combo-secondary.xlsx")
    assert (CHARTS / "p2-combo-secondary.xlsx").read_bytes() == data
