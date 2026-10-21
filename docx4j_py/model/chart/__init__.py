"""Charts put into words: docx4j's ``org.docx4j.model.chart``, its reading half. CR-011 section 6.

What a reader without the picture needs of a chart (WD REQ-511, docx4j CR-035
and CR-047): :class:`ChartTable`, the chart's title and its data as a table,
each number formatted by its format code as Excel formats it; and
:class:`ChartDescription`, its kind in Excel's names, the axes Office shows
with their titles, and the series on secondary axes, with the note a text
reader writes before the title::

    from docx4j_py.model.chart import ChartDescription, ChartModel

    d = ChartDescription.of(ChartModel.read(chart_part))
    d.note           # '[chart: clustered column; value axis: Revenue]'
    d.table.rows     # [['North', '4.3', ...], ...]

:class:`ChartModel` reads a ``c:chartSpace`` and :class:`ChartExModel` a
chartex one (waterfall, funnel, histogram, ...), each without modifying its
part. Drawing a chart, editing one and docx4j's model JSON are not ported.

Imported on first use: ``import docx4j_py`` does not import it.
"""

from __future__ import annotations

from docx4j_py.model.chart.chartex import ChartExModel
from docx4j_py.model.chart.description import ChartDescription
from docx4j_py.model.chart.model import ChartAxis, ChartData, ChartGroup, ChartModel, ChartSeries, ChartTitle
from docx4j_py.model.chart.table import ChartTable

__all__ = [
    "ChartAxis",
    "ChartData",
    "ChartDescription",
    "ChartExModel",
    "ChartGroup",
    "ChartModel",
    "ChartSeries",
    "ChartTable",
    "ChartTitle",
]
