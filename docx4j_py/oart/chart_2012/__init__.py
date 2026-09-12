from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child, ChildList
from docx4j_py.dml.chart import (
    CTBoolean as ChartCtboolean,
)
from docx4j_py.dml.chart import (
    CTChartLines as ChartCtchartLines,
)
from docx4j_py.dml.chart import (
    CTLayout as ChartCtlayout,
)
from docx4j_py.dml.chart import (
    CTNumFmt as ChartCtnumFmt,
)
from docx4j_py.dml.chart import (
    CTPivotSource as ChartCtpivotSource,
)
from docx4j_py.dml.chart import (
    CTTx as ChartCttx,
)
from docx4j_py.dml.main import CTShapeProperties as MainCtshapeProperties

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTFormulaRef(Child):
    class Meta:
        name = "CT_FormulaRef"

    sqref: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFullRef(Child):
    class Meta:
        name = "CT_FullRef"

    sqref: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLevelRef(Child):
    class Meta:
        name = "CT_LevelRef"

    sqref: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBoolean(ChartCtboolean):
    class Meta:
        name = "xForSave"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTCategoryFilterException(Child):
    class Meta:
        name = "CT_CategoryFilterException"

    sqref: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    sp_pr: None | MainCtshapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    explosion: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    invert_if_negative: None | ChartCtboolean = field(
        default=None,
        metadata={
            "name": "invertIfNegative",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    bubble3_d: None | ChartCtboolean = field(
        default=None,
        metadata={
            "name": "bubble3D",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    d_lbl: None | CTDLbl = field(
        default=None,
        metadata={
            "name": "dLbl",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartLines(ChartCtchartLines):
    class Meta:
        name = "leaderLines"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTDataLabelFieldTableEntry(Child):
    class Meta:
        name = "CT_DataLabelFieldTableEntry"

    txfld_guid: None | str = field(
        default=None,
        metadata={
            "name": "txfldGUID",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    f: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    dlbl_field_table_cache: None | CTStrData = field(
        default=None,
        metadata={
            "name": "dlblFieldTableCache",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredAreaSer(Child):
    class Meta:
        name = "CT_FilteredAreaSer"

    ser: None | CTAreaSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredBarSer(Child):
    class Meta:
        name = "CT_FilteredBarSer"

    ser: None | CTBarSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredBubbleSer(Child):
    class Meta:
        name = "CT_FilteredBubbleSer"

    ser: None | CTBubbleSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredCategoryTitle(Child):
    class Meta:
        name = "CT_FilteredCategoryTitle"

    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredLineSer(Child):
    class Meta:
        name = "CT_FilteredLineSer"

    ser: None | CTLineSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredPieSer(Child):
    class Meta:
        name = "CT_FilteredPieSer"

    ser: None | CTPieSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredRadarSer(Child):
    class Meta:
        name = "CT_FilteredRadarSer"

    ser: None | CTRadarSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredScatterSer(Child):
    class Meta:
        name = "CT_FilteredScatterSer"

    ser: None | CTScatterSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredSeriesTitle(Child):
    class Meta:
        name = "CT_FilteredSeriesTitle"

    tx: None | ChartCttx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFilteredSurfaceSer(Child):
    class Meta:
        name = "CT_FilteredSurfaceSer"

    ser: None | CTSurfaceSer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLayout(ChartCtlayout):
    class Meta:
        name = "layout"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTNumFmt(ChartCtnumFmt):
    class Meta:
        name = "numFmt"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTPivotSource(ChartCtpivotSource):
    class Meta:
        name = "pivotSource"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTSeriesDataLabelsRange(Child):
    class Meta:
        name = "CT_SeriesDataLabelsRange"

    f: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )
    dlbl_range_cache: None | CTStrData = field(
        default=None,
        metadata={
            "name": "dlblRangeCache",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeProperties(MainCtshapeProperties):
    class Meta:
        name = "spPr"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTTx(ChartCttx):
    class Meta:
        name = "tx"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class AutoCat(ChartCtboolean):
    class Meta:
        name = "autoCat"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FormulaRef(CTFormulaRef):
    class Meta:
        name = "formulaRef"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FullRef(CTFullRef):
    class Meta:
        name = "fullRef"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class LevelRef(CTLevelRef):
    class Meta:
        name = "levelRef"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class ShowDataLabelsRange(ChartCtboolean):
    class Meta:
        name = "showDataLabelsRange"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class ShowLeaderLines(ChartCtboolean):
    class Meta:
        name = "showLeaderLines"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CTCategoryFilterExceptions(Child):
    class Meta:
        name = "CT_CategoryFilterExceptions"

    category_filter_exception: list[CTCategoryFilterException] = field(
        default_factory=ChildList,
        metadata={
            "name": "categoryFilterException",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataLabelFieldTable(Child):
    class Meta:
        name = "CT_DataLabelFieldTable"

    dlbl_ftentry: list[CTDataLabelFieldTableEntry] = field(
        default_factory=ChildList,
        metadata={
            "name": "dlblFTEntry",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class DatalabelsRange(CTSeriesDataLabelsRange):
    class Meta:
        name = "datalabelsRange"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredAreaSeries(CTFilteredAreaSer):
    class Meta:
        name = "filteredAreaSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredBarSeries(CTFilteredBarSer):
    class Meta:
        name = "filteredBarSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredBubbleSeries(CTFilteredBubbleSer):
    class Meta:
        name = "filteredBubbleSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredCategoryTitle(CTFilteredCategoryTitle):
    class Meta:
        name = "filteredCategoryTitle"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredLineSeries(CTFilteredLineSer):
    class Meta:
        name = "filteredLineSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredPieSeries(CTFilteredPieSer):
    class Meta:
        name = "filteredPieSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredRadarSeries(CTFilteredRadarSer):
    class Meta:
        name = "filteredRadarSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredScatterSeries(CTFilteredScatterSer):
    class Meta:
        name = "filteredScatterSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredSeriesTitle(CTFilteredSeriesTitle):
    class Meta:
        name = "filteredSeriesTitle"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class FilteredSurfaceSeries(CTFilteredSurfaceSer):
    class Meta:
        name = "filteredSurfaceSeries"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class CategoryFilterExceptions(CTCategoryFilterExceptions):
    class Meta:
        name = "categoryFilterExceptions"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


@dataclass(slots=True, kw_only=True)
class DlblFieldTable(CTDataLabelFieldTable):
    class Meta:
        name = "dlblFieldTable"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chart"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.chart import (
    CTAreaSer,
    CTAxDataSource,
    CTBarSer,
    CTBubbleSer,
    CTDLbl,
    CTLineSer,
    CTMarker,
    CTPieSer,
    CTRadarSer,
    CTScatterSer,
    CTStrData,
    CTSurfaceSer,
    CTUnsignedInt,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.chart_2012 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.chart_2012.el",
    "element_name": "docx4j_py.traversal",
    "find": "docx4j_py.traversal",
    "iter_nodes": "docx4j_py.traversal",
    "link_parents": "docx4j_py.child",
    "text_of": "docx4j_py.traversal",
    "to_xml": "docx4j_py.fragments",
    "walk": "docx4j_py.traversal",
    "warm_up": "docx4j_py.runtime",
    "wml": "docx4j_py.fragments"
}


def __getattr__(name: str) -> object:
    """Import a Phase C helper, or the ``el`` submodule, on first use."""
    target = _PHASE_C.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    from importlib import import_module

    value = import_module(target) if name == "el" else getattr(import_module(target), name)
    globals()[name] = value
    return value
