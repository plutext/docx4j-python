from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child, ChildList
from docx4j_py.dml.chart import (
    CTBoolean as ChartCtboolean,
)
from docx4j_py.dml.chart import (
    CTDLbl as ChartCtdlbl,
)
from docx4j_py.dml.chart import (
    CTMarker as ChartCtmarker,
)
from docx4j_py.dml.chart import (
    CTUnsignedInt as ChartCtunsignedInt,
)
from docx4j_py.dml.main import CTShapeProperties as MainCtshapeProperties

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTBooleanFalse(Child):
    class Meta:
        name = "CT_BooleanFalse"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartUniqueID(Child):
    class Meta:
        name = "CT_ChartUniqueID"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLiteralDataChart(Child):
    class Meta:
        name = "CT_LiteralDataChart"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBoolean(ChartCtboolean):
    class Meta:
        name = "invertIfNegative"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTChartDataPointUniqueIDMapEntry(Child):
    class Meta:
        name = "CT_ChartDataPointUniqueIDMapEntry"

    ptidx: None | int = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    unique_id: None | CTChartUniqueID = field(
        default=None,
        metadata={
            "name": "uniqueID",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDLbl(ChartCtdlbl):
    class Meta:
        name = "dLbl"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTMarker(ChartCtmarker):
    class Meta:
        name = "marker"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTMultiLvlStrFilteredLiteralCache(Child):
    class Meta:
        name = "CT_MultiLvlStrFilteredLiteralCache"

    multi_lvl_str_cache: None | CTMultiLvlStrData = field(
        default=None,
        metadata={
            "name": "multiLvlStrCache",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumFilteredLiteralCache(Child):
    class Meta:
        name = "CT_NumFilteredLiteralCache"

    num_cache: None | CTNumData = field(
        default=None,
        metadata={
            "name": "numCache",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPivotOptions16(Child):
    class Meta:
        name = "CT_PivotOptions16"

    show_expand_collapse_field_buttons: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "showExpandCollapseFieldButtons",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeProperties(MainCtshapeProperties):
    class Meta:
        name = "spPr"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTStrFilteredLiteralCache(Child):
    class Meta:
        name = "CT_StrFilteredLiteralCache"

    str_cache: None | CTStrData = field(
        default=None,
        metadata={
            "name": "strCache",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUnsignedInt(ChartCtunsignedInt):
    class Meta:
        name = "explosion"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class Bubble3D(ChartCtboolean):
    class Meta:
        name = "bubble3D"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTCategoryFilterException(Child):
    class Meta:
        name = "CT_CategoryFilterException"

    unique_id: None | CTChartUniqueID = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    explosion: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    invert_if_negative: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "invertIfNegative",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    bubble3_d: None | Bubble3D = field(
        default=None,
        metadata={
            "name": "bubble3D",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )
    d_lbl: None | CTDLbl = field(
        default=None,
        metadata={
            "name": "dLbl",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartDataPointUniqueIDMap(Child):
    class Meta:
        name = "CT_ChartDataPointUniqueIDMap"

    ptentry: list[CTChartDataPointUniqueIDMapEntry] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class PivotOptions16(CTPivotOptions16):
    class Meta:
        name = "pivotOptions16"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CTCategoryFilterExceptions(Child):
    class Meta:
        name = "CT_CategoryFilterExceptions"

    category_filter_exception: list[CTCategoryFilterException] = field(
        default_factory=ChildList,
        metadata={
            "name": "categoryFilterException",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2014/chart",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class Datapointuniqueidmap(CTChartDataPointUniqueIDMap):
    class Meta:
        name = "datapointuniqueidmap"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


@dataclass(slots=True, kw_only=True)
class CategoryFilterExceptions(CTCategoryFilterExceptions):
    class Meta:
        name = "categoryFilterExceptions"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/chart"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.chart import (
    CTMultiLvlStrData,
    CTNumData,
    CTStrData,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.chart_2014 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.chart_2014.el",
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
