from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList
from docx4j_py.dml.chart_drawing import CTDrawing as ChartDrawingCtdrawing

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/chart"


@dataclass(slots=True, kw_only=True)
class CTAxisUnit(Child):
    class Meta:
        name = "CT_AxisUnit"

    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_exclusive": 0.0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBoolean(Child):
    class Meta:
        name = "CT_Boolean"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBubbleScale(Child):
    class Meta:
        name = "CT_BubbleScale"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100",
            "min_inclusive": 0,
            "max_inclusive": 300,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDepthPercent(Child):
    class Meta:
        name = "CT_DepthPercent"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100",
            "min_inclusive": 20,
            "max_inclusive": 2000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDouble(Child):
    class Meta:
        name = "CT_Double"

    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExtension(Child):
    class Meta:
        name = "CT_Extension"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )
    uri: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFirstSliceAng(Child):
    class Meta:
        name = "CT_FirstSliceAng"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 360,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGapAmount(Child):
    class Meta:
        name = "CT_GapAmount"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "150",
            "min_inclusive": 0,
            "max_inclusive": 500,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHPercent(Child):
    class Meta:
        name = "CT_HPercent"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100",
            "min_inclusive": 5,
            "max_inclusive": 500,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHeaderFooter(Child):
    class Meta:
        name = "CT_HeaderFooter"

    odd_header: None | str = field(
        default=None,
        metadata={
            "name": "oddHeader",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    odd_footer: None | str = field(
        default=None,
        metadata={
            "name": "oddFooter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    even_header: None | str = field(
        default=None,
        metadata={
            "name": "evenHeader",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    even_footer: None | str = field(
        default=None,
        metadata={
            "name": "evenFooter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    first_header: None | str = field(
        default=None,
        metadata={
            "name": "firstHeader",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    first_footer: None | str = field(
        default=None,
        metadata={
            "name": "firstFooter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    align_with_margins: None | bool = field(
        default=None,
        metadata={
            "name": "alignWithMargins",
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    different_odd_even: None | bool = field(
        default=None,
        metadata={
            "name": "differentOddEven",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    different_first: None | bool = field(
        default=None,
        metadata={
            "name": "differentFirst",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHoleSize(Child):
    class Meta:
        name = "CT_HoleSize"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "10",
            "min_inclusive": 10,
            "max_inclusive": 90,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLblOffset(Child):
    class Meta:
        name = "CT_LblOffset"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100",
            "min_inclusive": 0,
            "max_inclusive": 1000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLogBase(Child):
    class Meta:
        name = "CT_LogBase"

    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 2.0,
            "max_inclusive": 1000.0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarkerSize(Child):
    class Meta:
        name = "CT_MarkerSize"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "5",
            "min_inclusive": 2,
            "max_inclusive": 72,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumFmt(Child):
    class Meta:
        name = "CT_NumFmt"

    format_code: None | str = field(
        default=None,
        metadata={
            "name": "formatCode",
            "type": "Attribute",
        },
    )
    source_linked: None | bool = field(
        default=None,
        metadata={
            "name": "sourceLinked",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumVal(Child):
    class Meta:
        name = "CT_NumVal"

    v: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    format_code: None | str = field(
        default=None,
        metadata={
            "name": "formatCode",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOrder(Child):
    class Meta:
        name = "CT_Order"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "2",
            "min_inclusive": 2,
            "max_inclusive": 6,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOverlap(Child):
    class Meta:
        name = "CT_Overlap"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPageMargins(Child):
    class Meta:
        name = "CT_PageMargins"

    l: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    r: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    t: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    b: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    header: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    footer: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPeriod(Child):
    class Meta:
        name = "CT_Period"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "2",
            "min_inclusive": 2,
            "max_inclusive": 255,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPerspective(Child):
    class Meta:
        name = "CT_Perspective"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "30",
            "min_inclusive": 0,
            "max_inclusive": 240,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureStackUnit(Child):
    class Meta:
        name = "CT_PictureStackUnit"

    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_exclusive": 0.0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelId(Child):
    class Meta:
        name = "CT_RelId"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRotX(Child):
    class Meta:
        name = "CT_RotX"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -90,
            "max_inclusive": 90,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRotY(Child):
    class Meta:
        name = "CT_RotY"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 360,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSecondPieSize(Child):
    class Meta:
        name = "CT_SecondPieSize"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "75",
            "min_inclusive": 5,
            "max_inclusive": 200,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSkip(Child):
    class Meta:
        name = "CT_Skip"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStrVal(Child):
    class Meta:
        name = "CT_StrVal"

    v: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyle(Child):
    class Meta:
        name = "CT_Style"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 1,
            "max_inclusive": 48,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextLanguageID(Child):
    class Meta:
        name = "CT_TextLanguageID"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUnsignedInt(Child):
    class Meta:
        name = "CT_UnsignedInt"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class STAxPos(Enum):
    B = "b"
    L = "l"
    R = "r"
    T = "t"


class STBarDir(Enum):
    BAR = "bar"
    COL = "col"


class STBarGrouping(Enum):
    PERCENT_STACKED = "percentStacked"
    CLUSTERED = "clustered"
    STANDARD = "standard"
    STACKED = "stacked"


class STBuiltInUnit(Enum):
    HUNDREDS = "hundreds"
    THOUSANDS = "thousands"
    TEN_THOUSANDS = "tenThousands"
    HUNDRED_THOUSANDS = "hundredThousands"
    MILLIONS = "millions"
    TEN_MILLIONS = "tenMillions"
    HUNDRED_MILLIONS = "hundredMillions"
    BILLIONS = "billions"
    TRILLIONS = "trillions"


class STCrossBetween(Enum):
    BETWEEN = "between"
    MID_CAT = "midCat"


class STCrosses(Enum):
    AUTO_ZERO = "autoZero"
    MAX = "max"
    MIN = "min"


class STDLblPos(Enum):
    BEST_FIT = "bestFit"
    B = "b"
    CTR = "ctr"
    IN_BASE = "inBase"
    IN_END = "inEnd"
    L = "l"
    OUT_END = "outEnd"
    R = "r"
    T = "t"


class STDispBlanksAs(Enum):
    SPAN = "span"
    GAP = "gap"
    ZERO = "zero"


class STErrBarType(Enum):
    BOTH = "both"
    MINUS = "minus"
    PLUS = "plus"


class STErrDir(Enum):
    X = "x"
    Y = "y"


class STErrValType(Enum):
    CUST = "cust"
    FIXED_VAL = "fixedVal"
    PERCENTAGE = "percentage"
    STD_DEV = "stdDev"
    STD_ERR = "stdErr"


class STGrouping(Enum):
    PERCENT_STACKED = "percentStacked"
    STANDARD = "standard"
    STACKED = "stacked"


class STLayoutMode(Enum):
    EDGE = "edge"
    FACTOR = "factor"


class STLayoutTarget(Enum):
    INNER = "inner"
    OUTER = "outer"


class STLblAlgn(Enum):
    CTR = "ctr"
    L = "l"
    R = "r"


class STLegendPos(Enum):
    B = "b"
    TR = "tr"
    L = "l"
    R = "r"
    T = "t"


class STMarkerStyle(Enum):
    CIRCLE = "circle"
    DASH = "dash"
    DIAMOND = "diamond"
    DOT = "dot"
    NONE = "none"
    PICTURE = "picture"
    PLUS = "plus"
    SQUARE = "square"
    STAR = "star"
    TRIANGLE = "triangle"
    X = "x"


class STOfPieType(Enum):
    PIE = "pie"
    BAR = "bar"


class STOrientation(Enum):
    MAX_MIN = "maxMin"
    MIN_MAX = "minMax"


class STPageSetupOrientation(Enum):
    DEFAULT = "default"
    PORTRAIT = "portrait"
    LANDSCAPE = "landscape"


class STPictureFormat(Enum):
    STRETCH = "stretch"
    STACK = "stack"
    STACK_SCALE = "stackScale"


class STRadarStyle(Enum):
    STANDARD = "standard"
    MARKER = "marker"
    FILLED = "filled"


class STScatterStyle(Enum):
    NONE = "none"
    LINE = "line"
    LINE_MARKER = "lineMarker"
    MARKER = "marker"
    SMOOTH = "smooth"
    SMOOTH_MARKER = "smoothMarker"


class STShape(Enum):
    CONE = "cone"
    CONE_TO_MAX = "coneToMax"
    BOX = "box"
    CYLINDER = "cylinder"
    PYRAMID = "pyramid"
    PYRAMID_TO_MAX = "pyramidToMax"


class STSizeRepresents(Enum):
    AREA = "area"
    W = "w"


class STSplitType(Enum):
    AUTO = "auto"
    CUST = "cust"
    PERCENT = "percent"
    POS = "pos"
    VAL = "val"


class STTickLblPos(Enum):
    HIGH = "high"
    LOW = "low"
    NEXT_TO = "nextTo"
    NONE = "none"


class STTickMark(Enum):
    CROSS = "cross"
    IN = "in"
    NONE = "none"
    OUT = "out"


class STTimeUnit(Enum):
    DAYS = "days"
    MONTHS = "months"
    YEARS = "years"


class STTrendlineType(Enum):
    EXP = "exp"
    LINEAR = "linear"
    LOG = "log"
    MOVING_AVG = "movingAvg"
    POLY = "poly"
    POWER = "power"


@dataclass(slots=True, kw_only=True)
class CTAxPos(Child):
    class Meta:
        name = "CT_AxPos"

    val: None | STAxPos = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBandFmt(Child):
    class Meta:
        name = "CT_BandFmt"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBarDir(Child):
    class Meta:
        name = "CT_BarDir"

    val: None | STBarDir = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STBarDir.COL,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBarGrouping(Child):
    class Meta:
        name = "CT_BarGrouping"

    val: None | STBarGrouping = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STBarGrouping.CLUSTERED,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBuiltInUnit(Child):
    class Meta:
        name = "CT_BuiltInUnit"

    val: None | STBuiltInUnit = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STBuiltInUnit.THOUSANDS,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartLines(Child):
    class Meta:
        name = "CT_ChartLines"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCrossBetween(Child):
    class Meta:
        name = "CT_CrossBetween"

    val: None | STCrossBetween = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCrosses(Child):
    class Meta:
        name = "CT_Crosses"

    val: None | STCrosses = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustSplit(Child):
    class Meta:
        name = "CT_CustSplit"

    second_pie_pt: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "secondPiePt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDLblPos(Child):
    class Meta:
        name = "CT_DLblPos"

    val: None | STDLblPos = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDispBlanksAs(Child):
    class Meta:
        name = "CT_DispBlanksAs"

    val: None | STDispBlanksAs = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STDispBlanksAs.ZERO,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDrawing(ChartDrawingCtdrawing):
    class Meta:
        name = "userShapes"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"


@dataclass(slots=True, kw_only=True)
class CTErrBarType(Child):
    class Meta:
        name = "CT_ErrBarType"

    val: None | STErrBarType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STErrBarType.BOTH,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTErrDir(Child):
    class Meta:
        name = "CT_ErrDir"

    val: None | STErrDir = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTErrValType(Child):
    class Meta:
        name = "CT_ErrValType"

    val: None | STErrValType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STErrValType.FIXED_VAL,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExtensionList(Child):
    class Meta:
        name = "CT_ExtensionList"

    ext: list[CTExtension] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExternalData(Child):
    class Meta:
        name = "CT_ExternalData"

    auto_update: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "autoUpdate",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGrouping(Child):
    class Meta:
        name = "CT_Grouping"

    val: None | STGrouping = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STGrouping.STANDARD,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLayoutMode(Child):
    class Meta:
        name = "CT_LayoutMode"

    val: None | STLayoutMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STLayoutMode.FACTOR,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLayoutTarget(Child):
    class Meta:
        name = "CT_LayoutTarget"

    val: None | STLayoutTarget = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STLayoutTarget.OUTER,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLblAlgn(Child):
    class Meta:
        name = "CT_LblAlgn"

    val: None | STLblAlgn = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLegendPos(Child):
    class Meta:
        name = "CT_LegendPos"

    val: None | STLegendPos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STLegendPos.R,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLvl(Child):
    class Meta:
        name = "CT_Lvl"

    pt: list[CTStrVal] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarkerStyle(Child):
    class Meta:
        name = "CT_MarkerStyle"

    val: None | STMarkerStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOfPieType(Child):
    class Meta:
        name = "CT_OfPieType"

    val: None | STOfPieType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STOfPieType.PIE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOrientation(Child):
    class Meta:
        name = "CT_Orientation"

    val: None | STOrientation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STOrientation.MIN_MAX,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPageSetup(Child):
    class Meta:
        name = "CT_PageSetup"

    paper_size: None | int = field(
        default=None,
        metadata={
            "name": "paperSize",
            "type": "Attribute",
            "schema_default": "1",
        },
    )
    first_page_number: None | int = field(
        default=None,
        metadata={
            "name": "firstPageNumber",
            "type": "Attribute",
            "schema_default": "1",
        },
    )
    orientation: None | STPageSetupOrientation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STPageSetupOrientation.DEFAULT,
        },
    )
    black_and_white: None | bool = field(
        default=None,
        metadata={
            "name": "blackAndWhite",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    draft: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    use_first_page_number: None | bool = field(
        default=None,
        metadata={
            "name": "useFirstPageNumber",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    horizontal_dpi: None | int = field(
        default=None,
        metadata={
            "name": "horizontalDpi",
            "type": "Attribute",
            "schema_default": "600",
        },
    )
    vertical_dpi: None | int = field(
        default=None,
        metadata={
            "name": "verticalDpi",
            "type": "Attribute",
            "schema_default": "600",
        },
    )
    copies: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "1",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureFormat(Child):
    class Meta:
        name = "CT_PictureFormat"

    val: None | STPictureFormat = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTProtection(Child):
    class Meta:
        name = "CT_Protection"

    chart_object: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "chartObject",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    data: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    formatting: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    selection: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    user_interface: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "userInterface",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRadarStyle(Child):
    class Meta:
        name = "CT_RadarStyle"

    val: None | STRadarStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STRadarStyle.STANDARD,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScatterStyle(Child):
    class Meta:
        name = "CT_ScatterStyle"

    val: None | STScatterStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STScatterStyle.MARKER,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShape(Child):
    class Meta:
        name = "CT_Shape"

    val: None | STShape = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STShape.BOX,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSizeRepresents(Child):
    class Meta:
        name = "CT_SizeRepresents"

    val: None | STSizeRepresents = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STSizeRepresents.AREA,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSplitType(Child):
    class Meta:
        name = "CT_SplitType"

    val: None | STSplitType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STSplitType.AUTO,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTickLblPos(Child):
    class Meta:
        name = "CT_TickLblPos"

    val: None | STTickLblPos = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTickLblPos.NEXT_TO,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTickMark(Child):
    class Meta:
        name = "CT_TickMark"

    val: None | STTickMark = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTickMark.CROSS,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTimeUnit(Child):
    class Meta:
        name = "CT_TimeUnit"

    val: None | STTimeUnit = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTimeUnit.DAYS,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrendlineType(Child):
    class Meta:
        name = "CT_TrendlineType"

    val: None | STTrendlineType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTrendlineType.LINEAR,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUpDownBar(Child):
    class Meta:
        name = "CT_UpDownBar"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class Chart(CTRelId):
    class Meta:
        name = "chart"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"


@dataclass(slots=True, kw_only=True)
class Delete(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowBubbleSize(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowCatName(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowLeaderLines(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowLegendKey(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowPercent(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowSerName(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ShowVal(CTBoolean):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTBandFmts(Child):
    class Meta:
        name = "CT_BandFmts"

    band_fmt: list[CTBandFmt] = field(
        default_factory=ChildList,
        metadata={
            "name": "bandFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDTable(Child):
    class Meta:
        name = "CT_DTable"

    show_horz_border: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "showHorzBorder",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    show_vert_border: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "showVertBorder",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    show_outline: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "showOutline",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    show_keys: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "showKeys",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLegendEntry(Child):
    class Meta:
        name = "CT_LegendEntry"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    delete_or_tx_pr: None | CTBoolean | CTTextBody = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "delete",
                    "type": ForwardRef("CTBoolean"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "txPr",
                    "type": ForwardRef("CTTextBody"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTManualLayout(Child):
    class Meta:
        name = "CT_ManualLayout"

    layout_target: None | CTLayoutTarget = field(
        default=None,
        metadata={
            "name": "layoutTarget",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    x_mode: None | CTLayoutMode = field(
        default=None,
        metadata={
            "name": "xMode",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    y_mode: None | CTLayoutMode = field(
        default=None,
        metadata={
            "name": "yMode",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    w_mode: None | CTLayoutMode = field(
        default=None,
        metadata={
            "name": "wMode",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    h_mode: None | CTLayoutMode = field(
        default=None,
        metadata={
            "name": "hMode",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    x: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    y: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    w: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    h: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarker(Child):
    class Meta:
        name = "CT_Marker"

    symbol: None | CTMarkerStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    size: None | CTMarkerSize = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMultiLvlStrData(Child):
    class Meta:
        name = "CT_MultiLvlStrData"

    pt_count: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "ptCount",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    lvl: list[CTLvl] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumData(Child):
    class Meta:
        name = "CT_NumData"

    format_code: None | str = field(
        default=None,
        metadata={
            "name": "formatCode",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    pt_count: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "ptCount",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    pt: list[CTNumVal] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureOptions(Child):
    class Meta:
        name = "CT_PictureOptions"

    apply_to_front: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "applyToFront",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    apply_to_sides: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "applyToSides",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    apply_to_end: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "applyToEnd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    picture_format: None | CTPictureFormat = field(
        default=None,
        metadata={
            "name": "pictureFormat",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    picture_stack_unit: None | CTPictureStackUnit = field(
        default=None,
        metadata={
            "name": "pictureStackUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPivotSource(Child):
    class Meta:
        name = "CT_PivotSource"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    fmt_id: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "fmtId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: list[CTExtensionList] = field(
        default_factory=ChildList,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPrintSettings(Child):
    class Meta:
        name = "CT_PrintSettings"

    header_footer: None | CTHeaderFooter = field(
        default=None,
        metadata={
            "name": "headerFooter",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    page_margins: None | CTPageMargins = field(
        default=None,
        metadata={
            "name": "pageMargins",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    page_setup: None | CTPageSetup = field(
        default=None,
        metadata={
            "name": "pageSetup",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    legacy_drawing_hf: None | CTRelId = field(
        default=None,
        metadata={
            "name": "legacyDrawingHF",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScaling(Child):
    class Meta:
        name = "CT_Scaling"

    log_base: None | CTLogBase = field(
        default=None,
        metadata={
            "name": "logBase",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    orientation: None | CTOrientation = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    max: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    min: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStrData(Child):
    class Meta:
        name = "CT_StrData"

    pt_count: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "ptCount",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    pt: list[CTStrVal] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUpDownBars(Child):
    class Meta:
        name = "CT_UpDownBars"

    gap_width: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    up_bars: None | CTUpDownBar = field(
        default=None,
        metadata={
            "name": "upBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    down_bars: None | CTUpDownBar = field(
        default=None,
        metadata={
            "name": "downBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTView3D(Child):
    class Meta:
        name = "CT_View3D"

    rot_x: None | CTRotX = field(
        default=None,
        metadata={
            "name": "rotX",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    h_percent: None | CTHPercent = field(
        default=None,
        metadata={
            "name": "hPercent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    rot_y: None | CTRotY = field(
        default=None,
        metadata={
            "name": "rotY",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    depth_percent: None | CTDepthPercent = field(
        default=None,
        metadata={
            "name": "depthPercent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    r_ang_ax: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "rAngAx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    perspective: None | CTPerspective = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDPt(Child):
    class Meta:
        name = "CT_DPt"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    invert_if_negative: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "invertIfNegative",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    bubble3_d: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "bubble3D",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    explosion: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    picture_options: None | CTPictureOptions = field(
        default=None,
        metadata={
            "name": "pictureOptions",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLayout(Child):
    class Meta:
        name = "CT_Layout"

    manual_layout: None | CTManualLayout = field(
        default=None,
        metadata={
            "name": "manualLayout",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMultiLvlStrRef(Child):
    class Meta:
        name = "CT_MultiLvlStrRef"

    f: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    multi_lvl_str_cache: None | CTMultiLvlStrData = field(
        default=None,
        metadata={
            "name": "multiLvlStrCache",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumRef(Child):
    class Meta:
        name = "CT_NumRef"

    f: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    num_cache: None | CTNumData = field(
        default=None,
        metadata={
            "name": "numCache",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStrRef(Child):
    class Meta:
        name = "CT_StrRef"

    f: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    str_cache: None | CTStrData = field(
        default=None,
        metadata={
            "name": "strCache",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSurface(Child):
    class Meta:
        name = "CT_Surface"

    thickness: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    picture_options: None | CTPictureOptions = field(
        default=None,
        metadata={
            "name": "pictureOptions",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAxDataSource(Child):
    class Meta:
        name = "CT_AxDataSource"

    content: None | CTMultiLvlStrRef | CTNumRef | CTNumData | CTStrRef | CTStrData = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "multiLvlStrRef",
                    "type": ForwardRef("CTMultiLvlStrRef"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "numRef",
                    "type": ForwardRef("CTNumRef"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "numLit",
                    "type": ForwardRef("CTNumData"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "strRef",
                    "type": ForwardRef("CTStrRef"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "strLit",
                    "type": ForwardRef("CTStrData"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLegend(Child):
    class Meta:
        name = "CT_Legend"

    legend_pos: None | CTLegendPos = field(
        default=None,
        metadata={
            "name": "legendPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    legend_entry: list[CTLegendEntry] = field(
        default_factory=ChildList,
        metadata={
            "name": "legendEntry",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    layout: None | CTLayout = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    overlay: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumDataSource(Child):
    class Meta:
        name = "CT_NumDataSource"

    num_ref_or_num_lit: None | CTNumRef | CTNumData = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "numRef",
                    "type": ForwardRef("CTNumRef"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "numLit",
                    "type": ForwardRef("CTNumData"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSerTx(Child):
    class Meta:
        name = "CT_SerTx"

    str_ref_or_v: None | CTStrRef | str = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "strRef",
                    "type": ForwardRef("CTStrRef"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "v",
                    "type": str,
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTx(Child):
    class Meta:
        name = "CT_Tx"

    str_ref_or_rich: None | CTStrRef | CTTextBody = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "strRef",
                    "type": ForwardRef("CTStrRef"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "rich",
                    "type": ForwardRef("CTTextBody"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDLbl(Child):
    class Meta:
        name = "CT_DLbl"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    content: list[
        Delete
        | CTLayout
        | CTTx
        | CTNumFmt
        | CTShapeProperties
        | CTTextBody
        | CTDLblPos
        | ShowLegendKey
        | ShowVal
        | ShowCatName
        | ShowSerName
        | ShowPercent
        | ShowBubbleSize
        | str
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "delete",
                    "type": ForwardRef("Delete"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "layout",
                    "type": ForwardRef("CTLayout"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "tx",
                    "type": ForwardRef("CTTx"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "numFmt",
                    "type": ForwardRef("CTNumFmt"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "spPr",
                    "type": ForwardRef("CTShapeProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "txPr",
                    "type": ForwardRef("CTTextBody"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "dLblPos",
                    "type": ForwardRef("CTDLblPos"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showLegendKey",
                    "type": ForwardRef("ShowLegendKey"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showVal",
                    "type": ForwardRef("ShowVal"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showCatName",
                    "type": ForwardRef("ShowCatName"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showSerName",
                    "type": ForwardRef("ShowSerName"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showPercent",
                    "type": ForwardRef("ShowPercent"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showBubbleSize",
                    "type": ForwardRef("ShowBubbleSize"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "separator",
                    "type": str,
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
            "max_occurs": 13,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDispUnitsLbl(Child):
    class Meta:
        name = "CT_DispUnitsLbl"

    layout: None | CTLayout = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTErrBars(Child):
    class Meta:
        name = "CT_ErrBars"

    err_dir: None | CTErrDir = field(
        default=None,
        metadata={
            "name": "errDir",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_bar_type: None | CTErrBarType = field(
        default=None,
        metadata={
            "name": "errBarType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_val_type: None | CTErrValType = field(
        default=None,
        metadata={
            "name": "errValType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    no_end_cap: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "noEndCap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    plus: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minus: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSurfaceSer(Child):
    class Meta:
        name = "CT_SurfaceSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTitle(Child):
    class Meta:
        name = "CT_Title"

    tx: None | CTTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    layout: None | CTLayout = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    overlay: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrendlineLbl(Child):
    class Meta:
        name = "CT_TrendlineLbl"

    layout: None | CTLayout = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    num_fmt: None | CTNumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCatAx(Child):
    class Meta:
        name = "CT_CatAx"

    ax_id: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    scaling: None | CTScaling = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    delete: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_pos: None | CTAxPos = field(
        default=None,
        metadata={
            "name": "axPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "majorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "minorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    title: None | CTTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    num_fmt: None | CTNumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "majorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "minorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_lbl_pos: None | CTTickLblPos = field(
        default=None,
        metadata={
            "name": "tickLblPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cross_ax: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "crossAx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    crosses_or_crosses_at: None | CTCrosses | CTDouble = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "crosses",
                    "type": ForwardRef("CTCrosses"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "crossesAt",
                    "type": ForwardRef("CTDouble"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    auto: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    lbl_algn: None | CTLblAlgn = field(
        default=None,
        metadata={
            "name": "lblAlgn",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    lbl_offset: None | CTLblOffset = field(
        default=None,
        metadata={
            "name": "lblOffset",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_lbl_skip: None | CTSkip = field(
        default=None,
        metadata={
            "name": "tickLblSkip",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_mark_skip: None | CTSkip = field(
        default=None,
        metadata={
            "name": "tickMarkSkip",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    no_multi_lvl_lbl: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "noMultiLvlLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDLbls(Child):
    class Meta:
        name = "CT_DLbls"

    d_lbl: list[CTDLbl] = field(
        default_factory=ChildList,
        metadata={
            "name": "dLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    content: list[
        Delete
        | CTNumFmt
        | CTShapeProperties
        | CTTextBody
        | CTDLblPos
        | ShowLegendKey
        | ShowVal
        | ShowCatName
        | ShowSerName
        | ShowPercent
        | ShowBubbleSize
        | str
        | ShowLeaderLines
        | CTChartLines
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "delete",
                    "type": ForwardRef("Delete"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "numFmt",
                    "type": ForwardRef("CTNumFmt"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "spPr",
                    "type": ForwardRef("CTShapeProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "txPr",
                    "type": ForwardRef("CTTextBody"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "dLblPos",
                    "type": ForwardRef("CTDLblPos"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showLegendKey",
                    "type": ForwardRef("ShowLegendKey"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showVal",
                    "type": ForwardRef("ShowVal"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showCatName",
                    "type": ForwardRef("ShowCatName"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showSerName",
                    "type": ForwardRef("ShowSerName"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showPercent",
                    "type": ForwardRef("ShowPercent"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showBubbleSize",
                    "type": ForwardRef("ShowBubbleSize"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "separator",
                    "type": str,
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "showLeaderLines",
                    "type": ForwardRef("ShowLeaderLines"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "leaderLines",
                    "type": ForwardRef("CTChartLines"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
            "max_occurs": 13,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDateAx(Child):
    class Meta:
        name = "CT_DateAx"

    ax_id: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    scaling: None | CTScaling = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    delete: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_pos: None | CTAxPos = field(
        default=None,
        metadata={
            "name": "axPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "majorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "minorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    title: None | CTTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    num_fmt: None | CTNumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "majorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "minorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_lbl_pos: None | CTTickLblPos = field(
        default=None,
        metadata={
            "name": "tickLblPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cross_ax: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "crossAx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    crosses_or_crosses_at: None | CTCrosses | CTDouble = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "crosses",
                    "type": ForwardRef("CTCrosses"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "crossesAt",
                    "type": ForwardRef("CTDouble"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    auto: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    lbl_offset: None | CTLblOffset = field(
        default=None,
        metadata={
            "name": "lblOffset",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    base_time_unit: None | CTTimeUnit = field(
        default=None,
        metadata={
            "name": "baseTimeUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_unit: None | CTAxisUnit = field(
        default=None,
        metadata={
            "name": "majorUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_time_unit: None | CTTimeUnit = field(
        default=None,
        metadata={
            "name": "majorTimeUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_unit: None | CTAxisUnit = field(
        default=None,
        metadata={
            "name": "minorUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_time_unit: None | CTTimeUnit = field(
        default=None,
        metadata={
            "name": "minorTimeUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDispUnits(Child):
    class Meta:
        name = "CT_DispUnits"

    cust_unit_or_built_in_unit: None | CTDouble | CTBuiltInUnit = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "custUnit",
                    "type": ForwardRef("CTDouble"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "builtInUnit",
                    "type": ForwardRef("CTBuiltInUnit"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    disp_units_lbl: None | CTDispUnitsLbl = field(
        default=None,
        metadata={
            "name": "dispUnitsLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPivotFmt(Child):
    class Meta:
        name = "CT_PivotFmt"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbl: None | CTDLbl = field(
        default=None,
        metadata={
            "name": "dLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSerAx(Child):
    class Meta:
        name = "CT_SerAx"

    ax_id: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    scaling: None | CTScaling = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    delete: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_pos: None | CTAxPos = field(
        default=None,
        metadata={
            "name": "axPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "majorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "minorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    title: None | CTTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    num_fmt: None | CTNumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "majorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "minorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_lbl_pos: None | CTTickLblPos = field(
        default=None,
        metadata={
            "name": "tickLblPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cross_ax: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "crossAx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    crosses_or_crosses_at: None | CTCrosses | CTDouble = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "crosses",
                    "type": ForwardRef("CTCrosses"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "crossesAt",
                    "type": ForwardRef("CTDouble"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    tick_lbl_skip: None | CTSkip = field(
        default=None,
        metadata={
            "name": "tickLblSkip",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_mark_skip: None | CTSkip = field(
        default=None,
        metadata={
            "name": "tickMarkSkip",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSurface3DChart(Child):
    class Meta:
        name = "CT_Surface3DChart"

    wireframe: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTSurfaceSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    band_fmts: None | CTBandFmts = field(
        default=None,
        metadata={
            "name": "bandFmts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 3,
            "max_occurs": 3,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSurfaceChart(Child):
    class Meta:
        name = "CT_SurfaceChart"

    wireframe: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTSurfaceSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    band_fmts: None | CTBandFmts = field(
        default=None,
        metadata={
            "name": "bandFmts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 3,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrendline(Child):
    class Meta:
        name = "CT_Trendline"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline_type: None | CTTrendlineType = field(
        default=None,
        metadata={
            "name": "trendlineType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTOrder = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    period: None | CTPeriod = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    forward: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    backward: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    intercept: None | CTDouble = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    disp_rsqr: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "dispRSqr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    disp_eq: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "dispEq",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline_lbl: None | CTTrendlineLbl = field(
        default=None,
        metadata={
            "name": "trendlineLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAreaSer(Child):
    class Meta:
        name = "CT_AreaSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    picture_options: None | CTPictureOptions = field(
        default=None,
        metadata={
            "name": "pictureOptions",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline: list[CTTrendline] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_bars: list[CTErrBars] = field(
        default_factory=ChildList,
        metadata={
            "name": "errBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "max_occurs": 2,
        },
    )
    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBarSer(Child):
    class Meta:
        name = "CT_BarSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    invert_if_negative: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "invertIfNegative",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    picture_options: None | CTPictureOptions = field(
        default=None,
        metadata={
            "name": "pictureOptions",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline: list[CTTrendline] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_bars: None | CTErrBars = field(
        default=None,
        metadata={
            "name": "errBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    shape: None | CTShape = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBubbleSer(Child):
    class Meta:
        name = "CT_BubbleSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    invert_if_negative: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "invertIfNegative",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline: list[CTTrendline] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_bars: list[CTErrBars] = field(
        default_factory=ChildList,
        metadata={
            "name": "errBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "max_occurs": 2,
        },
    )
    x_val: None | CTAxDataSource = field(
        default=None,
        metadata={
            "name": "xVal",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    y_val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "name": "yVal",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    bubble_size: None | CTNumDataSource = field(
        default=None,
        metadata={
            "name": "bubbleSize",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    bubble3_d: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "bubble3D",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineSer(Child):
    class Meta:
        name = "CT_LineSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline: list[CTTrendline] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_bars: None | CTErrBars = field(
        default=None,
        metadata={
            "name": "errBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    smooth: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPieSer(Child):
    class Meta:
        name = "CT_PieSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    explosion: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPivotFmts(Child):
    class Meta:
        name = "CT_PivotFmts"

    pivot_fmt: list[CTPivotFmt] = field(
        default_factory=ChildList,
        metadata={
            "name": "pivotFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRadarSer(Child):
    class Meta:
        name = "CT_RadarSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cat: None | CTAxDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScatterSer(Child):
    class Meta:
        name = "CT_ScatterSer"

    idx: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    order: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx: None | CTSerTx = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    marker: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_pt: list[CTDPt] = field(
        default_factory=ChildList,
        metadata={
            "name": "dPt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    trendline: list[CTTrendline] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    err_bars: list[CTErrBars] = field(
        default_factory=ChildList,
        metadata={
            "name": "errBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "max_occurs": 2,
        },
    )
    x_val: None | CTAxDataSource = field(
        default=None,
        metadata={
            "name": "xVal",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    y_val: None | CTNumDataSource = field(
        default=None,
        metadata={
            "name": "yVal",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    smooth: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTValAx(Child):
    class Meta:
        name = "CT_ValAx"

    ax_id: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    scaling: None | CTScaling = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    delete: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_pos: None | CTAxPos = field(
        default=None,
        metadata={
            "name": "axPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "majorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_gridlines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "minorGridlines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    title: None | CTTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    num_fmt: None | CTNumFmt = field(
        default=None,
        metadata={
            "name": "numFmt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "majorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_tick_mark: None | CTTickMark = field(
        default=None,
        metadata={
            "name": "minorTickMark",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tick_lbl_pos: None | CTTickLblPos = field(
        default=None,
        metadata={
            "name": "tickLblPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cross_ax: None | CTUnsignedInt = field(
        default=None,
        metadata={
            "name": "crossAx",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    crosses_or_crosses_at: None | CTCrosses | CTDouble = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "crosses",
                    "type": ForwardRef("CTCrosses"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "crossesAt",
                    "type": ForwardRef("CTDouble"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    cross_between: None | CTCrossBetween = field(
        default=None,
        metadata={
            "name": "crossBetween",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    major_unit: None | CTAxisUnit = field(
        default=None,
        metadata={
            "name": "majorUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    minor_unit: None | CTAxisUnit = field(
        default=None,
        metadata={
            "name": "minorUnit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    disp_units: None | CTDispUnits = field(
        default=None,
        metadata={
            "name": "dispUnits",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTArea3DChart(Child):
    class Meta:
        name = "CT_Area3DChart"

    grouping: None | CTGrouping = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTAreaSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    drop_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "dropLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    gap_depth: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapDepth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 3,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAreaChart(Child):
    class Meta:
        name = "CT_AreaChart"

    grouping: None | CTGrouping = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTAreaSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    drop_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "dropLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBar3DChart(Child):
    class Meta:
        name = "CT_Bar3DChart"

    bar_dir: None | CTBarDir = field(
        default=None,
        metadata={
            "name": "barDir",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    grouping: None | CTBarGrouping = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTBarSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    gap_width: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    gap_depth: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapDepth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    shape: None | CTShape = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 3,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBarChart(Child):
    class Meta:
        name = "CT_BarChart"

    bar_dir: None | CTBarDir = field(
        default=None,
        metadata={
            "name": "barDir",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    grouping: None | CTBarGrouping = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTBarSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    gap_width: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    overlap: None | CTOverlap = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser_lines: list[CTChartLines] = field(
        default_factory=ChildList,
        metadata={
            "name": "serLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBubbleChart(Child):
    class Meta:
        name = "CT_BubbleChart"

    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTBubbleSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    bubble3_d: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "bubble3D",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    bubble_scale: None | CTBubbleScale = field(
        default=None,
        metadata={
            "name": "bubbleScale",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    show_neg_bubbles: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "showNegBubbles",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    size_represents: None | CTSizeRepresents = field(
        default=None,
        metadata={
            "name": "sizeRepresents",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDoughnutChart(Child):
    class Meta:
        name = "CT_DoughnutChart"

    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTPieSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    first_slice_ang: None | CTFirstSliceAng = field(
        default=None,
        metadata={
            "name": "firstSliceAng",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    hole_size: None | CTHoleSize = field(
        default=None,
        metadata={
            "name": "holeSize",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLine3DChart(Child):
    class Meta:
        name = "CT_Line3DChart"

    grouping: None | CTGrouping = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTLineSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    drop_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "dropLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    gap_depth: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapDepth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 3,
            "max_occurs": 3,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineChart(Child):
    class Meta:
        name = "CT_LineChart"

    grouping: None | CTGrouping = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTLineSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    drop_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "dropLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    hi_low_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "hiLowLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    up_down_bars: None | CTUpDownBars = field(
        default=None,
        metadata={
            "name": "upDownBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    marker: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    smooth: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOfPieChart(Child):
    class Meta:
        name = "CT_OfPieChart"

    of_pie_type: None | CTOfPieType = field(
        default=None,
        metadata={
            "name": "ofPieType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTPieSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    gap_width: None | CTGapAmount = field(
        default=None,
        metadata={
            "name": "gapWidth",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    split_type: None | CTSplitType = field(
        default=None,
        metadata={
            "name": "splitType",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    split_pos: None | CTDouble = field(
        default=None,
        metadata={
            "name": "splitPos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    cust_split: None | CTCustSplit = field(
        default=None,
        metadata={
            "name": "custSplit",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    second_pie_size: None | CTSecondPieSize = field(
        default=None,
        metadata={
            "name": "secondPieSize",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser_lines: list[CTChartLines] = field(
        default_factory=ChildList,
        metadata={
            "name": "serLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPie3DChart(Child):
    class Meta:
        name = "CT_Pie3DChart"

    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTPieSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPieChart(Child):
    class Meta:
        name = "CT_PieChart"

    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTPieSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    first_slice_ang: None | CTFirstSliceAng = field(
        default=None,
        metadata={
            "name": "firstSliceAng",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRadarChart(Child):
    class Meta:
        name = "CT_RadarChart"

    radar_style: None | CTRadarStyle = field(
        default=None,
        metadata={
            "name": "radarStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTRadarSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScatterChart(Child):
    class Meta:
        name = "CT_ScatterChart"

    scatter_style: None | CTScatterStyle = field(
        default=None,
        metadata={
            "name": "scatterStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    vary_colors: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "varyColors",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ser: list[CTScatterSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStockChart(Child):
    class Meta:
        name = "CT_StockChart"

    ser: list[CTLineSer] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 3,
            "max_occurs": 4,
        },
    )
    d_lbls: None | CTDLbls = field(
        default=None,
        metadata={
            "name": "dLbls",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    drop_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "dropLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    hi_low_lines: None | CTChartLines = field(
        default=None,
        metadata={
            "name": "hiLowLines",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    up_down_bars: None | CTUpDownBars = field(
        default=None,
        metadata={
            "name": "upDownBars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ax_id: list[CTUnsignedInt] = field(
        default_factory=ChildList,
        metadata={
            "name": "axId",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "min_occurs": 2,
            "max_occurs": 2,
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPlotArea(Child):
    class Meta:
        name = "CT_PlotArea"

    layout: None | CTLayout = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    content: list[
        CTAreaChart
        | CTArea3DChart
        | CTLineChart
        | CTLine3DChart
        | CTStockChart
        | CTRadarChart
        | CTScatterChart
        | CTPieChart
        | CTPie3DChart
        | CTDoughnutChart
        | CTBarChart
        | CTBar3DChart
        | CTOfPieChart
        | CTSurfaceChart
        | CTSurface3DChart
        | CTBubbleChart
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "areaChart",
                    "type": ForwardRef("CTAreaChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "area3DChart",
                    "type": ForwardRef("CTArea3DChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "lineChart",
                    "type": ForwardRef("CTLineChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "line3DChart",
                    "type": ForwardRef("CTLine3DChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "stockChart",
                    "type": ForwardRef("CTStockChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "radarChart",
                    "type": ForwardRef("CTRadarChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "scatterChart",
                    "type": ForwardRef("CTScatterChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "pieChart",
                    "type": ForwardRef("CTPieChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "pie3DChart",
                    "type": ForwardRef("CTPie3DChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "doughnutChart",
                    "type": ForwardRef("CTDoughnutChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "barChart",
                    "type": ForwardRef("CTBarChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "bar3DChart",
                    "type": ForwardRef("CTBar3DChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "ofPieChart",
                    "type": ForwardRef("CTOfPieChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "surfaceChart",
                    "type": ForwardRef("CTSurfaceChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "surface3DChart",
                    "type": ForwardRef("CTSurface3DChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "bubbleChart",
                    "type": ForwardRef("CTBubbleChart"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    content_1: list[CTValAx | CTCatAx | CTDateAx | CTSerAx] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "valAx",
                    "type": ForwardRef("CTValAx"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "catAx",
                    "type": ForwardRef("CTCatAx"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "dateAx",
                    "type": ForwardRef("CTDateAx"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
                {
                    "name": "serAx",
                    "type": ForwardRef("CTSerAx"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                },
            ),
        },
    )
    d_table: None | CTDTable = field(
        default=None,
        metadata={
            "name": "dTable",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChart(Child):
    class Meta:
        name = "CT_Chart"

    title: None | CTTitle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    auto_title_deleted: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "autoTitleDeleted",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    pivot_fmts: None | CTPivotFmts = field(
        default=None,
        metadata={
            "name": "pivotFmts",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    view3_d: None | CTView3D = field(
        default=None,
        metadata={
            "name": "view3D",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    floor: None | CTSurface = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    side_wall: None | CTSurface = field(
        default=None,
        metadata={
            "name": "sideWall",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    back_wall: None | CTSurface = field(
        default=None,
        metadata={
            "name": "backWall",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    plot_area: None | CTPlotArea = field(
        default=None,
        metadata={
            "name": "plotArea",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    legend: None | CTLegend = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    plot_vis_only: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "plotVisOnly",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    disp_blanks_as: None | CTDispBlanksAs = field(
        default=None,
        metadata={
            "name": "dispBlanksAs",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    show_dlbls_over_max: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "showDLblsOverMax",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChartSpace(Child):
    class Meta:
        name = "CT_ChartSpace"

    date1904: None | CTBoolean = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    lang: None | CTTextLanguageID = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    rounded_corners: None | CTBoolean = field(
        default=None,
        metadata={
            "name": "roundedCorners",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    style: None | CTStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    clr_map_ovr: None | CTColorMapping = field(
        default=None,
        metadata={
            "name": "clrMapOvr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    pivot_source: None | CTPivotSource = field(
        default=None,
        metadata={
            "name": "pivotSource",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    protection: None | CTProtection = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    chart: None | CTChart = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    tx_pr: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    external_data: None | CTExternalData = field(
        default=None,
        metadata={
            "name": "externalData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    print_settings: None | CTPrintSettings = field(
        default=None,
        metadata={
            "name": "printSettings",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    user_shapes: None | CTRelId = field(
        default=None,
        metadata={
            "name": "userShapes",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )
    ext_lst: None | CTExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class ChartSpace(CTChartSpace):
    class Meta:
        name = "chartSpace"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTColorMapping,
    CTShapeProperties,
    CTTextBody,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.chart import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.dml.chart.el",
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
