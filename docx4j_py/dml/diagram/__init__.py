from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList
from docx4j_py.dml.main import CTTextBody as MainCttextBody

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class CTAdj(Child):
    class Meta:
        name = "CT_Adj"

    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 1,
        },
    )
    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBulletEnabled(Child):
    class Meta:
        name = "CT_BulletEnabled"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCTCategory(Child):
    class Meta:
        name = "CT_CTCategory"

    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    pri: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCTDescription(Child):
    class Meta:
        name = "CT_CTDescription"

    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCTName(Child):
    class Meta:
        name = "CT_CTName"

    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCategory(Child):
    class Meta:
        name = "CT_Category"

    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    pri: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChildMax(Child):
    class Meta:
        name = "CT_ChildMax"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "-1",
            "min_inclusive": -1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChildPref(Child):
    class Meta:
        name = "CT_ChildPref"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "-1",
            "min_inclusive": -1,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDescription(Child):
    class Meta:
        name = "CT_Description"

    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTName(Child):
    class Meta:
        name = "CT_Name"

    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOrgChart(Child):
    class Meta:
        name = "CT_OrgChart"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelIds(Child):
    class Meta:
        name = "CT_RelIds"

    dm: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "schema_default": "",
        },
    )
    lo: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "schema_default": "",
        },
    )
    qs: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "schema_default": "",
        },
    )
    cs: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSDCategory(Child):
    class Meta:
        name = "CT_SDCategory"

    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    pri: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSDDescription(Child):
    class Meta:
        name = "CT_SDDescription"

    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSDName(Child):
    class Meta:
        name = "CT_SDName"

    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class STAlgorithmType(Enum):
    COMPOSITE = "composite"
    CONN = "conn"
    CYCLE = "cycle"
    HIER_CHILD = "hierChild"
    HIER_ROOT = "hierRoot"
    PYRA = "pyra"
    LIN = "lin"
    SP = "sp"
    TX = "tx"
    SNAKE = "snake"


class STAnimLvlStr(Enum):
    NONE = "none"
    LVL = "lvl"
    CTR = "ctr"


class STAnimOneStr(Enum):
    NONE = "none"
    ONE = "one"
    BRANCH = "branch"


class STArrowheadStyle(Enum):
    AUTO = "auto"
    ARR = "arr"
    NO_ARR = "noArr"


class STAutoTextRotation(Enum):
    NONE = "none"
    UPR = "upr"
    GRAV = "grav"


class STAxisType(Enum):
    SELF = "self"
    CH = "ch"
    DES = "des"
    DES_OR_SELF = "desOrSelf"
    PAR = "par"
    ANCST = "ancst"
    ANCST_OR_SELF = "ancstOrSelf"
    FOLLOW_SIB = "followSib"
    PRECED_SIB = "precedSib"
    FOLLOW = "follow"
    PRECED = "preced"
    ROOT = "root"
    NONE = "none"


class STBendPoint(Enum):
    BEG = "beg"
    DEF = "def"
    END = "end"


class STBoolOperator(Enum):
    NONE = "none"
    EQU = "equ"
    GTE = "gte"
    LTE = "lte"


class STBreakpoint(Enum):
    END_CNV = "endCnv"
    BAL = "bal"
    FIXED = "fixed"


class STCenterShapeMapping(Enum):
    NONE = "none"
    F_NODE = "fNode"


class STChildAlignment(Enum):
    T = "t"
    B = "b"
    L = "l"
    R = "r"


class STChildDirection(Enum):
    HORZ = "horz"
    VERT = "vert"


class STChildOrderType(Enum):
    B = "b"
    T = "t"


class STClrAppMethod(Enum):
    SPAN = "span"
    CYCLE = "cycle"
    REPEAT = "repeat"


class STConnectorPoint(Enum):
    AUTO = "auto"
    B_CTR = "bCtr"
    CTR = "ctr"
    MID_L = "midL"
    MID_R = "midR"
    T_CTR = "tCtr"
    B_L = "bL"
    B_R = "bR"
    T_L = "tL"
    T_R = "tR"
    RADIAL = "radial"


class STConnectorRouting(Enum):
    STRA = "stra"
    BEND = "bend"
    CURVE = "curve"
    LONG_CURVE = "longCurve"


class STConstraintRelationship(Enum):
    SELF = "self"
    CH = "ch"
    DES = "des"


class STConstraintType(Enum):
    NONE = "none"
    ALIGN_OFF = "alignOff"
    BEG_MARG = "begMarg"
    BEND_DIST = "bendDist"
    BEG_PAD = "begPad"
    B = "b"
    B_MARG = "bMarg"
    B_OFF = "bOff"
    CTR_X = "ctrX"
    CTR_XOFF = "ctrXOff"
    CTR_Y = "ctrY"
    CTR_YOFF = "ctrYOff"
    CONN_DIST = "connDist"
    DIAM = "diam"
    END_MARG = "endMarg"
    END_PAD = "endPad"
    H = "h"
    H_AR_H = "hArH"
    H_OFF = "hOff"
    L = "l"
    L_MARG = "lMarg"
    L_OFF = "lOff"
    R = "r"
    R_MARG = "rMarg"
    R_OFF = "rOff"
    PRIM_FONT_SZ = "primFontSz"
    PYRA_ACCT_RATIO = "pyraAcctRatio"
    SEC_FONT_SZ = "secFontSz"
    SIB_SP = "sibSp"
    SEC_SIB_SP = "secSibSp"
    SP = "sp"
    STEM_THICK = "stemThick"
    T = "t"
    T_MARG = "tMarg"
    T_OFF = "tOff"
    USER_A = "userA"
    USER_B = "userB"
    USER_C = "userC"
    USER_D = "userD"
    USER_E = "userE"
    USER_F = "userF"
    USER_G = "userG"
    USER_H = "userH"
    USER_I = "userI"
    USER_J = "userJ"
    USER_K = "userK"
    USER_L = "userL"
    USER_M = "userM"
    USER_N = "userN"
    USER_O = "userO"
    USER_P = "userP"
    USER_Q = "userQ"
    USER_R = "userR"
    USER_S = "userS"
    USER_T = "userT"
    USER_U = "userU"
    USER_V = "userV"
    USER_W = "userW"
    USER_X = "userX"
    USER_Y = "userY"
    USER_Z = "userZ"
    W = "w"
    W_AR_H = "wArH"
    W_OFF = "wOff"


class STContinueDirection(Enum):
    REV_DIR = "revDir"
    SAME_DIR = "sameDir"


class STCxnType(Enum):
    PAR_OF = "parOf"
    PRES_OF = "presOf"
    PRES_PAR_OF = "presParOf"
    UNKNOWN_RELATIONSHIP = "unknownRelationship"


class STDirection(Enum):
    NORM = "norm"
    REV = "rev"


class STElementType(Enum):
    ALL = "all"
    DOC = "doc"
    NODE = "node"
    NORM = "norm"
    NON_NORM = "nonNorm"
    ASST = "asst"
    NON_ASST = "nonAsst"
    PAR_TRANS = "parTrans"
    PRES = "pres"
    SIB_TRANS = "sibTrans"


class STFlowDirection(Enum):
    ROW = "row"
    COL = "col"


class STFunctionOperator(Enum):
    EQU = "equ"
    NEQ = "neq"
    GT = "gt"
    LT = "lt"
    GTE = "gte"
    LTE = "lte"


class STFunctionType(Enum):
    CNT = "cnt"
    POS = "pos"
    REV_POS = "revPos"
    POS_EVEN = "posEven"
    POS_ODD = "posOdd"
    VAR = "var"
    DEPTH = "depth"
    MAX_DEPTH = "maxDepth"


class STGrowDirection(Enum):
    T_L = "tL"
    T_R = "tR"
    B_L = "bL"
    B_R = "bR"


class STHierBranchStyle(Enum):
    L = "l"
    R = "r"
    HANG = "hang"
    STD = "std"
    INIT = "init"


class STHierarchyAlignment(Enum):
    T_L = "tL"
    T_R = "tR"
    T_CTR_CH = "tCtrCh"
    T_CTR_DES = "tCtrDes"
    B_L = "bL"
    B_R = "bR"
    B_CTR_CH = "bCtrCh"
    B_CTR_DES = "bCtrDes"
    L_T = "lT"
    L_B = "lB"
    L_CTR_CH = "lCtrCh"
    L_CTR_DES = "lCtrDes"
    R_T = "rT"
    R_B = "rB"
    R_CTR_CH = "rCtrCh"
    R_CTR_DES = "rCtrDes"


class STHorizontalAlignment(Enum):
    L = "l"
    CTR = "ctr"
    R = "r"
    NONE = "none"


class STHueDir(Enum):
    CW = "cw"
    CCW = "ccw"


class STLinearDirection(Enum):
    FROM_L = "fromL"
    FROM_R = "fromR"
    FROM_T = "fromT"
    FROM_B = "fromB"


class STNodeHorizontalAlignment(Enum):
    L = "l"
    CTR = "ctr"
    R = "r"


class STNodeVerticalAlignment(Enum):
    T = "t"
    MID = "mid"
    B = "b"


class STOffset(Enum):
    CTR = "ctr"
    OFF = "off"


class STParameterId(Enum):
    HORZ_ALIGN = "horzAlign"
    VERT_ALIGN = "vertAlign"
    CH_DIR = "chDir"
    CH_ALIGN = "chAlign"
    SEC_CH_ALIGN = "secChAlign"
    LIN_DIR = "linDir"
    SEC_LIN_DIR = "secLinDir"
    ST_ELEM = "stElem"
    BEND_PT = "bendPt"
    CONN_ROUT = "connRout"
    BEG_STY = "begSty"
    END_STY = "endSty"
    DIM = "dim"
    ROT_PATH = "rotPath"
    CTR_SHP_MAP = "ctrShpMap"
    NODE_HORZ_ALIGN = "nodeHorzAlign"
    NODE_VERT_ALIGN = "nodeVertAlign"
    FALLBACK = "fallback"
    TX_DIR = "txDir"
    PYRA_ACCT_POS = "pyraAcctPos"
    PYRA_ACCT_TX_MAR = "pyraAcctTxMar"
    TX_BL_DIR = "txBlDir"
    TX_ANCHOR_HORZ = "txAnchorHorz"
    TX_ANCHOR_VERT = "txAnchorVert"
    TX_ANCHOR_HORZ_CH = "txAnchorHorzCh"
    TX_ANCHOR_VERT_CH = "txAnchorVertCh"
    PAR_TX_LTRALIGN = "parTxLTRAlign"
    PAR_TX_RTLALIGN = "parTxRTLAlign"
    SHP_TX_LTRALIGN_CH = "shpTxLTRAlignCh"
    SHP_TX_RTLALIGN_CH = "shpTxRTLAlignCh"
    AUTO_TX_ROT = "autoTxRot"
    GR_DIR = "grDir"
    FLOW_DIR = "flowDir"
    CONT_DIR = "contDir"
    BKPT = "bkpt"
    OFF = "off"
    HIER_ALIGN = "hierAlign"
    BK_PT_FIXED_VAL = "bkPtFixedVal"
    ST_BULLET_LVL = "stBulletLvl"
    ST_ANG = "stAng"
    SPAN_ANG = "spanAng"
    AR = "ar"
    LN_SP_PAR = "lnSpPar"
    LN_SP_AF_PAR_P = "lnSpAfParP"
    LN_SP_CH = "lnSpCh"
    LN_SP_AF_CH_P = "lnSpAfChP"
    RT_SHORT_DIST = "rtShortDist"
    ALIGN_TX = "alignTx"
    PYRA_LVL_NODE = "pyraLvlNode"
    PYRA_ACCT_BKGD_NODE = "pyraAcctBkgdNode"
    PYRA_ACCT_TX_NODE = "pyraAcctTxNode"
    SRC_NODE = "srcNode"
    DST_NODE = "dstNode"
    BEG_PTS = "begPts"
    END_PTS = "endPts"


class STPtType(Enum):
    NODE = "node"
    ASST = "asst"
    DOC = "doc"
    PRES = "pres"
    PAR_TRANS = "parTrans"
    SIB_TRANS = "sibTrans"


class STPyramidAccentPosition(Enum):
    BEF = "bef"
    AFT = "aft"


class STPyramidAccentTextMargin(Enum):
    STEP = "step"
    STACK = "stack"


class STResizeHandlesStr(Enum):
    EXACT = "exact"
    REL = "rel"


class STRotationPath(Enum):
    NONE = "none"
    ALONG_PATH = "alongPath"


class STSecondaryChildAlignment(Enum):
    NONE = "none"
    T = "t"
    B = "b"
    L = "l"
    R = "r"


class STSecondaryLinearDirection(Enum):
    NONE = "none"
    FROM_L = "fromL"
    FROM_R = "fromR"
    FROM_T = "fromT"
    FROM_B = "fromB"


class STStartingElement(Enum):
    NODE = "node"
    TRANS = "trans"


class STTextAlignment(Enum):
    L = "l"
    CTR = "ctr"
    R = "r"


class STTextAnchorHorizontal(Enum):
    NONE = "none"
    CTR = "ctr"


class STTextAnchorVertical(Enum):
    T = "t"
    MID = "mid"
    B = "b"


class STTextBlockDirection(Enum):
    HORZ = "horz"
    VERT = "vert"


class STTextDirection(Enum):
    FROM_T = "fromT"
    FROM_B = "fromB"


class STVerticalAlignment(Enum):
    T = "t"
    MID = "mid"
    B = "b"
    NONE = "none"


class StConnectorDimension(Enum):
    VALUE_1_D = "1D"
    VALUE_2_D = "2D"
    CUST = "cust"


class StFallbackDimension(Enum):
    VALUE_1_D = "1D"
    VALUE_2_D = "2D"


class StFunctionArgument(Enum):
    NONE = "none"
    ORG_CHART = "orgChart"
    CH_MAX = "chMax"
    CH_PREF = "chPref"
    BUL_ENABLED = "bulEnabled"
    DIR = "dir"
    HIER_BRANCH = "hierBranch"
    ANIM_ONE = "animOne"
    ANIM_LVL = "animLvl"
    RESIZE_HANDLES = "resizeHandles"


class StLayoutShapeType(Enum):
    LINE = "line"
    LINE_INV = "lineInv"
    TRIANGLE = "triangle"
    RT_TRIANGLE = "rtTriangle"
    RECT = "rect"
    DIAMOND = "diamond"
    PARALLELOGRAM = "parallelogram"
    TRAPEZOID = "trapezoid"
    NON_ISOSCELES_TRAPEZOID = "nonIsoscelesTrapezoid"
    PENTAGON = "pentagon"
    HEXAGON = "hexagon"
    HEPTAGON = "heptagon"
    OCTAGON = "octagon"
    DECAGON = "decagon"
    DODECAGON = "dodecagon"
    STAR4 = "star4"
    STAR5 = "star5"
    STAR6 = "star6"
    STAR7 = "star7"
    STAR8 = "star8"
    STAR10 = "star10"
    STAR12 = "star12"
    STAR16 = "star16"
    STAR24 = "star24"
    STAR32 = "star32"
    ROUND_RECT = "roundRect"
    ROUND1_RECT = "round1Rect"
    ROUND2_SAME_RECT = "round2SameRect"
    ROUND2_DIAG_RECT = "round2DiagRect"
    SNIP_ROUND_RECT = "snipRoundRect"
    SNIP1_RECT = "snip1Rect"
    SNIP2_SAME_RECT = "snip2SameRect"
    SNIP2_DIAG_RECT = "snip2DiagRect"
    PLAQUE = "plaque"
    ELLIPSE = "ellipse"
    TEARDROP = "teardrop"
    HOME_PLATE = "homePlate"
    CHEVRON = "chevron"
    PIE_WEDGE = "pieWedge"
    PIE = "pie"
    BLOCK_ARC = "blockArc"
    DONUT = "donut"
    NO_SMOKING = "noSmoking"
    RIGHT_ARROW = "rightArrow"
    LEFT_ARROW = "leftArrow"
    UP_ARROW = "upArrow"
    DOWN_ARROW = "downArrow"
    STRIPED_RIGHT_ARROW = "stripedRightArrow"
    NOTCHED_RIGHT_ARROW = "notchedRightArrow"
    BENT_UP_ARROW = "bentUpArrow"
    LEFT_RIGHT_ARROW = "leftRightArrow"
    UP_DOWN_ARROW = "upDownArrow"
    LEFT_UP_ARROW = "leftUpArrow"
    LEFT_RIGHT_UP_ARROW = "leftRightUpArrow"
    QUAD_ARROW = "quadArrow"
    LEFT_ARROW_CALLOUT = "leftArrowCallout"
    RIGHT_ARROW_CALLOUT = "rightArrowCallout"
    UP_ARROW_CALLOUT = "upArrowCallout"
    DOWN_ARROW_CALLOUT = "downArrowCallout"
    LEFT_RIGHT_ARROW_CALLOUT = "leftRightArrowCallout"
    UP_DOWN_ARROW_CALLOUT = "upDownArrowCallout"
    QUAD_ARROW_CALLOUT = "quadArrowCallout"
    BENT_ARROW = "bentArrow"
    UTURN_ARROW = "uturnArrow"
    CIRCULAR_ARROW = "circularArrow"
    LEFT_CIRCULAR_ARROW = "leftCircularArrow"
    LEFT_RIGHT_CIRCULAR_ARROW = "leftRightCircularArrow"
    CURVED_RIGHT_ARROW = "curvedRightArrow"
    CURVED_LEFT_ARROW = "curvedLeftArrow"
    CURVED_UP_ARROW = "curvedUpArrow"
    CURVED_DOWN_ARROW = "curvedDownArrow"
    SWOOSH_ARROW = "swooshArrow"
    CUBE = "cube"
    CAN = "can"
    LIGHTNING_BOLT = "lightningBolt"
    HEART = "heart"
    SUN = "sun"
    MOON = "moon"
    SMILEY_FACE = "smileyFace"
    IRREGULAR_SEAL1 = "irregularSeal1"
    IRREGULAR_SEAL2 = "irregularSeal2"
    FOLDED_CORNER = "foldedCorner"
    BEVEL = "bevel"
    FRAME = "frame"
    HALF_FRAME = "halfFrame"
    CORNER = "corner"
    DIAG_STRIPE = "diagStripe"
    CHORD = "chord"
    ARC = "arc"
    LEFT_BRACKET = "leftBracket"
    RIGHT_BRACKET = "rightBracket"
    LEFT_BRACE = "leftBrace"
    RIGHT_BRACE = "rightBrace"
    BRACKET_PAIR = "bracketPair"
    BRACE_PAIR = "bracePair"
    STRAIGHT_CONNECTOR1 = "straightConnector1"
    BENT_CONNECTOR2 = "bentConnector2"
    BENT_CONNECTOR3 = "bentConnector3"
    BENT_CONNECTOR4 = "bentConnector4"
    BENT_CONNECTOR5 = "bentConnector5"
    CURVED_CONNECTOR2 = "curvedConnector2"
    CURVED_CONNECTOR3 = "curvedConnector3"
    CURVED_CONNECTOR4 = "curvedConnector4"
    CURVED_CONNECTOR5 = "curvedConnector5"
    CALLOUT1 = "callout1"
    CALLOUT2 = "callout2"
    CALLOUT3 = "callout3"
    ACCENT_CALLOUT1 = "accentCallout1"
    ACCENT_CALLOUT2 = "accentCallout2"
    ACCENT_CALLOUT3 = "accentCallout3"
    BORDER_CALLOUT1 = "borderCallout1"
    BORDER_CALLOUT2 = "borderCallout2"
    BORDER_CALLOUT3 = "borderCallout3"
    ACCENT_BORDER_CALLOUT1 = "accentBorderCallout1"
    ACCENT_BORDER_CALLOUT2 = "accentBorderCallout2"
    ACCENT_BORDER_CALLOUT3 = "accentBorderCallout3"
    WEDGE_RECT_CALLOUT = "wedgeRectCallout"
    WEDGE_ROUND_RECT_CALLOUT = "wedgeRoundRectCallout"
    WEDGE_ELLIPSE_CALLOUT = "wedgeEllipseCallout"
    CLOUD_CALLOUT = "cloudCallout"
    CLOUD = "cloud"
    RIBBON = "ribbon"
    RIBBON2 = "ribbon2"
    ELLIPSE_RIBBON = "ellipseRibbon"
    ELLIPSE_RIBBON2 = "ellipseRibbon2"
    LEFT_RIGHT_RIBBON = "leftRightRibbon"
    VERTICAL_SCROLL = "verticalScroll"
    HORIZONTAL_SCROLL = "horizontalScroll"
    WAVE = "wave"
    DOUBLE_WAVE = "doubleWave"
    PLUS = "plus"
    FLOW_CHART_PROCESS = "flowChartProcess"
    FLOW_CHART_DECISION = "flowChartDecision"
    FLOW_CHART_INPUT_OUTPUT = "flowChartInputOutput"
    FLOW_CHART_PREDEFINED_PROCESS = "flowChartPredefinedProcess"
    FLOW_CHART_INTERNAL_STORAGE = "flowChartInternalStorage"
    FLOW_CHART_DOCUMENT = "flowChartDocument"
    FLOW_CHART_MULTIDOCUMENT = "flowChartMultidocument"
    FLOW_CHART_TERMINATOR = "flowChartTerminator"
    FLOW_CHART_PREPARATION = "flowChartPreparation"
    FLOW_CHART_MANUAL_INPUT = "flowChartManualInput"
    FLOW_CHART_MANUAL_OPERATION = "flowChartManualOperation"
    FLOW_CHART_CONNECTOR = "flowChartConnector"
    FLOW_CHART_PUNCHED_CARD = "flowChartPunchedCard"
    FLOW_CHART_PUNCHED_TAPE = "flowChartPunchedTape"
    FLOW_CHART_SUMMING_JUNCTION = "flowChartSummingJunction"
    FLOW_CHART_OR = "flowChartOr"
    FLOW_CHART_COLLATE = "flowChartCollate"
    FLOW_CHART_SORT = "flowChartSort"
    FLOW_CHART_EXTRACT = "flowChartExtract"
    FLOW_CHART_MERGE = "flowChartMerge"
    FLOW_CHART_OFFLINE_STORAGE = "flowChartOfflineStorage"
    FLOW_CHART_ONLINE_STORAGE = "flowChartOnlineStorage"
    FLOW_CHART_MAGNETIC_TAPE = "flowChartMagneticTape"
    FLOW_CHART_MAGNETIC_DISK = "flowChartMagneticDisk"
    FLOW_CHART_MAGNETIC_DRUM = "flowChartMagneticDrum"
    FLOW_CHART_DISPLAY = "flowChartDisplay"
    FLOW_CHART_DELAY = "flowChartDelay"
    FLOW_CHART_ALTERNATE_PROCESS = "flowChartAlternateProcess"
    FLOW_CHART_OFFPAGE_CONNECTOR = "flowChartOffpageConnector"
    ACTION_BUTTON_BLANK = "actionButtonBlank"
    ACTION_BUTTON_HOME = "actionButtonHome"
    ACTION_BUTTON_HELP = "actionButtonHelp"
    ACTION_BUTTON_INFORMATION = "actionButtonInformation"
    ACTION_BUTTON_FORWARD_NEXT = "actionButtonForwardNext"
    ACTION_BUTTON_BACK_PREVIOUS = "actionButtonBackPrevious"
    ACTION_BUTTON_END = "actionButtonEnd"
    ACTION_BUTTON_BEGINNING = "actionButtonBeginning"
    ACTION_BUTTON_RETURN = "actionButtonReturn"
    ACTION_BUTTON_DOCUMENT = "actionButtonDocument"
    ACTION_BUTTON_SOUND = "actionButtonSound"
    ACTION_BUTTON_MOVIE = "actionButtonMovie"
    GEAR6 = "gear6"
    GEAR9 = "gear9"
    FUNNEL = "funnel"
    MATH_PLUS = "mathPlus"
    MATH_MINUS = "mathMinus"
    MATH_MULTIPLY = "mathMultiply"
    MATH_DIVIDE = "mathDivide"
    MATH_EQUAL = "mathEqual"
    MATH_NOT_EQUAL = "mathNotEqual"
    CORNER_TABS = "cornerTabs"
    SQUARE_TABS = "squareTabs"
    PLAQUE_TABS = "plaqueTabs"
    CHART_X = "chartX"
    CHART_STAR = "chartStar"
    CHART_PLUS = "chartPlus"
    NONE = "none"
    CONN = "conn"


@dataclass(slots=True, kw_only=True)
class CTAdjLst(Child):
    class Meta:
        name = "CT_AdjLst"

    adj: list[CTAdj] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimLvl(Child):
    class Meta:
        name = "CT_AnimLvl"

    val: None | STAnimLvlStr = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STAnimLvlStr.NONE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimOne(Child):
    class Meta:
        name = "CT_AnimOne"

    val: None | STAnimOneStr = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STAnimOneStr.ONE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCTCategories(Child):
    class Meta:
        name = "CT_CTCategories"

    cat: list[CTCTCategory] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCategories(Child):
    class Meta:
        name = "CT_Categories"

    cat: list[CTCategory] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColors(Child):
    class Meta:
        name = "CT_Colors"

    content: list[
        CTScRgbColor | CTSRgbColor | CTHslColor | CTSystemColor | CTSchemeColor | CTPresetColor
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "scrgbClr",
                    "type": ForwardRef("CTScRgbColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hslClr",
                    "type": ForwardRef("CTHslColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sysClr",
                    "type": ForwardRef("CTSystemColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "prstClr",
                    "type": ForwardRef("CTPresetColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    meth: None | STClrAppMethod = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STClrAppMethod.SPAN,
        },
    )
    hue_dir: None | STHueDir = field(
        default=None,
        metadata={
            "name": "hueDir",
            "type": "Attribute",
            "schema_default": STHueDir.CW,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConstraint(Child):
    class Meta:
        name = "CT_Constraint"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    type_value: None | STConstraintType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    for_value: None | STConstraintRelationship = field(
        default=None,
        metadata={
            "name": "for",
            "type": "Attribute",
            "schema_default": STConstraintRelationship.SELF,
        },
    )
    for_name: None | str = field(
        default=None,
        metadata={
            "name": "forName",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    pt_type: None | STElementType = field(
        default=None,
        metadata={
            "name": "ptType",
            "type": "Attribute",
            "schema_default": STElementType.ALL,
        },
    )
    ref_type: None | STConstraintType = field(
        default=None,
        metadata={
            "name": "refType",
            "type": "Attribute",
            "schema_default": STConstraintType.NONE,
        },
    )
    ref_for: None | STConstraintRelationship = field(
        default=None,
        metadata={
            "name": "refFor",
            "type": "Attribute",
            "schema_default": STConstraintRelationship.SELF,
        },
    )
    ref_for_name: None | str = field(
        default=None,
        metadata={
            "name": "refForName",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    ref_pt_type: None | STElementType = field(
        default=None,
        metadata={
            "name": "refPtType",
            "type": "Attribute",
            "schema_default": STElementType.ALL,
        },
    )
    op: None | STBoolOperator = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STBoolOperator.NONE,
        },
    )
    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    fact: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "1",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCxn(Child):
    class Meta:
        name = "CT_Cxn"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    model_id: None | int | str = field(
        default=None,
        metadata={
            "name": "modelId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    type_value: None | STCxnType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": STCxnType.PAR_OF,
        },
    )
    src_id: None | int | str = field(
        default=None,
        metadata={
            "name": "srcId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    dest_id: None | int | str = field(
        default=None,
        metadata={
            "name": "destId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    src_ord: None | int = field(
        default=None,
        metadata={
            "name": "srcOrd",
            "type": "Attribute",
        },
    )
    dest_ord: None | int = field(
        default=None,
        metadata={
            "name": "destOrd",
            "type": "Attribute",
        },
    )
    par_trans_id: None | int | str = field(
        default=None,
        metadata={
            "name": "parTransId",
            "type": "Attribute",
            "schema_default": "0",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    sib_trans_id: None | int | str = field(
        default=None,
        metadata={
            "name": "sibTransId",
            "type": "Attribute",
            "schema_default": "0",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    pres_id: None | str = field(
        default=None,
        metadata={
            "name": "presId",
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDirection(Child):
    class Meta:
        name = "CT_Direction"

    val: None | STDirection = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STDirection.NORM,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHierBranchStyle(Child):
    class Meta:
        name = "CT_HierBranchStyle"

    val: None | STHierBranchStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STHierBranchStyle.STD,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumericRule(Child):
    class Meta:
        name = "CT_NumericRule"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    type_value: None | STConstraintType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    for_value: None | STConstraintRelationship = field(
        default=None,
        metadata={
            "name": "for",
            "type": "Attribute",
            "schema_default": STConstraintRelationship.SELF,
        },
    )
    for_name: None | str = field(
        default=None,
        metadata={
            "name": "forName",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    pt_type: None | STElementType = field(
        default=None,
        metadata={
            "name": "ptType",
            "type": "Attribute",
            "schema_default": STElementType.ALL,
        },
    )
    val: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "NaN",
        },
    )
    fact: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "NaN",
        },
    )
    max: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "NaN",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTParameter(Child):
    class Meta:
        name = "CT_Parameter"

    type_value: None | STParameterId = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    val: (
        None
        | STHorizontalAlignment
        | STVerticalAlignment
        | STChildDirection
        | STChildAlignment
        | STSecondaryChildAlignment
        | STLinearDirection
        | STSecondaryLinearDirection
        | STStartingElement
        | STBendPoint
        | STConnectorRouting
        | STArrowheadStyle
        | StConnectorDimension
        | STRotationPath
        | STCenterShapeMapping
        | STNodeHorizontalAlignment
        | STNodeVerticalAlignment
        | StFallbackDimension
        | STTextDirection
        | STPyramidAccentPosition
        | STPyramidAccentTextMargin
        | STTextBlockDirection
        | STTextAnchorHorizontal
        | STTextAnchorVertical
        | STTextAlignment
        | STAutoTextRotation
        | STGrowDirection
        | STFlowDirection
        | STContinueDirection
        | STBreakpoint
        | STOffset
        | STHierarchyAlignment
        | int
        | float
        | bool
        | str
        | STConnectorPoint
    ) = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresentationOf(Child):
    class Meta:
        name = "CT_PresentationOf"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    axis: list[STAxisType] = field(
        default_factory=lambda: [
            STAxisType.NONE,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    pt_type: list[STElementType] = field(
        default_factory=lambda: [
            STElementType.ALL,
        ],
        metadata={
            "name": "ptType",
            "type": "Attribute",
            "tokens": True,
        },
    )
    hide_last_trans: list[bool] = field(
        default_factory=lambda: [
            True,
        ],
        metadata={
            "name": "hideLastTrans",
            "type": "Attribute",
            "tokens": True,
        },
    )
    st: list[int] = field(
        default_factory=lambda: [
            1,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    cnt: list[int] = field(
        default_factory=lambda: [
            0,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    step: list[int] = field(
        default_factory=lambda: [
            1,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTResizeHandles(Child):
    class Meta:
        name = "CT_ResizeHandles"

    val: None | STResizeHandlesStr = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STResizeHandlesStr.REL,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSDCategories(Child):
    class Meta:
        name = "CT_SDCategories"

    cat: list[CTSDCategory] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBody(MainCttextBody):
    class Meta:
        name = "t"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class CTTextProps(Child):
    class Meta:
        name = "CT_TextProps"

    sp3d_or_flat_tx: None | CTShape3D | CTFlatText = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "sp3d",
                    "type": ForwardRef("CTShape3D"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "flatTx",
                    "type": ForwardRef("CTFlatText"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class RelIds(CTRelIds):
    class Meta:
        name = "relIds"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class CTAlgorithm(Child):
    class Meta:
        name = "CT_Algorithm"

    param: list[CTParameter] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    type_value: None | STAlgorithmType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    rev: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCTStyleLabel(Child):
    class Meta:
        name = "CT_CTStyleLabel"

    fill_clr_lst: None | CTColors = field(
        default=None,
        metadata={
            "name": "fillClrLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    lin_clr_lst: None | CTColors = field(
        default=None,
        metadata={
            "name": "linClrLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    effect_clr_lst: None | CTColors = field(
        default=None,
        metadata={
            "name": "effectClrLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    tx_lin_clr_lst: None | CTColors = field(
        default=None,
        metadata={
            "name": "txLinClrLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    tx_fill_clr_lst: None | CTColors = field(
        default=None,
        metadata={
            "name": "txFillClrLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    tx_effect_clr_lst: None | CTColors = field(
        default=None,
        metadata={
            "name": "txEffectClrLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorTransformHeader(Child):
    class Meta:
        name = "CT_ColorTransformHeader"

    title: list[CTCTName] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    desc: list[CTCTDescription] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    cat_lst: None | CTCTCategories = field(
        default=None,
        metadata={
            "name": "catLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
            "schema_default": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    res_id: None | int = field(
        default=None,
        metadata={
            "name": "resId",
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConstraints(Child):
    class Meta:
        name = "CT_Constraints"

    constr: list[CTConstraint] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCxnList(Child):
    class Meta:
        name = "CT_CxnList"

    cxn: list[CTCxn] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDiagramDefinitionHeader(Child):
    class Meta:
        name = "CT_DiagramDefinitionHeader"

    title: list[CTName] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    desc: list[CTDescription] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    cat_lst: None | CTCategories = field(
        default=None,
        metadata={
            "name": "catLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
            "schema_default": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    def_style: None | str = field(
        default=None,
        metadata={
            "name": "defStyle",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    res_id: None | int = field(
        default=None,
        metadata={
            "name": "resId",
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLayoutVariablePropertySet(Child):
    class Meta:
        name = "CT_LayoutVariablePropertySet"

    org_chart: None | CTOrgChart = field(
        default=None,
        metadata={
            "name": "orgChart",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ch_max: None | CTChildMax = field(
        default=None,
        metadata={
            "name": "chMax",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ch_pref: None | CTChildPref = field(
        default=None,
        metadata={
            "name": "chPref",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    bullet_enabled: None | CTBulletEnabled = field(
        default=None,
        metadata={
            "name": "bulletEnabled",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    dir: None | CTDirection = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    hier_branch: None | CTHierBranchStyle = field(
        default=None,
        metadata={
            "name": "hierBranch",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    anim_one: None | CTAnimOne = field(
        default=None,
        metadata={
            "name": "animOne",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    anim_lvl: None | CTAnimLvl = field(
        default=None,
        metadata={
            "name": "animLvl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    resize_handles: None | CTResizeHandles = field(
        default=None,
        metadata={
            "name": "resizeHandles",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRules(Child):
    class Meta:
        name = "CT_Rules"

    rule: list[CTNumericRule] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShape(Child):
    class Meta:
        name = "CT_Shape"

    adj_lst: None | CTAdjLst = field(
        default=None,
        metadata={
            "name": "adjLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    rot: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    type_value: None | StLayoutShapeType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": StLayoutShapeType.NONE,
        },
    )
    blip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "schema_default": "",
        },
    )
    z_order_off: None | int = field(
        default=None,
        metadata={
            "name": "zOrderOff",
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    hide_geom: None | bool = field(
        default=None,
        metadata={
            "name": "hideGeom",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    lk_tx_entry: None | bool = field(
        default=None,
        metadata={
            "name": "lkTxEntry",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    blip_phldr: None | bool = field(
        default=None,
        metadata={
            "name": "blipPhldr",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleDefinitionHeader(Child):
    class Meta:
        name = "CT_StyleDefinitionHeader"

    title: list[CTSDName] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    desc: list[CTSDDescription] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    cat_lst: None | CTSDCategories = field(
        default=None,
        metadata={
            "name": "catLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
            "schema_default": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    res_id: None | int = field(
        default=None,
        metadata={
            "name": "resId",
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleLabel(Child):
    class Meta:
        name = "CT_StyleLabel"

    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    sp3d: None | CTShape3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    tx_pr: None | CTTextProps = field(
        default=None,
        metadata={
            "name": "txPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorTransform(Child):
    class Meta:
        name = "CT_ColorTransform"

    title: list[CTCTName] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    desc: list[CTCTDescription] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    cat_lst: None | CTCTCategories = field(
        default=None,
        metadata={
            "name": "catLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    style_lbl: list[CTCTStyleLabel] = field(
        default_factory=ChildList,
        metadata={
            "name": "styleLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
            "schema_default": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorTransformHeaderLst(Child):
    class Meta:
        name = "CT_ColorTransformHeaderLst"

    colors_def_hdr: list[CTColorTransformHeader] = field(
        default_factory=ChildList,
        metadata={
            "name": "colorsDefHdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDiagramDefinitionHeaderLst(Child):
    class Meta:
        name = "CT_DiagramDefinitionHeaderLst"

    layout_def_hdr: list[CTDiagramDefinitionHeader] = field(
        default_factory=ChildList,
        metadata={
            "name": "layoutDefHdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTElemPropSet(Child):
    class Meta:
        name = "CT_ElemPropSet"

    pres_layout_vars: None | CTLayoutVariablePropertySet = field(
        default=None,
        metadata={
            "name": "presLayoutVars",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    pres_assoc_id: None | int | str = field(
        default=None,
        metadata={
            "name": "presAssocID",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    pres_name: None | str = field(
        default=None,
        metadata={
            "name": "presName",
            "type": "Attribute",
        },
    )
    pres_style_lbl: None | str = field(
        default=None,
        metadata={
            "name": "presStyleLbl",
            "type": "Attribute",
        },
    )
    pres_style_idx: None | int = field(
        default=None,
        metadata={
            "name": "presStyleIdx",
            "type": "Attribute",
        },
    )
    pres_style_cnt: None | int = field(
        default=None,
        metadata={
            "name": "presStyleCnt",
            "type": "Attribute",
        },
    )
    lo_type_id: None | str = field(
        default=None,
        metadata={
            "name": "loTypeId",
            "type": "Attribute",
        },
    )
    lo_cat_id: None | str = field(
        default=None,
        metadata={
            "name": "loCatId",
            "type": "Attribute",
        },
    )
    qs_type_id: None | str = field(
        default=None,
        metadata={
            "name": "qsTypeId",
            "type": "Attribute",
        },
    )
    qs_cat_id: None | str = field(
        default=None,
        metadata={
            "name": "qsCatId",
            "type": "Attribute",
        },
    )
    cs_type_id: None | str = field(
        default=None,
        metadata={
            "name": "csTypeId",
            "type": "Attribute",
        },
    )
    cs_cat_id: None | str = field(
        default=None,
        metadata={
            "name": "csCatId",
            "type": "Attribute",
        },
    )
    coherent3_doff: None | bool = field(
        default=None,
        metadata={
            "name": "coherent3DOff",
            "type": "Attribute",
        },
    )
    phldr_t: None | str = field(
        default=None,
        metadata={
            "name": "phldrT",
            "type": "Attribute",
        },
    )
    phldr: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cust_ang: None | int = field(
        default=None,
        metadata={
            "name": "custAng",
            "type": "Attribute",
        },
    )
    cust_flip_vert: None | bool = field(
        default=None,
        metadata={
            "name": "custFlipVert",
            "type": "Attribute",
        },
    )
    cust_flip_hor: None | bool = field(
        default=None,
        metadata={
            "name": "custFlipHor",
            "type": "Attribute",
        },
    )
    cust_sz_x: None | int = field(
        default=None,
        metadata={
            "name": "custSzX",
            "type": "Attribute",
        },
    )
    cust_sz_y: None | int = field(
        default=None,
        metadata={
            "name": "custSzY",
            "type": "Attribute",
        },
    )
    cust_scale_x: None | int = field(
        default=None,
        metadata={
            "name": "custScaleX",
            "type": "Attribute",
        },
    )
    cust_scale_y: None | int = field(
        default=None,
        metadata={
            "name": "custScaleY",
            "type": "Attribute",
        },
    )
    cust_t: None | bool = field(
        default=None,
        metadata={
            "name": "custT",
            "type": "Attribute",
        },
    )
    cust_lin_fact_x: None | int = field(
        default=None,
        metadata={
            "name": "custLinFactX",
            "type": "Attribute",
        },
    )
    cust_lin_fact_y: None | int = field(
        default=None,
        metadata={
            "name": "custLinFactY",
            "type": "Attribute",
        },
    )
    cust_lin_fact_neighbor_x: None | int = field(
        default=None,
        metadata={
            "name": "custLinFactNeighborX",
            "type": "Attribute",
        },
    )
    cust_lin_fact_neighbor_y: None | int = field(
        default=None,
        metadata={
            "name": "custLinFactNeighborY",
            "type": "Attribute",
        },
    )
    cust_rad_scale_rad: None | int = field(
        default=None,
        metadata={
            "name": "custRadScaleRad",
            "type": "Attribute",
        },
    )
    cust_rad_scale_inc: None | int = field(
        default=None,
        metadata={
            "name": "custRadScaleInc",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLayoutNode(Child):
    class Meta:
        name = "CT_LayoutNode"

    content: list[
        CTAlgorithm
        | CTShape
        | CTPresentationOf
        | CTConstraints
        | CTRules
        | CTLayoutVariablePropertySet
        | CTForEach
        | CTLayoutNode
        | CTChoose
        | CTOfficeArtExtensionList
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "alg",
                    "type": ForwardRef("CTAlgorithm"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "shape",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "presOf",
                    "type": ForwardRef("CTPresentationOf"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "constrLst",
                    "type": ForwardRef("CTConstraints"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "ruleLst",
                    "type": ForwardRef("CTRules"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "varLst",
                    "type": ForwardRef("CTLayoutVariablePropertySet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "forEach",
                    "type": ForwardRef("CTForEach"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "layoutNode",
                    "type": ForwardRef("CTLayoutNode"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "choose",
                    "type": ForwardRef("CTChoose"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "extLst",
                    "type": ForwardRef("CTOfficeArtExtensionList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
            ),
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    style_lbl: None | str = field(
        default=None,
        metadata={
            "name": "styleLbl",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    ch_order: None | STChildOrderType = field(
        default=None,
        metadata={
            "name": "chOrder",
            "type": "Attribute",
            "schema_default": STChildOrderType.B,
        },
    )
    move_with: None | str = field(
        default=None,
        metadata={
            "name": "moveWith",
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleDefinition(Child):
    class Meta:
        name = "CT_StyleDefinition"

    title: list[CTSDName] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    desc: list[CTSDDescription] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    cat_lst: None | CTSDCategories = field(
        default=None,
        metadata={
            "name": "catLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    style_lbl: list[CTStyleLabel] = field(
        default_factory=ChildList,
        metadata={
            "name": "styleLbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
            "schema_default": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleDefinitionHeaderLst(Child):
    class Meta:
        name = "CT_StyleDefinitionHeaderLst"

    style_def_hdr: list[CTStyleDefinitionHeader] = field(
        default_factory=ChildList,
        metadata={
            "name": "styleDefHdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class ColorsDefHdr(CTColorTransformHeader):
    class Meta:
        name = "colorsDefHdr"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class LayoutDefHdr(CTDiagramDefinitionHeader):
    class Meta:
        name = "layoutDefHdr"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class StyleDefHdr(CTStyleDefinitionHeader):
    class Meta:
        name = "styleDefHdr"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class CTForEach(Child):
    class Meta:
        name = "CT_ForEach"

    content: list[
        CTAlgorithm
        | CTShape
        | CTPresentationOf
        | CTConstraints
        | CTRules
        | CTForEach
        | CTLayoutNode
        | CTChoose
        | CTOfficeArtExtensionList
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "alg",
                    "type": ForwardRef("CTAlgorithm"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "shape",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "presOf",
                    "type": ForwardRef("CTPresentationOf"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "constrLst",
                    "type": ForwardRef("CTConstraints"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "ruleLst",
                    "type": ForwardRef("CTRules"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "forEach",
                    "type": ForwardRef("CTForEach"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "layoutNode",
                    "type": ForwardRef("CTLayoutNode"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "choose",
                    "type": ForwardRef("CTChoose"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "extLst",
                    "type": ForwardRef("CTOfficeArtExtensionList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
            ),
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    ref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    axis: list[STAxisType] = field(
        default_factory=lambda: [
            STAxisType.NONE,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    pt_type: list[STElementType] = field(
        default_factory=lambda: [
            STElementType.ALL,
        ],
        metadata={
            "name": "ptType",
            "type": "Attribute",
            "tokens": True,
        },
    )
    hide_last_trans: list[bool] = field(
        default_factory=lambda: [
            True,
        ],
        metadata={
            "name": "hideLastTrans",
            "type": "Attribute",
            "tokens": True,
        },
    )
    st: list[int] = field(
        default_factory=lambda: [
            1,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    cnt: list[int] = field(
        default_factory=lambda: [
            0,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    step: list[int] = field(
        default_factory=lambda: [
            1,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPt(Child):
    class Meta:
        name = "CT_Pt"

    pr_set: None | CTElemPropSet = field(
        default=None,
        metadata={
            "name": "prSet",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    t: None | CTTextBody = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    model_id: None | int | str = field(
        default=None,
        metadata={
            "name": "modelId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    type_value: None | STPtType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": STPtType.NODE,
        },
    )
    cxn_id: None | int | str = field(
        default=None,
        metadata={
            "name": "cxnId",
            "type": "Attribute",
            "schema_default": "0",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class ColorsDef(CTColorTransform):
    class Meta:
        name = "colorsDef"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class ColorsDefHdrLst(CTColorTransformHeaderLst):
    class Meta:
        name = "colorsDefHdrLst"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class LayoutDefHdrLst(CTDiagramDefinitionHeaderLst):
    class Meta:
        name = "layoutDefHdrLst"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class StyleDef(CTStyleDefinition):
    class Meta:
        name = "styleDef"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class StyleDefHdrLst(CTStyleDefinitionHeaderLst):
    class Meta:
        name = "styleDefHdrLst"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class CTOtherwise(Child):
    class Meta:
        name = "CT_Otherwise"

    content: list[
        CTAlgorithm
        | CTShape
        | CTPresentationOf
        | CTConstraints
        | CTRules
        | CTForEach
        | CTLayoutNode
        | CTChoose
        | CTOfficeArtExtensionList
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "alg",
                    "type": ForwardRef("CTAlgorithm"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "shape",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "presOf",
                    "type": ForwardRef("CTPresentationOf"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "constrLst",
                    "type": ForwardRef("CTConstraints"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "ruleLst",
                    "type": ForwardRef("CTRules"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "forEach",
                    "type": ForwardRef("CTForEach"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "layoutNode",
                    "type": ForwardRef("CTLayoutNode"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "choose",
                    "type": ForwardRef("CTChoose"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "extLst",
                    "type": ForwardRef("CTOfficeArtExtensionList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
            ),
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPtList(Child):
    class Meta:
        name = "CT_PtList"

    pt: list[CTPt] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWhen(Child):
    class Meta:
        name = "CT_When"

    content: list[
        CTAlgorithm
        | CTShape
        | CTPresentationOf
        | CTConstraints
        | CTRules
        | CTForEach
        | CTLayoutNode
        | CTChoose
        | CTOfficeArtExtensionList
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "alg",
                    "type": ForwardRef("CTAlgorithm"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "shape",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "presOf",
                    "type": ForwardRef("CTPresentationOf"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "constrLst",
                    "type": ForwardRef("CTConstraints"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "ruleLst",
                    "type": ForwardRef("CTRules"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "forEach",
                    "type": ForwardRef("CTForEach"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "layoutNode",
                    "type": ForwardRef("CTLayoutNode"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "choose",
                    "type": ForwardRef("CTChoose"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
                {
                    "name": "extLst",
                    "type": ForwardRef("CTOfficeArtExtensionList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                },
            ),
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    axis: list[STAxisType] = field(
        default_factory=lambda: [
            STAxisType.NONE,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    pt_type: list[STElementType] = field(
        default_factory=lambda: [
            STElementType.ALL,
        ],
        metadata={
            "name": "ptType",
            "type": "Attribute",
            "tokens": True,
        },
    )
    hide_last_trans: list[bool] = field(
        default_factory=lambda: [
            True,
        ],
        metadata={
            "name": "hideLastTrans",
            "type": "Attribute",
            "tokens": True,
        },
    )
    st: list[int] = field(
        default_factory=lambda: [
            1,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    cnt: list[int] = field(
        default_factory=lambda: [
            0,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    step: list[int] = field(
        default_factory=lambda: [
            1,
        ],
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )
    func: None | STFunctionType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    arg: None | StFunctionArgument = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": StFunctionArgument.NONE,
        },
    )
    op: None | STFunctionOperator = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    val: (
        None
        | int
        | bool
        | STDirection
        | STHierBranchStyle
        | STAnimOneStr
        | STAnimLvlStr
        | STResizeHandlesStr
    ) = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTChoose(Child):
    class Meta:
        name = "CT_Choose"

    if_value: list[CTWhen] = field(
        default_factory=ChildList,
        metadata={
            "name": "if",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            "min_occurs": 1,
        },
    )
    else_value: None | CTOtherwise = field(
        default=None,
        metadata={
            "name": "else",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataModel(Child):
    class Meta:
        name = "CT_DataModel"

    pt_lst: None | CTPtList = field(
        default=None,
        metadata={
            "name": "ptLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    cxn_lst: None | CTCxnList = field(
        default=None,
        metadata={
            "name": "cxnLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    bg: None | CTBackgroundFormatting = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    whole: None | CTWholeE2OFormatting = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSampleData(Child):
    class Meta:
        name = "CT_SampleData"

    data_model: None | CTDataModel = field(
        default=None,
        metadata={
            "name": "dataModel",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    use_def: None | bool = field(
        default=None,
        metadata={
            "name": "useDef",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class DataModel(CTDataModel):
    class Meta:
        name = "dataModel"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


@dataclass(slots=True, kw_only=True)
class CTDiagramDefinition(Child):
    class Meta:
        name = "CT_DiagramDefinition"

    title: list[CTName] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    desc: list[CTDescription] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    cat_lst: None | CTCategories = field(
        default=None,
        metadata={
            "name": "catLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    samp_data: None | CTSampleData = field(
        default=None,
        metadata={
            "name": "sampData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    style_data: None | CTSampleData = field(
        default=None,
        metadata={
            "name": "styleData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    clr_data: None | CTSampleData = field(
        default=None,
        metadata={
            "name": "clrData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    layout_node: None | CTLayoutNode = field(
        default=None,
        metadata={
            "name": "layoutNode",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    unique_id: None | str = field(
        default=None,
        metadata={
            "name": "uniqueId",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
            "schema_default": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        },
    )
    def_style: None | str = field(
        default=None,
        metadata={
            "name": "defStyle",
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class LayoutDef(CTDiagramDefinition):
    class Meta:
        name = "layoutDef"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTBackgroundFormatting,
    CTFlatText,
    CTHslColor,
    CTOfficeArtExtensionList,
    CTPresetColor,
    CTScene3D,
    CTSchemeColor,
    CTScRgbColor,
    CTShape3D,
    CTShapeProperties,
    CTShapeStyle,
    CTSRgbColor,
    CTSystemColor,
    CTWholeE2OFormatting,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.diagram import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.dml.diagram.el",
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
