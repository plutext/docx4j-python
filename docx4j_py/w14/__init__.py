from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList
from docx4j_py.wml import (
    CTEmpty as WmlCtempty,
)
from docx4j_py.wml import (
    CTMarkup as WmlCtmarkup,
)
from docx4j_py.wml import (
    CTTrackChange as WmlCttrackChange,
)

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTDefaultImageDpi(Child):
    class Meta:
        name = "CT_DefaultImageDpi"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineJoinMiterProperties(Child):
    class Meta:
        name = "CT_LineJoinMiterProperties"

    lim: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLongHexNumber(Child):
    class Meta:
        name = "CT_LongHexNumber"

    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPercentage(Child):
    class Meta:
        name = "CT_Percentage"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPositiveFixedPercentage(Child):
    class Meta:
        name = "CT_PositiveFixedPercentage"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPositivePercentage(Child):
    class Meta:
        name = "CT_PositivePercentage"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelativeRect(Child):
    class Meta:
        name = "CT_RelativeRect"

    l: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    t: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    r: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    b: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtCheckboxSymbol(Child):
    class Meta:
        name = "CT_SdtCheckboxSymbol"

    font: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSphereCoords(Child):
    class Meta:
        name = "CT_SphereCoords"

    lat: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    lon: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    rev: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )


class STBevelPresetType(Enum):
    RELAXED_INSET = "relaxedInset"
    CIRCLE = "circle"
    SLOPE = "slope"
    CROSS = "cross"
    ANGLE = "angle"
    SOFT_ROUND = "softRound"
    CONVEX = "convex"
    COOL_SLANT = "coolSlant"
    DIVOT = "divot"
    RIBLET = "riblet"
    HARD_EDGE = "hardEdge"
    ART_DECO = "artDeco"


class STCompoundLine(Enum):
    SNG = "sng"
    DBL = "dbl"
    THICK_THIN = "thickThin"
    THIN_THICK = "thinThick"
    TRI = "tri"


class STLigatures(Enum):
    NONE = "none"
    STANDARD = "standard"
    CONTEXTUAL = "contextual"
    HISTORICAL = "historical"
    DISCRETIONAL = "discretional"
    STANDARD_CONTEXTUAL = "standardContextual"
    STANDARD_HISTORICAL = "standardHistorical"
    CONTEXTUAL_HISTORICAL = "contextualHistorical"
    STANDARD_DISCRETIONAL = "standardDiscretional"
    CONTEXTUAL_DISCRETIONAL = "contextualDiscretional"
    HISTORICAL_DISCRETIONAL = "historicalDiscretional"
    STANDARD_CONTEXTUAL_HISTORICAL = "standardContextualHistorical"
    STANDARD_CONTEXTUAL_DISCRETIONAL = "standardContextualDiscretional"
    STANDARD_HISTORICAL_DISCRETIONAL = "standardHistoricalDiscretional"
    CONTEXTUAL_HISTORICAL_DISCRETIONAL = "contextualHistoricalDiscretional"
    ALL = "all"


class STLightRigDirection(Enum):
    TL = "tl"
    T = "t"
    TR = "tr"
    L = "l"
    R = "r"
    BL = "bl"
    B = "b"
    BR = "br"


class STLightRigType(Enum):
    LEGACY_FLAT1 = "legacyFlat1"
    LEGACY_FLAT2 = "legacyFlat2"
    LEGACY_FLAT3 = "legacyFlat3"
    LEGACY_FLAT4 = "legacyFlat4"
    LEGACY_NORMAL1 = "legacyNormal1"
    LEGACY_NORMAL2 = "legacyNormal2"
    LEGACY_NORMAL3 = "legacyNormal3"
    LEGACY_NORMAL4 = "legacyNormal4"
    LEGACY_HARSH1 = "legacyHarsh1"
    LEGACY_HARSH2 = "legacyHarsh2"
    LEGACY_HARSH3 = "legacyHarsh3"
    LEGACY_HARSH4 = "legacyHarsh4"
    THREE_PT = "threePt"
    BALANCED = "balanced"
    SOFT = "soft"
    HARSH = "harsh"
    FLOOD = "flood"
    CONTRASTING = "contrasting"
    MORNING = "morning"
    SUNRISE = "sunrise"
    SUNSET = "sunset"
    CHILLY = "chilly"
    FREEZING = "freezing"
    FLAT = "flat"
    TWO_PT = "twoPt"
    GLOW = "glow"
    BRIGHT_ROOM = "brightRoom"


class STLineCap(Enum):
    RND = "rnd"
    SQ = "sq"
    FLAT = "flat"


class STNumForm(Enum):
    DEFAULT = "default"
    LINING = "lining"
    OLD_STYLE = "oldStyle"


class STNumSpacing(Enum):
    DEFAULT = "default"
    PROPORTIONAL = "proportional"
    TABULAR = "tabular"


class STPathShadeType(Enum):
    SHAPE = "shape"
    CIRCLE = "circle"
    RECT = "rect"


class STPenAlignment(Enum):
    CTR = "ctr"
    IN = "in"


class STPresetCameraType(Enum):
    LEGACY_OBLIQUE_TOP_LEFT = "legacyObliqueTopLeft"
    LEGACY_OBLIQUE_TOP = "legacyObliqueTop"
    LEGACY_OBLIQUE_TOP_RIGHT = "legacyObliqueTopRight"
    LEGACY_OBLIQUE_LEFT = "legacyObliqueLeft"
    LEGACY_OBLIQUE_FRONT = "legacyObliqueFront"
    LEGACY_OBLIQUE_RIGHT = "legacyObliqueRight"
    LEGACY_OBLIQUE_BOTTOM_LEFT = "legacyObliqueBottomLeft"
    LEGACY_OBLIQUE_BOTTOM = "legacyObliqueBottom"
    LEGACY_OBLIQUE_BOTTOM_RIGHT = "legacyObliqueBottomRight"
    LEGACY_PERSPECTIVE_TOP_LEFT = "legacyPerspectiveTopLeft"
    LEGACY_PERSPECTIVE_TOP = "legacyPerspectiveTop"
    LEGACY_PERSPECTIVE_TOP_RIGHT = "legacyPerspectiveTopRight"
    LEGACY_PERSPECTIVE_LEFT = "legacyPerspectiveLeft"
    LEGACY_PERSPECTIVE_FRONT = "legacyPerspectiveFront"
    LEGACY_PERSPECTIVE_RIGHT = "legacyPerspectiveRight"
    LEGACY_PERSPECTIVE_BOTTOM_LEFT = "legacyPerspectiveBottomLeft"
    LEGACY_PERSPECTIVE_BOTTOM = "legacyPerspectiveBottom"
    LEGACY_PERSPECTIVE_BOTTOM_RIGHT = "legacyPerspectiveBottomRight"
    ORTHOGRAPHIC_FRONT = "orthographicFront"
    ISOMETRIC_TOP_UP = "isometricTopUp"
    ISOMETRIC_TOP_DOWN = "isometricTopDown"
    ISOMETRIC_BOTTOM_UP = "isometricBottomUp"
    ISOMETRIC_BOTTOM_DOWN = "isometricBottomDown"
    ISOMETRIC_LEFT_UP = "isometricLeftUp"
    ISOMETRIC_LEFT_DOWN = "isometricLeftDown"
    ISOMETRIC_RIGHT_UP = "isometricRightUp"
    ISOMETRIC_RIGHT_DOWN = "isometricRightDown"
    ISOMETRIC_OFF_AXIS1_LEFT = "isometricOffAxis1Left"
    ISOMETRIC_OFF_AXIS1_RIGHT = "isometricOffAxis1Right"
    ISOMETRIC_OFF_AXIS1_TOP = "isometricOffAxis1Top"
    ISOMETRIC_OFF_AXIS2_LEFT = "isometricOffAxis2Left"
    ISOMETRIC_OFF_AXIS2_RIGHT = "isometricOffAxis2Right"
    ISOMETRIC_OFF_AXIS2_TOP = "isometricOffAxis2Top"
    ISOMETRIC_OFF_AXIS3_LEFT = "isometricOffAxis3Left"
    ISOMETRIC_OFF_AXIS3_RIGHT = "isometricOffAxis3Right"
    ISOMETRIC_OFF_AXIS3_BOTTOM = "isometricOffAxis3Bottom"
    ISOMETRIC_OFF_AXIS4_LEFT = "isometricOffAxis4Left"
    ISOMETRIC_OFF_AXIS4_RIGHT = "isometricOffAxis4Right"
    ISOMETRIC_OFF_AXIS4_BOTTOM = "isometricOffAxis4Bottom"
    OBLIQUE_TOP_LEFT = "obliqueTopLeft"
    OBLIQUE_TOP = "obliqueTop"
    OBLIQUE_TOP_RIGHT = "obliqueTopRight"
    OBLIQUE_LEFT = "obliqueLeft"
    OBLIQUE_RIGHT = "obliqueRight"
    OBLIQUE_BOTTOM_LEFT = "obliqueBottomLeft"
    OBLIQUE_BOTTOM = "obliqueBottom"
    OBLIQUE_BOTTOM_RIGHT = "obliqueBottomRight"
    PERSPECTIVE_FRONT = "perspectiveFront"
    PERSPECTIVE_LEFT = "perspectiveLeft"
    PERSPECTIVE_RIGHT = "perspectiveRight"
    PERSPECTIVE_ABOVE = "perspectiveAbove"
    PERSPECTIVE_BELOW = "perspectiveBelow"
    PERSPECTIVE_ABOVE_LEFT_FACING = "perspectiveAboveLeftFacing"
    PERSPECTIVE_ABOVE_RIGHT_FACING = "perspectiveAboveRightFacing"
    PERSPECTIVE_CONTRASTING_LEFT_FACING = "perspectiveContrastingLeftFacing"
    PERSPECTIVE_CONTRASTING_RIGHT_FACING = "perspectiveContrastingRightFacing"
    PERSPECTIVE_HEROIC_LEFT_FACING = "perspectiveHeroicLeftFacing"
    PERSPECTIVE_HEROIC_RIGHT_FACING = "perspectiveHeroicRightFacing"
    PERSPECTIVE_HEROIC_EXTREME_LEFT_FACING = "perspectiveHeroicExtremeLeftFacing"
    PERSPECTIVE_HEROIC_EXTREME_RIGHT_FACING = "perspectiveHeroicExtremeRightFacing"
    PERSPECTIVE_RELAXED = "perspectiveRelaxed"
    PERSPECTIVE_RELAXED_MODERATELY = "perspectiveRelaxedModerately"


class STPresetLineDashVal(Enum):
    SOLID = "solid"
    DOT = "dot"
    SYS_DOT = "sysDot"
    DASH = "dash"
    SYS_DASH = "sysDash"
    LG_DASH = "lgDash"
    DASH_DOT = "dashDot"
    SYS_DASH_DOT = "sysDashDot"
    LG_DASH_DOT = "lgDashDot"
    LG_DASH_DOT_DOT = "lgDashDotDot"
    SYS_DASH_DOT_DOT = "sysDashDotDot"


class STPresetMaterialType(Enum):
    LEGACY_MATTE = "legacyMatte"
    LEGACY_PLASTIC = "legacyPlastic"
    LEGACY_METAL = "legacyMetal"
    LEGACY_WIREFRAME = "legacyWireframe"
    MATTE = "matte"
    PLASTIC = "plastic"
    METAL = "metal"
    WARM_MATTE = "warmMatte"
    TRANSLUCENT_POWDER = "translucentPowder"
    POWDER = "powder"
    DK_EDGE = "dkEdge"
    SOFT_EDGE = "softEdge"
    CLEAR = "clear"
    FLAT = "flat"
    SOFTMETAL = "softmetal"
    NONE = "none"


class STRectAlignment(Enum):
    NONE = "none"
    TL = "tl"
    T = "t"
    TR = "tr"
    L = "l"
    CTR = "ctr"
    R = "r"
    BL = "bl"
    B = "b"
    BR = "br"


class STSchemeColorVal(Enum):
    BG1 = "bg1"
    TX1 = "tx1"
    BG2 = "bg2"
    TX2 = "tx2"
    ACCENT1 = "accent1"
    ACCENT2 = "accent2"
    ACCENT3 = "accent3"
    ACCENT4 = "accent4"
    ACCENT5 = "accent5"
    ACCENT6 = "accent6"
    HLINK = "hlink"
    FOL_HLINK = "folHlink"
    DK1 = "dk1"
    LT1 = "lt1"
    DK2 = "dk2"
    LT2 = "lt2"
    PH_CLR = "phClr"


class StOnOff(Enum):
    TRUE = "true"
    FALSE = "false"
    VALUE_0 = "0"
    VALUE_1 = "1"


@dataclass(slots=True, kw_only=True)
class CTBevel(Child):
    class Meta:
        name = "CT_Bevel"

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    prst: None | STBevelPresetType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCamera(Child):
    class Meta:
        name = "CT_Camera"

    prst: None | STPresetCameraType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEmpty(WmlCtempty):
    class Meta:
        name = "entityPicker"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTLigatures(Child):
    class Meta:
        name = "CT_Ligatures"

    val: None | STLigatures = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLightRig(Child):
    class Meta:
        name = "CT_LightRig"

    rot: None | CTSphereCoords = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    rig: None | STLightRigType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    dir: None | STLightRigDirection = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLinearShadeProperties(Child):
    class Meta:
        name = "CT_LinearShadeProperties"

    ang: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    scaled: None | StOnOff = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarkup(WmlCtmarkup):
    class Meta:
        name = "customXmlConflictInsRangeEnd"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTNumForm(Child):
    class Meta:
        name = "CT_NumForm"

    val: None | STNumForm = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumSpacing(Child):
    class Meta:
        name = "CT_NumSpacing"

    val: None | STNumSpacing = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOnOff(Child):
    class Meta:
        name = "CT_OnOff"

    val: None | StOnOff = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPathShadeProperties(Child):
    class Meta:
        name = "CT_PathShadeProperties"

    fill_to_rect: None | CTRelativeRect = field(
        default=None,
        metadata={
            "name": "fillToRect",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    path: None | STPathShadeType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresetLineDashProperties(Child):
    class Meta:
        name = "CT_PresetLineDashProperties"

    val: None | STPresetLineDashVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTReflection(Child):
    class Meta:
        name = "CT_Reflection"

    blur_rad: None | int = field(
        default=None,
        metadata={
            "name": "blurRad",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    st_a: None | int = field(
        default=None,
        metadata={
            "name": "stA",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    st_pos: None | int = field(
        default=None,
        metadata={
            "name": "stPos",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    end_a: None | int = field(
        default=None,
        metadata={
            "name": "endA",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    end_pos: None | int = field(
        default=None,
        metadata={
            "name": "endPos",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    dist: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dir: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    fade_dir: None | int = field(
        default=None,
        metadata={
            "name": "fadeDir",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    sx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    sy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    kx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    ky: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    algn: None | STRectAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleSet(Child):
    class Meta:
        name = "CT_StyleSet"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    val: None | StOnOff = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTrackChange(WmlCttrackChange):
    class Meta:
        name = "customXmlConflictInsRangeStart"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTWordContentPartNonVisual(Child):
    class Meta:
        name = "CT_WordContentPartNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    c_nv_content_part_pr: None | CTNonVisualInkContentPartProperties = field(
        default=None,
        metadata={
            "name": "cNvContentPartPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Alpha(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Bevel(WmlCtempty):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CustomXmlConflictDelRangeEnd(WmlCtmarkup):
    class Meta:
        name = "customXmlConflictDelRangeEnd"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CustomXmlConflictDelRangeStart(WmlCttrackChange):
    class Meta:
        name = "customXmlConflictDelRangeStart"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class DefaultImageDpi(CTDefaultImageDpi):
    class Meta:
        name = "defaultImageDpi"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class DocId(CTLongHexNumber):
    class Meta:
        name = "docId"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class Lum(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class LumMod(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class LumOff(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Round(WmlCtempty):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Sat(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class SatMod(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class SatOff(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Shade(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Tint(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTSRgbColor(Child):
    class Meta:
        name = "CT_SRgbColor"

    content: list[
        Tint | Shade | Alpha | CTPositivePercentage | Sat | SatOff | SatMod | Lum | LumOff | LumMod
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("CTPositivePercentage"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    val: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "length": 3,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScene3D(Child):
    class Meta:
        name = "CT_Scene3D"

    camera: None | CTCamera = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    light_rig: None | CTLightRig = field(
        default=None,
        metadata={
            "name": "lightRig",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSchemeColor(Child):
    class Meta:
        name = "CT_SchemeColor"

    content: list[
        Tint | Shade | Alpha | CTPositivePercentage | Sat | SatOff | SatMod | Lum | LumOff | LumMod
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("CTPositivePercentage"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    val: None | STSchemeColorVal = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSdtCheckbox(Child):
    class Meta:
        name = "CT_SdtCheckbox"

    checked: None | CTOnOff = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    checked_state: None | CTSdtCheckboxSymbol = field(
        default=None,
        metadata={
            "name": "checkedState",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    unchecked_state: None | CTSdtCheckboxSymbol = field(
        default=None,
        metadata={
            "name": "uncheckedState",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStylisticSets(Child):
    class Meta:
        name = "CT_StylisticSets"

    style_set: list[CTStyleSet] = field(
        default_factory=ChildList,
        metadata={
            "name": "styleSet",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWordContentPart(Child):
    class Meta:
        name = "CT_WordContentPart"

    nv_content_part_pr: None | CTWordContentPartNonVisual = field(
        default=None,
        metadata={
            "name": "nvContentPartPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    bw_mode: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwMode",
            "type": "Attribute",
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
class ConflictMode(CTOnOff):
    class Meta:
        name = "conflictMode"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class DiscardImageEditingData(CTOnOff):
    class Meta:
        name = "discardImageEditingData"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTColor(Child):
    class Meta:
        name = "CT_Color"

    srgb_clr_or_scheme_clr: None | CTSRgbColor | CTSchemeColor = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGlow(Child):
    class Meta:
        name = "CT_Glow"

    srgb_clr_or_scheme_clr: None | CTSRgbColor | CTSchemeColor = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGradientStop(Child):
    class Meta:
        name = "CT_GradientStop"

    srgb_clr_or_scheme_clr: None | CTSRgbColor | CTSchemeColor = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    pos: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShadow(Child):
    class Meta:
        name = "CT_Shadow"

    srgb_clr_or_scheme_clr: None | CTSRgbColor | CTSchemeColor = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    blur_rad: None | int = field(
        default=None,
        metadata={
            "name": "blurRad",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dist: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dir: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    sx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    sy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    kx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    ky: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    algn: None | STRectAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSolidColorFillProperties(Child):
    class Meta:
        name = "CT_SolidColorFillProperties"

    srgb_clr_or_scheme_clr: None | CTSRgbColor | CTSchemeColor = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class Checkbox(CTSdtCheckbox):
    class Meta:
        name = "checkbox"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class ContentPart(CTWordContentPart):
    class Meta:
        name = "contentPart"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTGradientStopList(Child):
    class Meta:
        name = "CT_GradientStopList"

    gs: list[CTGradientStop] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_occurs": 2,
            "max_occurs": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTProps3D(Child):
    class Meta:
        name = "CT_Props3D"

    bevel_t: None | CTBevel = field(
        default=None,
        metadata={
            "name": "bevelT",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    bevel_b: None | CTBevel = field(
        default=None,
        metadata={
            "name": "bevelB",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    extrusion_clr: None | CTColor = field(
        default=None,
        metadata={
            "name": "extrusionClr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    contour_clr: None | CTColor = field(
        default=None,
        metadata={
            "name": "contourClr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    extrusion_h: None | int = field(
        default=None,
        metadata={
            "name": "extrusionH",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    contour_w: None | int = field(
        default=None,
        metadata={
            "name": "contourW",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    prst_material: None | STPresetMaterialType = field(
        default=None,
        metadata={
            "name": "prstMaterial",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Shadow(CTShadow):
    class Meta:
        name = "shadow"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordml"


@dataclass(slots=True, kw_only=True)
class CTGradientFillProperties(Child):
    class Meta:
        name = "CT_GradientFillProperties"

    gs_lst: None | CTGradientStopList = field(
        default=None,
        metadata={
            "name": "gsLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    lin_or_path: None | CTLinearShadeProperties | CTPathShadeProperties = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "lin",
                    "type": ForwardRef("CTLinearShadeProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "path",
                    "type": ForwardRef("CTPathShadeProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFillTextEffect(Child):
    class Meta:
        name = "CT_FillTextEffect"

    no_fill_or_solid_fill_or_grad_fill: (
        None | WmlCtempty | CTSolidColorFillProperties | CTGradientFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("WmlCtempty"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextOutlineEffect(Child):
    class Meta:
        name = "CT_TextOutlineEffect"

    no_fill_or_solid_fill_or_grad_fill: (
        None | WmlCtempty | CTSolidColorFillProperties | CTGradientFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("WmlCtempty"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    prst_dash: None | CTPresetLineDashProperties = field(
        default=None,
        metadata={
            "name": "prstDash",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    round_or_bevel_or_miter: None | Round | Bevel | CTLineJoinMiterProperties = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "round",
                    "type": ForwardRef("Round"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "bevel",
                    "type": ForwardRef("Bevel"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "miter",
                    "type": ForwardRef("CTLineJoinMiterProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
            ),
        },
    )
    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
            "min_inclusive": 0,
            "max_inclusive": 20116800,
        },
    )
    cap: None | STLineCap = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    cmpd: None | STCompoundLine = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    algn: None | STPenAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTNonVisualDrawingProps,
    CTOfficeArtExtensionList,
    CTTransform2D,
    STBlackWhiteMode,
)
from docx4j_py.oart.main_2010 import CTNonVisualInkContentPartProperties


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.w14 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.w14.el",
    "element_name": "docx4j_py.traversal",
    "find": "docx4j_py.traversal",
    "iter_nodes": "docx4j_py.traversal",
    "link_parents": "docx4j_py.child",
    "run_items_of": "docx4j_py.traversal",
    "text_of": "docx4j_py.traversal",
    "to_xml": "docx4j_py.fragments",
    "walk": "docx4j_py.traversal",
    "walk_all": "docx4j_py.traversal",
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
