from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTAdjPoint2D(Child):
    class Meta:
        name = "CT_AdjPoint2D"

    x: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    y: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAlphaBiLevelEffect(Child):
    class Meta:
        name = "CT_AlphaBiLevelEffect"

    thresh: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAlphaCeilingEffect(Child):
    class Meta:
        name = "CT_AlphaCeilingEffect"


@dataclass(slots=True, kw_only=True)
class CTAlphaFloorEffect(Child):
    class Meta:
        name = "CT_AlphaFloorEffect"


@dataclass(slots=True, kw_only=True)
class CTAlphaModulateFixedEffect(Child):
    class Meta:
        name = "CT_AlphaModulateFixedEffect"

    amt: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAlphaOutsetEffect(Child):
    class Meta:
        name = "CT_AlphaOutsetEffect"

    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAlphaReplaceEffect(Child):
    class Meta:
        name = "CT_AlphaReplaceEffect"

    a: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAngle(Child):
    class Meta:
        name = "CT_Angle"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAudioCDTime(Child):
    class Meta:
        name = "CT_AudioCDTime"

    track: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    time: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBiLevelEffect(Child):
    class Meta:
        name = "CT_BiLevelEffect"

    thresh: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBlurEffect(Child):
    class Meta:
        name = "CT_BlurEffect"

    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    grow: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTComplementTransform(Child):
    class Meta:
        name = "CT_ComplementTransform"


@dataclass(slots=True, kw_only=True)
class CTConnection(Child):
    class Meta:
        name = "CT_Connection"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDashStop(Child):
    class Meta:
        name = "CT_DashStop"

    d: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
        },
    )
    sp: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEffectReference(Child):
    class Meta:
        name = "CT_EffectReference"

    ref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEmbeddedWAVAudioFile(Child):
    class Meta:
        name = "CT_EmbeddedWAVAudioFile"

    embed: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    built_in: None | bool = field(
        default=None,
        metadata={
            "name": "builtIn",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEmptyElement(Child):
    class Meta:
        name = "CT_EmptyElement"

    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFixedPercentage(Child):
    class Meta:
        name = "CT_FixedPercentage"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFlatText(Child):
    class Meta:
        name = "CT_FlatText"

    z: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGammaTransform(Child):
    class Meta:
        name = "CT_GammaTransform"


@dataclass(slots=True, kw_only=True)
class CTGeomGuide(Child):
    class Meta:
        name = "CT_GeomGuide"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fmla: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeomRect(Child):
    class Meta:
        name = "CT_GeomRect"

    l: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    t: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    r: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    b: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGrayscaleEffect(Child):
    class Meta:
        name = "CT_GrayscaleEffect"


@dataclass(slots=True, kw_only=True)
class CTGrayscaleTransform(Child):
    class Meta:
        name = "CT_GrayscaleTransform"


@dataclass(slots=True, kw_only=True)
class CTGroupFillProperties(Child):
    class Meta:
        name = "CT_GroupFillProperties"


@dataclass(slots=True, kw_only=True)
class CTGvmlUseShapeRectangle(Child):
    class Meta:
        name = "CT_GvmlUseShapeRectangle"


@dataclass(slots=True, kw_only=True)
class CTHSLEffect(Child):
    class Meta:
        name = "CT_HSLEffect"

    hue: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    sat: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )
    lum: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInverseGammaTransform(Child):
    class Meta:
        name = "CT_InverseGammaTransform"


@dataclass(slots=True, kw_only=True)
class CTInverseTransform(Child):
    class Meta:
        name = "CT_InverseTransform"


@dataclass(slots=True, kw_only=True)
class CTLineJoinBevel(Child):
    class Meta:
        name = "CT_LineJoinBevel"


@dataclass(slots=True, kw_only=True)
class CTLineJoinMiterProperties(Child):
    class Meta:
        name = "CT_LineJoinMiterProperties"

    lim: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineJoinRound(Child):
    class Meta:
        name = "CT_LineJoinRound"


@dataclass(slots=True, kw_only=True)
class CTLinearShadeProperties(Child):
    class Meta:
        name = "CT_LinearShadeProperties"

    ang: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    scaled: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLuminanceEffect(Child):
    class Meta:
        name = "CT_LuminanceEffect"

    bright: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )
    contrast: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNoFillProperties(Child):
    class Meta:
        name = "CT_NoFillProperties"


@dataclass(slots=True, kw_only=True)
class CTOfficeArtExtension(Child):
    class Meta:
        name = "CT_OfficeArtExtension"

    any_element: None | object = field(
        default=None,
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
class CTPath2DArcTo(Child):
    class Meta:
        name = "CT_Path2DArcTo"

    w_r: None | int | str = field(
        default=None,
        metadata={
            "name": "wR",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    h_r: None | int | str = field(
        default=None,
        metadata={
            "name": "hR",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    st_ang: None | int | str = field(
        default=None,
        metadata={
            "name": "stAng",
            "type": "Attribute",
        },
    )
    sw_ang: None | int | str = field(
        default=None,
        metadata={
            "name": "swAng",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2DClose(Child):
    class Meta:
        name = "CT_Path2DClose"


@dataclass(slots=True, kw_only=True)
class CTPercentage(Child):
    class Meta:
        name = "CT_Percentage"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPoint2D(Child):
    class Meta:
        name = "CT_Point2D"

    x: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    y: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPoint3D(Child):
    class Meta:
        name = "CT_Point3D"

    x: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    y: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    z: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPositiveFixedAngle(Child):
    class Meta:
        name = "CT_PositiveFixedAngle"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
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
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPositiveSize2D(Child):
    class Meta:
        name = "CT_PositiveSize2D"

    cx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    cy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRatio(Child):
    class Meta:
        name = "CT_Ratio"

    n: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    d: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelativeOffsetEffect(Child):
    class Meta:
        name = "CT_RelativeOffsetEffect"

    tx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    ty: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
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
            "schema_default": "0",
        },
    )
    t: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    r: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    b: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSoftEdgesEffect(Child):
    class Meta:
        name = "CT_SoftEdgesEffect"

    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
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
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    lon: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    rev: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBulletColorFollowText(Child):
    class Meta:
        name = "CT_TextBulletColorFollowText"


@dataclass(slots=True, kw_only=True)
class CTTextBulletSizeFollowText(Child):
    class Meta:
        name = "CT_TextBulletSizeFollowText"


@dataclass(slots=True, kw_only=True)
class CTTextBulletSizePercent(Child):
    class Meta:
        name = "CT_TextBulletSizePercent"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 25000,
            "max_inclusive": 400000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBulletSizePoint(Child):
    class Meta:
        name = "CT_TextBulletSizePoint"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 100,
            "max_inclusive": 400000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBulletTypefaceFollowText(Child):
    class Meta:
        name = "CT_TextBulletTypefaceFollowText"


@dataclass(slots=True, kw_only=True)
class CTTextCharBullet(Child):
    class Meta:
        name = "CT_TextCharBullet"

    char: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextNoAutofit(Child):
    class Meta:
        name = "CT_TextNoAutofit"


@dataclass(slots=True, kw_only=True)
class CTTextNoBullet(Child):
    class Meta:
        name = "CT_TextNoBullet"


@dataclass(slots=True, kw_only=True)
class CTTextNormalAutofit(Child):
    class Meta:
        name = "CT_TextNormalAutofit"

    font_scale: None | int = field(
        default=None,
        metadata={
            "name": "fontScale",
            "type": "Attribute",
            "schema_default": "100000",
            "min_inclusive": 1000,
            "max_inclusive": 100000,
        },
    )
    ln_spc_reduction: None | int = field(
        default=None,
        metadata={
            "name": "lnSpcReduction",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 13200000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextShapeAutofit(Child):
    class Meta:
        name = "CT_TextShapeAutofit"


@dataclass(slots=True, kw_only=True)
class CTTextSpacingPercent(Child):
    class Meta:
        name = "CT_TextSpacingPercent"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 13200000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextSpacingPoint(Child):
    class Meta:
        name = "CT_TextSpacingPoint"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 158400,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextUnderlineFillFollowText(Child):
    class Meta:
        name = "CT_TextUnderlineFillFollowText"


@dataclass(slots=True, kw_only=True)
class CTTextUnderlineLineFollowText(Child):
    class Meta:
        name = "CT_TextUnderlineLineFollowText"


@dataclass(slots=True, kw_only=True)
class CTTintEffect(Child):
    class Meta:
        name = "CT_TintEffect"

    hue: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    amt: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTransformEffect(Child):
    class Meta:
        name = "CT_TransformEffect"

    sx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
        },
    )
    sy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
        },
    )
    kx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    ky: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    tx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    ty: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTVector3D(Child):
    class Meta:
        name = "CT_Vector3D"

    dx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    dy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    dz: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CtFontCollectionFont(Child):
    class Meta:
        global_type = False

    script: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    typeface: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class GraphicData(Child):
    class Meta:
        name = "CT_GraphicalObjectData"

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


class STBlackWhiteMode(Enum):
    CLR = "clr"
    AUTO = "auto"
    GRAY = "gray"
    LT_GRAY = "ltGray"
    INV_GRAY = "invGray"
    GRAY_WHITE = "grayWhite"
    BLACK_GRAY = "blackGray"
    BLACK_WHITE = "blackWhite"
    BLACK = "black"
    WHITE = "white"
    HIDDEN = "hidden"


class STBlendMode(Enum):
    OVER = "over"
    MULT = "mult"
    SCREEN = "screen"
    DARKEN = "darken"
    LIGHTEN = "lighten"


class STBlipCompression(Enum):
    EMAIL = "email"
    SCREEN = "screen"
    PRINT = "print"
    HQPRINT = "hqprint"
    NONE = "none"


class STChartBuildStep(Enum):
    CATEGORY = "category"
    PT_IN_CATEGORY = "ptInCategory"
    SERIES = "series"
    PT_IN_SERIES = "ptInSeries"
    ALL_PTS = "allPts"
    GRID_LEGEND = "gridLegend"


class STColorSchemeIndex(Enum):
    DK1 = "dk1"
    LT1 = "lt1"
    DK2 = "dk2"
    LT2 = "lt2"
    ACCENT1 = "accent1"
    ACCENT2 = "accent2"
    ACCENT3 = "accent3"
    ACCENT4 = "accent4"
    ACCENT5 = "accent5"
    ACCENT6 = "accent6"
    HLINK = "hlink"
    FOL_HLINK = "folHlink"


class STCompoundLine(Enum):
    SNG = "sng"
    DBL = "dbl"
    THICK_THIN = "thickThin"
    THIN_THICK = "thinThick"
    TRI = "tri"


class STDgmBuildStep(Enum):
    SP = "sp"
    BG = "bg"


class STEffectContainerType(Enum):
    SIB = "sib"
    TREE = "tree"


class STFontCollectionIndex(Enum):
    MAJOR = "major"
    MINOR = "minor"
    NONE = "none"


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


class STLineEndLength(Enum):
    SM = "sm"
    MED = "med"
    LG = "lg"


class STLineEndType(Enum):
    NONE = "none"
    TRIANGLE = "triangle"
    STEALTH = "stealth"
    DIAMOND = "diamond"
    OVAL = "oval"
    ARROW = "arrow"


class STLineEndWidth(Enum):
    SM = "sm"
    MED = "med"
    LG = "lg"


class STOnOffStyleType(Enum):
    ON = "on"
    OFF = "off"
    DEF = "def"


class STPathFillMode(Enum):
    NONE = "none"
    NORM = "norm"
    LIGHTEN = "lighten"
    LIGHTEN_LESS = "lightenLess"
    DARKEN = "darken"
    DARKEN_LESS = "darkenLess"


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


class STPresetColorVal(Enum):
    ALICE_BLUE = "aliceBlue"
    ANTIQUE_WHITE = "antiqueWhite"
    AQUA = "aqua"
    AQUAMARINE = "aquamarine"
    AZURE = "azure"
    BEIGE = "beige"
    BISQUE = "bisque"
    BLACK = "black"
    BLANCHED_ALMOND = "blanchedAlmond"
    BLUE = "blue"
    BLUE_VIOLET = "blueViolet"
    BROWN = "brown"
    BURLY_WOOD = "burlyWood"
    CADET_BLUE = "cadetBlue"
    CHARTREUSE = "chartreuse"
    CHOCOLATE = "chocolate"
    CORAL = "coral"
    CORNFLOWER_BLUE = "cornflowerBlue"
    CORNSILK = "cornsilk"
    CRIMSON = "crimson"
    CYAN = "cyan"
    DK_BLUE = "dkBlue"
    DK_CYAN = "dkCyan"
    DK_GOLDENROD = "dkGoldenrod"
    DK_GRAY = "dkGray"
    DK_GREEN = "dkGreen"
    DK_KHAKI = "dkKhaki"
    DK_MAGENTA = "dkMagenta"
    DK_OLIVE_GREEN = "dkOliveGreen"
    DK_ORANGE = "dkOrange"
    DK_ORCHID = "dkOrchid"
    DK_RED = "dkRed"
    DK_SALMON = "dkSalmon"
    DK_SEA_GREEN = "dkSeaGreen"
    DK_SLATE_BLUE = "dkSlateBlue"
    DK_SLATE_GRAY = "dkSlateGray"
    DK_TURQUOISE = "dkTurquoise"
    DK_VIOLET = "dkViolet"
    DEEP_PINK = "deepPink"
    DEEP_SKY_BLUE = "deepSkyBlue"
    DIM_GRAY = "dimGray"
    DODGER_BLUE = "dodgerBlue"
    FIREBRICK = "firebrick"
    FLORAL_WHITE = "floralWhite"
    FOREST_GREEN = "forestGreen"
    FUCHSIA = "fuchsia"
    GAINSBORO = "gainsboro"
    GHOST_WHITE = "ghostWhite"
    GOLD = "gold"
    GOLDENROD = "goldenrod"
    GRAY = "gray"
    GREEN = "green"
    GREEN_YELLOW = "greenYellow"
    HONEYDEW = "honeydew"
    HOT_PINK = "hotPink"
    INDIAN_RED = "indianRed"
    INDIGO = "indigo"
    IVORY = "ivory"
    KHAKI = "khaki"
    LAVENDER = "lavender"
    LAVENDER_BLUSH = "lavenderBlush"
    LAWN_GREEN = "lawnGreen"
    LEMON_CHIFFON = "lemonChiffon"
    LT_BLUE = "ltBlue"
    LT_CORAL = "ltCoral"
    LT_CYAN = "ltCyan"
    LT_GOLDENROD_YELLOW = "ltGoldenrodYellow"
    LT_GRAY = "ltGray"
    LT_GREEN = "ltGreen"
    LT_PINK = "ltPink"
    LT_SALMON = "ltSalmon"
    LT_SEA_GREEN = "ltSeaGreen"
    LT_SKY_BLUE = "ltSkyBlue"
    LT_SLATE_GRAY = "ltSlateGray"
    LT_STEEL_BLUE = "ltSteelBlue"
    LT_YELLOW = "ltYellow"
    LIME = "lime"
    LIME_GREEN = "limeGreen"
    LINEN = "linen"
    MAGENTA = "magenta"
    MAROON = "maroon"
    MED_AQUAMARINE = "medAquamarine"
    MED_BLUE = "medBlue"
    MED_ORCHID = "medOrchid"
    MED_PURPLE = "medPurple"
    MED_SEA_GREEN = "medSeaGreen"
    MED_SLATE_BLUE = "medSlateBlue"
    MED_SPRING_GREEN = "medSpringGreen"
    MED_TURQUOISE = "medTurquoise"
    MED_VIOLET_RED = "medVioletRed"
    MIDNIGHT_BLUE = "midnightBlue"
    MINT_CREAM = "mintCream"
    MISTY_ROSE = "mistyRose"
    MOCCASIN = "moccasin"
    NAVAJO_WHITE = "navajoWhite"
    NAVY = "navy"
    OLD_LACE = "oldLace"
    OLIVE = "olive"
    OLIVE_DRAB = "oliveDrab"
    ORANGE = "orange"
    ORANGE_RED = "orangeRed"
    ORCHID = "orchid"
    PALE_GOLDENROD = "paleGoldenrod"
    PALE_GREEN = "paleGreen"
    PALE_TURQUOISE = "paleTurquoise"
    PALE_VIOLET_RED = "paleVioletRed"
    PAPAYA_WHIP = "papayaWhip"
    PEACH_PUFF = "peachPuff"
    PERU = "peru"
    PINK = "pink"
    PLUM = "plum"
    POWDER_BLUE = "powderBlue"
    PURPLE = "purple"
    RED = "red"
    ROSY_BROWN = "rosyBrown"
    ROYAL_BLUE = "royalBlue"
    SADDLE_BROWN = "saddleBrown"
    SALMON = "salmon"
    SANDY_BROWN = "sandyBrown"
    SEA_GREEN = "seaGreen"
    SEA_SHELL = "seaShell"
    SIENNA = "sienna"
    SILVER = "silver"
    SKY_BLUE = "skyBlue"
    SLATE_BLUE = "slateBlue"
    SLATE_GRAY = "slateGray"
    SNOW = "snow"
    SPRING_GREEN = "springGreen"
    STEEL_BLUE = "steelBlue"
    TAN = "tan"
    TEAL = "teal"
    THISTLE = "thistle"
    TOMATO = "tomato"
    TURQUOISE = "turquoise"
    VIOLET = "violet"
    WHEAT = "wheat"
    WHITE = "white"
    WHITE_SMOKE = "whiteSmoke"
    YELLOW = "yellow"
    YELLOW_GREEN = "yellowGreen"


class STPresetLineDashVal(Enum):
    SOLID = "solid"
    DOT = "dot"
    DASH = "dash"
    LG_DASH = "lgDash"
    DASH_DOT = "dashDot"
    LG_DASH_DOT = "lgDashDot"
    LG_DASH_DOT_DOT = "lgDashDotDot"
    SYS_DASH = "sysDash"
    SYS_DOT = "sysDot"
    SYS_DASH_DOT = "sysDashDot"
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


class STPresetPatternVal(Enum):
    PCT5 = "pct5"
    PCT10 = "pct10"
    PCT20 = "pct20"
    PCT25 = "pct25"
    PCT30 = "pct30"
    PCT40 = "pct40"
    PCT50 = "pct50"
    PCT60 = "pct60"
    PCT70 = "pct70"
    PCT75 = "pct75"
    PCT80 = "pct80"
    PCT90 = "pct90"
    HORZ = "horz"
    VERT = "vert"
    LT_HORZ = "ltHorz"
    LT_VERT = "ltVert"
    DK_HORZ = "dkHorz"
    DK_VERT = "dkVert"
    NAR_HORZ = "narHorz"
    NAR_VERT = "narVert"
    DASH_HORZ = "dashHorz"
    DASH_VERT = "dashVert"
    CROSS = "cross"
    DN_DIAG = "dnDiag"
    UP_DIAG = "upDiag"
    LT_DN_DIAG = "ltDnDiag"
    LT_UP_DIAG = "ltUpDiag"
    DK_DN_DIAG = "dkDnDiag"
    DK_UP_DIAG = "dkUpDiag"
    WD_DN_DIAG = "wdDnDiag"
    WD_UP_DIAG = "wdUpDiag"
    DASH_DN_DIAG = "dashDnDiag"
    DASH_UP_DIAG = "dashUpDiag"
    DIAG_CROSS = "diagCross"
    SM_CHECK = "smCheck"
    LG_CHECK = "lgCheck"
    SM_GRID = "smGrid"
    LG_GRID = "lgGrid"
    DOT_GRID = "dotGrid"
    SM_CONFETTI = "smConfetti"
    LG_CONFETTI = "lgConfetti"
    HORZ_BRICK = "horzBrick"
    DIAG_BRICK = "diagBrick"
    SOLID_DMND = "solidDmnd"
    OPEN_DMND = "openDmnd"
    DOT_DMND = "dotDmnd"
    PLAID = "plaid"
    SPHERE = "sphere"
    WEAVE = "weave"
    DIVOT = "divot"
    SHINGLE = "shingle"
    WAVE = "wave"
    TRELLIS = "trellis"
    ZIG_ZAG = "zigZag"


class STPresetShadowVal(Enum):
    SHDW1 = "shdw1"
    SHDW2 = "shdw2"
    SHDW3 = "shdw3"
    SHDW4 = "shdw4"
    SHDW5 = "shdw5"
    SHDW6 = "shdw6"
    SHDW7 = "shdw7"
    SHDW8 = "shdw8"
    SHDW9 = "shdw9"
    SHDW10 = "shdw10"
    SHDW11 = "shdw11"
    SHDW12 = "shdw12"
    SHDW13 = "shdw13"
    SHDW14 = "shdw14"
    SHDW15 = "shdw15"
    SHDW16 = "shdw16"
    SHDW17 = "shdw17"
    SHDW18 = "shdw18"
    SHDW19 = "shdw19"
    SHDW20 = "shdw20"


class STRectAlignment(Enum):
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
    PH_CLR = "phClr"
    DK1 = "dk1"
    LT1 = "lt1"
    DK2 = "dk2"
    LT2 = "lt2"


class STShapeType(Enum):
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


class STTextAlignType(Enum):
    L = "l"
    CTR = "ctr"
    R = "r"
    JUST = "just"
    JUST_LOW = "justLow"
    DIST = "dist"
    THAI_DIST = "thaiDist"


class STTextAnchoringType(Enum):
    T = "t"
    CTR = "ctr"
    B = "b"
    JUST = "just"
    DIST = "dist"


class STTextAutonumberScheme(Enum):
    ALPHA_LC_PAREN_BOTH = "alphaLcParenBoth"
    ALPHA_UC_PAREN_BOTH = "alphaUcParenBoth"
    ALPHA_LC_PAREN_R = "alphaLcParenR"
    ALPHA_UC_PAREN_R = "alphaUcParenR"
    ALPHA_LC_PERIOD = "alphaLcPeriod"
    ALPHA_UC_PERIOD = "alphaUcPeriod"
    ARABIC_PAREN_BOTH = "arabicParenBoth"
    ARABIC_PAREN_R = "arabicParenR"
    ARABIC_PERIOD = "arabicPeriod"
    ARABIC_PLAIN = "arabicPlain"
    ROMAN_LC_PAREN_BOTH = "romanLcParenBoth"
    ROMAN_UC_PAREN_BOTH = "romanUcParenBoth"
    ROMAN_LC_PAREN_R = "romanLcParenR"
    ROMAN_UC_PAREN_R = "romanUcParenR"
    ROMAN_LC_PERIOD = "romanLcPeriod"
    ROMAN_UC_PERIOD = "romanUcPeriod"
    CIRCLE_NUM_DB_PLAIN = "circleNumDbPlain"
    CIRCLE_NUM_WD_BLACK_PLAIN = "circleNumWdBlackPlain"
    CIRCLE_NUM_WD_WHITE_PLAIN = "circleNumWdWhitePlain"
    ARABIC_DB_PERIOD = "arabicDbPeriod"
    ARABIC_DB_PLAIN = "arabicDbPlain"
    EA1_CHS_PERIOD = "ea1ChsPeriod"
    EA1_CHS_PLAIN = "ea1ChsPlain"
    EA1_CHT_PERIOD = "ea1ChtPeriod"
    EA1_CHT_PLAIN = "ea1ChtPlain"
    EA1_JPN_CHS_DB_PERIOD = "ea1JpnChsDbPeriod"
    EA1_JPN_KOR_PLAIN = "ea1JpnKorPlain"
    EA1_JPN_KOR_PERIOD = "ea1JpnKorPeriod"
    ARABIC1_MINUS = "arabic1Minus"
    ARABIC2_MINUS = "arabic2Minus"
    HEBREW2_MINUS = "hebrew2Minus"
    THAI_ALPHA_PERIOD = "thaiAlphaPeriod"
    THAI_ALPHA_PAREN_R = "thaiAlphaParenR"
    THAI_ALPHA_PAREN_BOTH = "thaiAlphaParenBoth"
    THAI_NUM_PERIOD = "thaiNumPeriod"
    THAI_NUM_PAREN_R = "thaiNumParenR"
    THAI_NUM_PAREN_BOTH = "thaiNumParenBoth"
    HINDI_ALPHA_PERIOD = "hindiAlphaPeriod"
    HINDI_NUM_PERIOD = "hindiNumPeriod"
    HINDI_NUM_PAREN_R = "hindiNumParenR"
    HINDI_ALPHA1_PERIOD = "hindiAlpha1Period"


class STTextCapsType(Enum):
    NONE = "none"
    SMALL = "small"
    ALL = "all"


class STTextFontAlignType(Enum):
    AUTO = "auto"
    T = "t"
    CTR = "ctr"
    BASE = "base"
    B = "b"


class STTextHorzOverflowType(Enum):
    OVERFLOW = "overflow"
    CLIP = "clip"


class STTextShapeType(Enum):
    TEXT_NO_SHAPE = "textNoShape"
    TEXT_PLAIN = "textPlain"
    TEXT_STOP = "textStop"
    TEXT_TRIANGLE = "textTriangle"
    TEXT_TRIANGLE_INVERTED = "textTriangleInverted"
    TEXT_CHEVRON = "textChevron"
    TEXT_CHEVRON_INVERTED = "textChevronInverted"
    TEXT_RING_INSIDE = "textRingInside"
    TEXT_RING_OUTSIDE = "textRingOutside"
    TEXT_ARCH_UP = "textArchUp"
    TEXT_ARCH_DOWN = "textArchDown"
    TEXT_CIRCLE = "textCircle"
    TEXT_BUTTON = "textButton"
    TEXT_ARCH_UP_POUR = "textArchUpPour"
    TEXT_ARCH_DOWN_POUR = "textArchDownPour"
    TEXT_CIRCLE_POUR = "textCirclePour"
    TEXT_BUTTON_POUR = "textButtonPour"
    TEXT_CURVE_UP = "textCurveUp"
    TEXT_CURVE_DOWN = "textCurveDown"
    TEXT_CAN_UP = "textCanUp"
    TEXT_CAN_DOWN = "textCanDown"
    TEXT_WAVE1 = "textWave1"
    TEXT_WAVE2 = "textWave2"
    TEXT_DOUBLE_WAVE1 = "textDoubleWave1"
    TEXT_WAVE4 = "textWave4"
    TEXT_INFLATE = "textInflate"
    TEXT_DEFLATE = "textDeflate"
    TEXT_INFLATE_BOTTOM = "textInflateBottom"
    TEXT_DEFLATE_BOTTOM = "textDeflateBottom"
    TEXT_INFLATE_TOP = "textInflateTop"
    TEXT_DEFLATE_TOP = "textDeflateTop"
    TEXT_DEFLATE_INFLATE = "textDeflateInflate"
    TEXT_DEFLATE_INFLATE_DEFLATE = "textDeflateInflateDeflate"
    TEXT_FADE_RIGHT = "textFadeRight"
    TEXT_FADE_LEFT = "textFadeLeft"
    TEXT_FADE_UP = "textFadeUp"
    TEXT_FADE_DOWN = "textFadeDown"
    TEXT_SLANT_UP = "textSlantUp"
    TEXT_SLANT_DOWN = "textSlantDown"
    TEXT_CASCADE_UP = "textCascadeUp"
    TEXT_CASCADE_DOWN = "textCascadeDown"


class STTextStrikeType(Enum):
    NO_STRIKE = "noStrike"
    SNG_STRIKE = "sngStrike"
    DBL_STRIKE = "dblStrike"


class STTextTabAlignType(Enum):
    L = "l"
    CTR = "ctr"
    R = "r"
    DEC = "dec"


class STTextUnderlineType(Enum):
    NONE = "none"
    WORDS = "words"
    SNG = "sng"
    DBL = "dbl"
    HEAVY = "heavy"
    DOTTED = "dotted"
    DOTTED_HEAVY = "dottedHeavy"
    DASH = "dash"
    DASH_HEAVY = "dashHeavy"
    DASH_LONG = "dashLong"
    DASH_LONG_HEAVY = "dashLongHeavy"
    DOT_DASH = "dotDash"
    DOT_DASH_HEAVY = "dotDashHeavy"
    DOT_DOT_DASH = "dotDotDash"
    DOT_DOT_DASH_HEAVY = "dotDotDashHeavy"
    WAVY = "wavy"
    WAVY_HEAVY = "wavyHeavy"
    WAVY_DBL = "wavyDbl"


class STTextVertOverflowType(Enum):
    OVERFLOW = "overflow"
    ELLIPSIS = "ellipsis"
    CLIP = "clip"


class STTextVerticalType(Enum):
    HORZ = "horz"
    VERT = "vert"
    VERT270 = "vert270"
    WORD_ART_VERT = "wordArtVert"
    EA_VERT = "eaVert"
    MONGOLIAN_VERT = "mongolianVert"
    WORD_ART_VERT_RTL = "wordArtVertRtl"


class STTextWrappingType(Enum):
    NONE = "none"
    SQUARE = "square"


class STTileFlipMode(Enum):
    NONE = "none"
    X = "x"
    Y = "y"
    XY = "xy"


class StAnimationChartBuildType(Enum):
    ALL_AT_ONCE = "allAtOnce"
    SERIES = "series"
    CATEGORY = "category"
    SERIES_EL = "seriesEl"
    CATEGORY_EL = "categoryEl"


class StAnimationDgmBuildType(Enum):
    ALL_AT_ONCE = "allAtOnce"
    ONE = "one"
    LVL_ONE = "lvlOne"
    LVL_AT_ONCE = "lvlAtOnce"


class StSystemColorVal(Enum):
    SCROLL_BAR = "scrollBar"
    BACKGROUND = "background"
    ACTIVE_CAPTION = "activeCaption"
    INACTIVE_CAPTION = "inactiveCaption"
    MENU = "menu"
    WINDOW = "window"
    WINDOW_FRAME = "windowFrame"
    MENU_TEXT = "menuText"
    WINDOW_TEXT = "windowText"
    CAPTION_TEXT = "captionText"
    ACTIVE_BORDER = "activeBorder"
    INACTIVE_BORDER = "inactiveBorder"
    APP_WORKSPACE = "appWorkspace"
    HIGHLIGHT = "highlight"
    HIGHLIGHT_TEXT = "highlightText"
    BTN_FACE = "btnFace"
    BTN_SHADOW = "btnShadow"
    GRAY_TEXT = "grayText"
    BTN_TEXT = "btnText"
    INACTIVE_CAPTION_TEXT = "inactiveCaptionText"
    BTN_HIGHLIGHT = "btnHighlight"
    VALUE_3D_DK_SHADOW = "3dDkShadow"
    VALUE_3D_LIGHT = "3dLight"
    INFO_TEXT = "infoText"
    INFO_BK = "infoBk"
    HOT_LIGHT = "hotLight"
    GRADIENT_ACTIVE_CAPTION = "gradientActiveCaption"
    GRADIENT_INACTIVE_CAPTION = "gradientInactiveCaption"
    MENU_HIGHLIGHT = "menuHighlight"
    MENU_BAR = "menuBar"


@dataclass(slots=True, kw_only=True)
class TextFont(Child):
    class Meta:
        name = "CT_TextFont"

    typeface: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    panose: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    pitch_family: None | int = field(
        default=None,
        metadata={
            "name": "pitchFamily",
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    charset: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "1",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimationChartBuildProperties(Child):
    class Meta:
        name = "CT_AnimationChartBuildProperties"

    bld: None | StAnimationChartBuildType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": StAnimationChartBuildType.ALL_AT_ONCE,
        },
    )
    anim_bg: None | bool = field(
        default=None,
        metadata={
            "name": "animBg",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimationChartElement(Child):
    class Meta:
        name = "CT_AnimationChartElement"

    series_idx: None | int = field(
        default=None,
        metadata={
            "name": "seriesIdx",
            "type": "Attribute",
            "schema_default": "-1",
        },
    )
    category_idx: None | int = field(
        default=None,
        metadata={
            "name": "categoryIdx",
            "type": "Attribute",
            "schema_default": "-1",
        },
    )
    bld_step: None | STChartBuildStep = field(
        default=None,
        metadata={
            "name": "bldStep",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimationDgmBuildProperties(Child):
    class Meta:
        name = "CT_AnimationDgmBuildProperties"

    bld: None | StAnimationDgmBuildType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": StAnimationDgmBuildType.ALL_AT_ONCE,
        },
    )
    rev: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimationDgmElement(Child):
    class Meta:
        name = "CT_AnimationDgmElement"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "{00000000-0000-0000-0000-000000000000}",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    bld_step: None | STDgmBuildStep = field(
        default=None,
        metadata={
            "name": "bldStep",
            "type": "Attribute",
            "schema_default": STDgmBuildStep.SP,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBevel(Child):
    class Meta:
        name = "CT_Bevel"

    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "76200",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "76200",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    prst: None | STBevelPresetType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STBevelPresetType.CIRCLE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCamera(Child):
    class Meta:
        name = "CT_Camera"

    rot: None | CTSphereCoords = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prst: None | STPresetCameraType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fov: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 10800000,
        },
    )
    zoom: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
            "min_inclusive": 0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConnectionSite(Child):
    class Meta:
        name = "CT_ConnectionSite"

    pos: None | CTAdjPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ang: None | int | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDashStopList(Child):
    class Meta:
        name = "CT_DashStopList"

    ds: list[CTDashStop] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGeomGuideList(Child):
    class Meta:
        name = "CT_GeomGuideList"

    gd: list[CTGeomGuide] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupTransform2D(Child):
    class Meta:
        name = "CT_GroupTransform2D"

    off: None | CTPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ch_off: None | CTPoint2D = field(
        default=None,
        metadata={
            "name": "chOff",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ch_ext: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "name": "chExt",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    rot: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    flip_h: None | bool = field(
        default=None,
        metadata={
            "name": "flipH",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    flip_v: None | bool = field(
        default=None,
        metadata={
            "name": "flipV",
            "type": "Attribute",
            "schema_default": "false",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    rig: None | STLightRigType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dir: None | STLightRigDirection = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineEndProperties(Child):
    class Meta:
        name = "CT_LineEndProperties"

    type_value: None | STLineEndType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    w: None | STLineEndWidth = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    len: None | STLineEndLength = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOfficeArtExtensionList(Child):
    class Meta:
        name = "CT_OfficeArtExtensionList"

    ext: list[CTOfficeArtExtension] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2DCubicBezierTo(Child):
    class Meta:
        name = "CT_Path2DCubicBezierTo"

    pt: list[CTAdjPoint2D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "min_occurs": 3,
            "max_occurs": 3,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2DLineTo(Child):
    class Meta:
        name = "CT_Path2DLineTo"

    pt: None | CTAdjPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2DMoveTo(Child):
    class Meta:
        name = "CT_Path2DMoveTo"

    pt: None | CTAdjPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2DQuadBezierTo(Child):
    class Meta:
        name = "CT_Path2DQuadBezierTo"

    pt: list[CTAdjPoint2D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "min_occurs": 2,
            "max_occurs": 2,
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    path: None | STPathShadeType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPolarAdjustHandle(Child):
    class Meta:
        name = "CT_PolarAdjustHandle"

    pos: None | CTAdjPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    gd_ref_r: None | str = field(
        default=None,
        metadata={
            "name": "gdRefR",
            "type": "Attribute",
        },
    )
    min_r: None | int | str = field(
        default=None,
        metadata={
            "name": "minR",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    max_r: None | int | str = field(
        default=None,
        metadata={
            "name": "maxR",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    gd_ref_ang: None | str = field(
        default=None,
        metadata={
            "name": "gdRefAng",
            "type": "Attribute",
        },
    )
    min_ang: None | int | str = field(
        default=None,
        metadata={
            "name": "minAng",
            "type": "Attribute",
        },
    )
    max_ang: None | int | str = field(
        default=None,
        metadata={
            "name": "maxAng",
            "type": "Attribute",
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
        },
    )


@dataclass(slots=True, kw_only=True)
class CTReflectionEffect(Child):
    class Meta:
        name = "CT_ReflectionEffect"

    blur_rad: None | int = field(
        default=None,
        metadata={
            "name": "blurRad",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    st_a: None | int = field(
        default=None,
        metadata={
            "name": "stA",
            "type": "Attribute",
            "schema_default": "100000",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    st_pos: None | int = field(
        default=None,
        metadata={
            "name": "stPos",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    end_a: None | int = field(
        default=None,
        metadata={
            "name": "endA",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    end_pos: None | int = field(
        default=None,
        metadata={
            "name": "endPos",
            "type": "Attribute",
            "schema_default": "100000",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    dist: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dir: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    fade_dir: None | int = field(
        default=None,
        metadata={
            "name": "fadeDir",
            "type": "Attribute",
            "schema_default": "5400000",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    sx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
        },
    )
    sy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
        },
    )
    kx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    ky: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    algn: None | STRectAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STRectAlignment.B,
        },
    )
    rot_with_shape: None | bool = field(
        default=None,
        metadata={
            "name": "rotWithShape",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScale2D(Child):
    class Meta:
        name = "CT_Scale2D"

    sx: None | CTRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sy: None | CTRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStretchInfoProperties(Child):
    class Meta:
        name = "CT_StretchInfoProperties"

    fill_rect: None | CTRelativeRect = field(
        default=None,
        metadata={
            "name": "fillRect",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextAutonumberBullet(Child):
    class Meta:
        name = "CT_TextAutonumberBullet"

    type_value: None | STTextAutonumberScheme = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    start_at: None | int = field(
        default=None,
        metadata={
            "name": "startAt",
            "type": "Attribute",
            "schema_default": "1",
            "min_inclusive": 1,
            "max_inclusive": 32767,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextSpacing(Child):
    class Meta:
        name = "CT_TextSpacing"

    spc_pct_or_spc_pts: None | CTTextSpacingPercent | CTTextSpacingPoint = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "spcPct",
                    "type": ForwardRef("CTTextSpacingPercent"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "spcPts",
                    "type": ForwardRef("CTTextSpacingPoint"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextTabStop(Child):
    class Meta:
        name = "CT_TextTabStop"

    pos: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    algn: None | STTextTabAlignType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTileInfoProperties(Child):
    class Meta:
        name = "CT_TileInfoProperties"

    tx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    ty: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    sx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    sy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    flip: None | STTileFlipMode = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    algn: None | STRectAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTransform2D(Child):
    class Meta:
        name = "CT_Transform2D"

    off: None | CTPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    rot: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    flip_h: None | bool = field(
        default=None,
        metadata={
            "name": "flipH",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    flip_v: None | bool = field(
        default=None,
        metadata={
            "name": "flipV",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTXYAdjustHandle(Child):
    class Meta:
        name = "CT_XYAdjustHandle"

    pos: None | CTAdjPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    gd_ref_x: None | str = field(
        default=None,
        metadata={
            "name": "gdRefX",
            "type": "Attribute",
        },
    )
    min_x: None | int | str = field(
        default=None,
        metadata={
            "name": "minX",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    max_x: None | int | str = field(
        default=None,
        metadata={
            "name": "maxX",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    gd_ref_y: None | str = field(
        default=None,
        metadata={
            "name": "gdRefY",
            "type": "Attribute",
        },
    )
    min_y: None | int | str = field(
        default=None,
        metadata={
            "name": "minY",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    max_y: None | int | str = field(
        default=None,
        metadata={
            "name": "maxY",
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class Graphic1(Child):
    class Meta:
        name = "CT_GraphicalObject"

    graphic_data: None | GraphicData = field(
        default=None,
        metadata={
            "name": "graphicData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class AlphaMod1(CTPositivePercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class AlphaMod2(CTPositivePercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Alpha1(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Alpha2(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class BlueMod1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class BlueMod2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class BlueOff1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class BlueOff2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Blue1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Blue2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class GreenMod1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class GreenMod2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class GreenOff1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class GreenOff2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Green1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Green2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class HueMod1(CTPositivePercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class HueMod2(CTPositivePercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class LumMod1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class LumMod2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class LumOff1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class LumOff2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Lum1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Lum2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RedMod1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RedMod2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RedOff1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class RedOff2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Red1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Red2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class SatMod1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class SatMod2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class SatOff1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class SatOff2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Sat1(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Sat2(CTPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Shade1(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Shade2(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class ThemeManager(CTEmptyElement):
    class Meta:
        name = "themeManager"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(slots=True, kw_only=True)
class Tint1(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class Tint2(CTPositiveFixedPercentage):
    class Meta:
        global_type = False


@dataclass(slots=True, kw_only=True)
class CTAdjustHandleList(Child):
    class Meta:
        name = "CT_AdjustHandleList"

    ah_xy_or_ah_polar: list[CTXYAdjustHandle | CTPolarAdjustHandle] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "ahXY",
                    "type": ForwardRef("CTXYAdjustHandle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "ahPolar",
                    "type": ForwardRef("CTPolarAdjustHandle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimationElementChoice(Child):
    class Meta:
        name = "CT_AnimationElementChoice"

    dgm_or_chart: None | CTAnimationDgmElement | CTAnimationChartElement = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "dgm",
                    "type": ForwardRef("CTAnimationDgmElement"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "chart",
                    "type": ForwardRef("CTAnimationChartElement"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAnimationGraphicalObjectBuildProperties(Child):
    class Meta:
        name = "CT_AnimationGraphicalObjectBuildProperties"

    bld_dgm_or_bld_chart: (
        None | CTAnimationDgmBuildProperties | CTAnimationChartBuildProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "bldDgm",
                    "type": ForwardRef("CTAnimationDgmBuildProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "bldChart",
                    "type": ForwardRef("CTAnimationChartBuildProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAudioCD(Child):
    class Meta:
        name = "CT_AudioCD"

    st: None | CTAudioCDTime = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    end: None | CTAudioCDTime = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAudioFile(Child):
    class Meta:
        name = "CT_AudioFile"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    link: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBackdrop(Child):
    class Meta:
        name = "CT_Backdrop"

    anchor: None | CTPoint3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    norm: None | CTVector3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    up: None | CTVector3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCell3D(Child):
    class Meta:
        name = "CT_Cell3D"

    bevel: None | CTBevel = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    light_rig: None | CTLightRig = field(
        default=None,
        metadata={
            "name": "lightRig",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prst_material: None | STPresetMaterialType = field(
        default=None,
        metadata={
            "name": "prstMaterial",
            "type": "Attribute",
            "schema_default": STPresetMaterialType.PLASTIC,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorMapping(Child):
    class Meta:
        name = "CT_ColorMapping"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bg1: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    tx1: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    bg2: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    tx2: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accent1: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accent2: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accent3: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accent4: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accent5: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accent6: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    hlink: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fol_hlink: None | STColorSchemeIndex = field(
        default=None,
        metadata={
            "name": "folHlink",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConnectionSiteList(Child):
    class Meta:
        name = "CT_ConnectionSiteList"

    cxn: list[CTConnectionSite] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConnectorLocking(Child):
    class Meta:
        name = "CT_ConnectorLocking"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    no_grp: None | bool = field(
        default=None,
        metadata={
            "name": "noGrp",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_select: None | bool = field(
        default=None,
        metadata={
            "name": "noSelect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_rot: None | bool = field(
        default=None,
        metadata={
            "name": "noRot",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_aspect: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeAspect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_move: None | bool = field(
        default=None,
        metadata={
            "name": "noMove",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_resize: None | bool = field(
        default=None,
        metadata={
            "name": "noResize",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_edit_points: None | bool = field(
        default=None,
        metadata={
            "name": "noEditPoints",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_adjust_handles: None | bool = field(
        default=None,
        metadata={
            "name": "noAdjustHandles",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_arrowheads: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeArrowheads",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_shape_type: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeShapeType",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicalObjectFrameLocking(Child):
    class Meta:
        name = "CT_GraphicalObjectFrameLocking"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    no_grp: None | bool = field(
        default=None,
        metadata={
            "name": "noGrp",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_drilldown: None | bool = field(
        default=None,
        metadata={
            "name": "noDrilldown",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_select: None | bool = field(
        default=None,
        metadata={
            "name": "noSelect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_aspect: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeAspect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_move: None | bool = field(
        default=None,
        metadata={
            "name": "noMove",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_resize: None | bool = field(
        default=None,
        metadata={
            "name": "noResize",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupLocking(Child):
    class Meta:
        name = "CT_GroupLocking"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    no_grp: None | bool = field(
        default=None,
        metadata={
            "name": "noGrp",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_ungrp: None | bool = field(
        default=None,
        metadata={
            "name": "noUngrp",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_select: None | bool = field(
        default=None,
        metadata={
            "name": "noSelect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_rot: None | bool = field(
        default=None,
        metadata={
            "name": "noRot",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_aspect: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeAspect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_move: None | bool = field(
        default=None,
        metadata={
            "name": "noMove",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_resize: None | bool = field(
        default=None,
        metadata={
            "name": "noResize",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHslColor(Child):
    class Meta:
        name = "CT_HslColor"

    content: list[
        Tint1
        | Shade1
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha1
        | CTFixedPercentage
        | AlphaMod1
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod1
        | Sat1
        | SatOff1
        | SatMod1
        | Lum1
        | LumOff1
        | LumMod1
        | Red1
        | RedOff1
        | RedMod1
        | Green1
        | GreenOff1
        | GreenMod1
        | Blue1
        | BlueOff1
        | BlueMod1
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "comp",
                    "type": ForwardRef("CTComplementTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "inv",
                    "type": ForwardRef("CTInverseTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gray",
                    "type": ForwardRef("CTGrayscaleTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hue",
                    "type": ForwardRef("CTPositiveFixedAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueOff",
                    "type": ForwardRef("CTAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("HueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gamma",
                    "type": ForwardRef("CTGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "invGamma",
                    "type": ForwardRef("CTInverseGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    hue_attribute: None | int = field(
        default=None,
        metadata={
            "name": "hue",
            "type": "Attribute",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    sat_attribute: None | int = field(
        default=None,
        metadata={
            "name": "sat",
            "type": "Attribute",
        },
    )
    lum_attribute: None | int = field(
        default=None,
        metadata={
            "name": "lum",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHyperlink(Child):
    class Meta:
        name = "CT_Hyperlink"

    snd: None | CTEmbeddedWAVAudioFile = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    invalid_url: None | str = field(
        default=None,
        metadata={
            "name": "invalidUrl",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    action: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    tgt_frame: None | str = field(
        default=None,
        metadata={
            "name": "tgtFrame",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    tooltip: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    history: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    highlight_click: None | bool = field(
        default=None,
        metadata={
            "name": "highlightClick",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    end_snd: None | bool = field(
        default=None,
        metadata={
            "name": "endSnd",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2D(Child):
    class Meta:
        name = "CT_Path2D"

    content: list[
        CTPath2DClose
        | CTPath2DMoveTo
        | CTPath2DLineTo
        | CTPath2DArcTo
        | CTPath2DQuadBezierTo
        | CTPath2DCubicBezierTo
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "close",
                    "type": ForwardRef("CTPath2DClose"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "moveTo",
                    "type": ForwardRef("CTPath2DMoveTo"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lnTo",
                    "type": ForwardRef("CTPath2DLineTo"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "arcTo",
                    "type": ForwardRef("CTPath2DArcTo"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "quadBezTo",
                    "type": ForwardRef("CTPath2DQuadBezierTo"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "cubicBezTo",
                    "type": ForwardRef("CTPath2DCubicBezierTo"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    fill: None | STPathFillMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STPathFillMode.NORM,
        },
    )
    stroke: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    extrusion_ok: None | bool = field(
        default=None,
        metadata={
            "name": "extrusionOk",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureLocking(Child):
    class Meta:
        name = "CT_PictureLocking"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    no_grp: None | bool = field(
        default=None,
        metadata={
            "name": "noGrp",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_select: None | bool = field(
        default=None,
        metadata={
            "name": "noSelect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_rot: None | bool = field(
        default=None,
        metadata={
            "name": "noRot",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_aspect: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeAspect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_move: None | bool = field(
        default=None,
        metadata={
            "name": "noMove",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_resize: None | bool = field(
        default=None,
        metadata={
            "name": "noResize",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_edit_points: None | bool = field(
        default=None,
        metadata={
            "name": "noEditPoints",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_adjust_handles: None | bool = field(
        default=None,
        metadata={
            "name": "noAdjustHandles",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_arrowheads: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeArrowheads",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_shape_type: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeShapeType",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_crop: None | bool = field(
        default=None,
        metadata={
            "name": "noCrop",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresetColor(Child):
    class Meta:
        name = "CT_PresetColor"

    content: list[
        Tint1
        | Shade1
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha1
        | CTFixedPercentage
        | AlphaMod1
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod1
        | Sat1
        | SatOff1
        | SatMod1
        | Lum1
        | LumOff1
        | LumMod1
        | Red1
        | RedOff1
        | RedMod1
        | Green1
        | GreenOff1
        | GreenMod1
        | Blue1
        | BlueOff1
        | BlueMod1
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "comp",
                    "type": ForwardRef("CTComplementTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "inv",
                    "type": ForwardRef("CTInverseTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gray",
                    "type": ForwardRef("CTGrayscaleTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hue",
                    "type": ForwardRef("CTPositiveFixedAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueOff",
                    "type": ForwardRef("CTAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("HueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gamma",
                    "type": ForwardRef("CTGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "invGamma",
                    "type": ForwardRef("CTInverseGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    val: None | STPresetColorVal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresetGeometry2D(Child):
    class Meta:
        name = "CT_PresetGeometry2D"

    av_lst: None | CTGeomGuideList = field(
        default=None,
        metadata={
            "name": "avLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prst: None | STShapeType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresetTextShape(Child):
    class Meta:
        name = "CT_PresetTextShape"

    av_lst: None | CTGeomGuideList = field(
        default=None,
        metadata={
            "name": "avLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prst: None | STTextShapeType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTQuickTimeFile(Child):
    class Meta:
        name = "CT_QuickTimeFile"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    link: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSRgbColor(Child):
    class Meta:
        name = "CT_SRgbColor"

    content: list[
        Tint1
        | Shade1
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha1
        | CTFixedPercentage
        | AlphaMod1
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod1
        | Sat1
        | SatOff1
        | SatMod1
        | Lum1
        | LumOff1
        | LumMod1
        | Red1
        | RedOff1
        | RedMod1
        | Green1
        | GreenOff1
        | GreenMod1
        | Blue1
        | BlueOff1
        | BlueMod1
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "comp",
                    "type": ForwardRef("CTComplementTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "inv",
                    "type": ForwardRef("CTInverseTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gray",
                    "type": ForwardRef("CTGrayscaleTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hue",
                    "type": ForwardRef("CTPositiveFixedAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueOff",
                    "type": ForwardRef("CTAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("HueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gamma",
                    "type": ForwardRef("CTGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "invGamma",
                    "type": ForwardRef("CTInverseGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    val: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScRgbColor(Child):
    class Meta:
        name = "CT_ScRgbColor"

    content: list[
        Tint1
        | Shade1
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha1
        | CTFixedPercentage
        | AlphaMod1
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod1
        | Sat1
        | SatOff1
        | SatMod1
        | Lum1
        | LumOff1
        | LumMod1
        | Red1
        | RedOff1
        | RedMod1
        | Green1
        | GreenOff1
        | GreenMod1
        | Blue1
        | BlueOff1
        | BlueMod1
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "comp",
                    "type": ForwardRef("CTComplementTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "inv",
                    "type": ForwardRef("CTInverseTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gray",
                    "type": ForwardRef("CTGrayscaleTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hue",
                    "type": ForwardRef("CTPositiveFixedAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueOff",
                    "type": ForwardRef("CTAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("HueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gamma",
                    "type": ForwardRef("CTGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "invGamma",
                    "type": ForwardRef("CTInverseGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    r: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    g: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    b: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSchemeColor(Child):
    class Meta:
        name = "CT_SchemeColor"

    content: list[
        Tint1
        | Shade1
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha1
        | CTFixedPercentage
        | AlphaMod1
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod1
        | Sat1
        | SatOff1
        | SatMod1
        | Lum1
        | LumOff1
        | LumMod1
        | Red1
        | RedOff1
        | RedMod1
        | Green1
        | GreenOff1
        | GreenMod1
        | Blue1
        | BlueOff1
        | BlueMod1
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "comp",
                    "type": ForwardRef("CTComplementTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "inv",
                    "type": ForwardRef("CTInverseTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gray",
                    "type": ForwardRef("CTGrayscaleTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hue",
                    "type": ForwardRef("CTPositiveFixedAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueOff",
                    "type": ForwardRef("CTAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("HueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gamma",
                    "type": ForwardRef("CTGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "invGamma",
                    "type": ForwardRef("CTInverseGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    val: None | STSchemeColorVal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeLocking(Child):
    class Meta:
        name = "CT_ShapeLocking"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    no_grp: None | bool = field(
        default=None,
        metadata={
            "name": "noGrp",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_select: None | bool = field(
        default=None,
        metadata={
            "name": "noSelect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_rot: None | bool = field(
        default=None,
        metadata={
            "name": "noRot",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_aspect: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeAspect",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_move: None | bool = field(
        default=None,
        metadata={
            "name": "noMove",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_resize: None | bool = field(
        default=None,
        metadata={
            "name": "noResize",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_edit_points: None | bool = field(
        default=None,
        metadata={
            "name": "noEditPoints",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_adjust_handles: None | bool = field(
        default=None,
        metadata={
            "name": "noAdjustHandles",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_arrowheads: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeArrowheads",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_change_shape_type: None | bool = field(
        default=None,
        metadata={
            "name": "noChangeShapeType",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    no_text_edit: None | bool = field(
        default=None,
        metadata={
            "name": "noTextEdit",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSystemColor(Child):
    class Meta:
        name = "CT_SystemColor"

    content: list[
        Tint1
        | Shade1
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha1
        | CTFixedPercentage
        | AlphaMod1
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod1
        | Sat1
        | SatOff1
        | SatMod1
        | Lum1
        | LumOff1
        | LumMod1
        | Red1
        | RedOff1
        | RedMod1
        | Green1
        | GreenOff1
        | GreenMod1
        | Blue1
        | BlueOff1
        | BlueMod1
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "comp",
                    "type": ForwardRef("CTComplementTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "inv",
                    "type": ForwardRef("CTInverseTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gray",
                    "type": ForwardRef("CTGrayscaleTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alpha",
                    "type": ForwardRef("Alpha1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hue",
                    "type": ForwardRef("CTPositiveFixedAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueOff",
                    "type": ForwardRef("CTAngle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hueMod",
                    "type": ForwardRef("HueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod1"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gamma",
                    "type": ForwardRef("CTGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "invGamma",
                    "type": ForwardRef("CTInverseGammaTransform"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    val: None | StSystemColorVal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    last_clr: None | bytes = field(
        default=None,
        metadata={
            "name": "lastClr",
            "type": "Attribute",
            "length": 3,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableCol(Child):
    class Meta:
        name = "CT_TableCol"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextTabStopList(Child):
    class Meta:
        name = "CT_TextTabStopList"

    tab: list[CTTextTabStop] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "max_occurs": 32,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTVideoFile(Child):
    class Meta:
        name = "CT_VideoFile"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    link: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class FontCollection(Child):
    class Meta:
        name = "CT_FontCollection"

    latin: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ea: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    cs: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    font: list[CtFontCollectionFont] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Graphic(Graphic1):
    class Meta:
        name = "graphic"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTAlphaInverseEffect(Child):
    class Meta:
        name = "CT_AlphaInverseEffect"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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


@dataclass(slots=True, kw_only=True)
class CTColor(Child):
    class Meta:
        name = "CT_Color"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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


@dataclass(slots=True, kw_only=True)
class CTColorMRU(Child):
    class Meta:
        name = "CT_ColorMRU"

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
                    "max_occurs": 10,
                },
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 10,
                },
                {
                    "name": "hslClr",
                    "type": ForwardRef("CTHslColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 10,
                },
                {
                    "name": "sysClr",
                    "type": ForwardRef("CTSystemColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 10,
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 10,
                },
                {
                    "name": "prstClr",
                    "type": ForwardRef("CTPresetColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 10,
                },
            ),
            "max_occurs": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorMappingOverride(Child):
    class Meta:
        name = "CT_ColorMappingOverride"

    master_clr_mapping_or_override_clr_mapping: None | CTEmptyElement | CTColorMapping = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "masterClrMapping",
                    "type": ForwardRef("CTEmptyElement"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "overrideClrMapping",
                    "type": ForwardRef("CTColorMapping"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorReplaceEffect(Child):
    class Meta:
        name = "CT_ColorReplaceEffect"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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


@dataclass(slots=True, kw_only=True)
class CTCustomColor(Child):
    class Meta:
        name = "CT_CustomColor"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDuotoneEffect(Child):
    class Meta:
        name = "CT_DuotoneEffect"

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
                    "max_occurs": 2,
                },
                {
                    "name": "srgbClr",
                    "type": ForwardRef("CTSRgbColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 2,
                },
                {
                    "name": "hslClr",
                    "type": ForwardRef("CTHslColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 2,
                },
                {
                    "name": "sysClr",
                    "type": ForwardRef("CTSystemColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 2,
                },
                {
                    "name": "schemeClr",
                    "type": ForwardRef("CTSchemeColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 2,
                },
                {
                    "name": "prstClr",
                    "type": ForwardRef("CTPresetColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "max_occurs": 2,
                },
            ),
            "max_occurs": 2,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFontReference(Child):
    class Meta:
        name = "CT_FontReference"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    idx: None | STFontCollectionIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGlowEffect(Child):
    class Meta:
        name = "CT_GlowEffect"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGradientStop(Child):
    class Meta:
        name = "CT_GradientStop"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    pos: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInnerShadowEffect(Child):
    class Meta:
        name = "CT_InnerShadowEffect"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    blur_rad: None | int = field(
        default=None,
        metadata={
            "name": "blurRad",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dist: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dir: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNonVisualConnectorProperties(Child):
    class Meta:
        name = "CT_NonVisualConnectorProperties"

    cxn_sp_locks: None | CTConnectorLocking = field(
        default=None,
        metadata={
            "name": "cxnSpLocks",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    st_cxn: None | CTConnection = field(
        default=None,
        metadata={
            "name": "stCxn",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    end_cxn: None | CTConnection = field(
        default=None,
        metadata={
            "name": "endCxn",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNonVisualDrawingProps(Child):
    class Meta:
        name = "CT_NonVisualDrawingProps"

    hlink_click: None | CTHyperlink = field(
        default=None,
        metadata={
            "name": "hlinkClick",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    hlink_hover: None | CTHyperlink = field(
        default=None,
        metadata={
            "name": "hlinkHover",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    descr: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    hidden: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNonVisualDrawingShapeProps(Child):
    class Meta:
        name = "CT_NonVisualDrawingShapeProps"

    sp_locks: None | CTShapeLocking = field(
        default=None,
        metadata={
            "name": "spLocks",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tx_box: None | bool = field(
        default=None,
        metadata={
            "name": "txBox",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNonVisualGraphicFrameProperties(Child):
    class Meta:
        name = "CT_NonVisualGraphicFrameProperties"

    graphic_frame_locks: None | CTGraphicalObjectFrameLocking = field(
        default=None,
        metadata={
            "name": "graphicFrameLocks",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNonVisualGroupDrawingShapeProps(Child):
    class Meta:
        name = "CT_NonVisualGroupDrawingShapeProps"

    grp_sp_locks: None | CTGroupLocking = field(
        default=None,
        metadata={
            "name": "grpSpLocks",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNonVisualPictureProperties(Child):
    class Meta:
        name = "CT_NonVisualPictureProperties"

    pic_locks: None | CTPictureLocking = field(
        default=None,
        metadata={
            "name": "picLocks",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prefer_relative_resize: None | bool = field(
        default=None,
        metadata={
            "name": "preferRelativeResize",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOuterShadowEffect(Child):
    class Meta:
        name = "CT_OuterShadowEffect"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    blur_rad: None | int = field(
        default=None,
        metadata={
            "name": "blurRad",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dist: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dir: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
        },
    )
    sx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
        },
    )
    sy: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100000",
        },
    )
    kx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    ky: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_exclusive": -5400000,
            "max_exclusive": 5400000,
        },
    )
    algn: None | STRectAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STRectAlignment.B,
        },
    )
    rot_with_shape: None | bool = field(
        default=None,
        metadata={
            "name": "rotWithShape",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath2DList(Child):
    class Meta:
        name = "CT_Path2DList"

    path: list[CTPath2D] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPresetShadowEffect(Child):
    class Meta:
        name = "CT_PresetShadowEffect"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    prst: None | STPresetShadowVal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dist: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    dir: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_exclusive": 21600000,
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    light_rig: None | CTLightRig = field(
        default=None,
        metadata={
            "name": "lightRig",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    backdrop: None | CTBackdrop = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSolidColorFillProperties(Child):
    class Meta:
        name = "CT_SolidColorFillProperties"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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


@dataclass(slots=True, kw_only=True)
class CTStyleMatrixReference(Child):
    class Meta:
        name = "CT_StyleMatrixReference"

    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableGrid(Child):
    class Meta:
        name = "CT_TableGrid"

    grid_col: list[CTTableCol] = field(
        default_factory=ChildList,
        metadata={
            "name": "gridCol",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtBaseStylesOverrideFontScheme(Child):
    class Meta:
        global_type = False

    major_font: None | FontCollection = field(
        default=None,
        metadata={
            "name": "majorFont",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    minor_font: None | FontCollection = field(
        default=None,
        metadata={
            "name": "minorFont",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CtBaseStylesFontScheme(Child):
    class Meta:
        global_type = False

    major_font: None | FontCollection = field(
        default=None,
        metadata={
            "name": "majorFont",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    minor_font: None | FontCollection = field(
        default=None,
        metadata={
            "name": "minorFont",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorChangeEffect(Child):
    class Meta:
        name = "CT_ColorChangeEffect"

    clr_from: None | CTColor = field(
        default=None,
        metadata={
            "name": "clrFrom",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    clr_to: None | CTColor = field(
        default=None,
        metadata={
            "name": "clrTo",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    use_a: None | bool = field(
        default=None,
        metadata={
            "name": "useA",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorScheme(Child):
    class Meta:
        name = "CT_ColorScheme"

    dk1: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lt1: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    dk2: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lt2: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    accent1: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    accent2: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    accent3: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    accent4: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    accent5: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    accent6: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    hlink: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    fol_hlink: None | CTColor = field(
        default=None,
        metadata={
            "name": "folHlink",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomColorList(Child):
    class Meta:
        name = "CT_CustomColorList"

    cust_clr: list[CTCustomColor] = field(
        default_factory=ChildList,
        metadata={
            "name": "custClr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCustomGeometry2D(Child):
    class Meta:
        name = "CT_CustomGeometry2D"

    av_lst: None | CTGeomGuideList = field(
        default=None,
        metadata={
            "name": "avLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    gd_lst: None | CTGeomGuideList = field(
        default=None,
        metadata={
            "name": "gdLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ah_lst: None | CTAdjustHandleList = field(
        default=None,
        metadata={
            "name": "ahLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    cxn_lst: None | CTConnectionSiteList = field(
        default=None,
        metadata={
            "name": "cxnLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    rect: None | CTGeomRect = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    path_lst: None | CTPath2DList = field(
        default=None,
        metadata={
            "name": "pathLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGradientStopList(Child):
    class Meta:
        name = "CT_GradientStopList"

    gs: list[CTGradientStop] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "min_occurs": 2,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlConnectorNonVisual(Child):
    class Meta:
        name = "CT_GvmlConnectorNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    c_nv_cxn_sp_pr: None | CTNonVisualConnectorProperties = field(
        default=None,
        metadata={
            "name": "cNvCxnSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlGraphicFrameNonVisual(Child):
    class Meta:
        name = "CT_GvmlGraphicFrameNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    c_nv_graphic_frame_pr: None | CTNonVisualGraphicFrameProperties = field(
        default=None,
        metadata={
            "name": "cNvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlGroupShapeNonVisual(Child):
    class Meta:
        name = "CT_GvmlGroupShapeNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    c_nv_grp_sp_pr: None | CTNonVisualGroupDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlPictureNonVisual(Child):
    class Meta:
        name = "CT_GvmlPictureNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    c_nv_pic_pr: None | CTNonVisualPictureProperties = field(
        default=None,
        metadata={
            "name": "cNvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlShapeNonVisual(Child):
    class Meta:
        name = "CT_GvmlShapeNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    c_nv_sp_pr: None | CTNonVisualDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPatternFillProperties(Child):
    class Meta:
        name = "CT_PatternFillProperties"

    fg_clr: None | CTColor = field(
        default=None,
        metadata={
            "name": "fgClr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bg_clr: None | CTColor = field(
        default=None,
        metadata={
            "name": "bgClr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prst: None | STPresetPatternVal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShape3D(Child):
    class Meta:
        name = "CT_Shape3D"

    bevel_t: None | CTBevel = field(
        default=None,
        metadata={
            "name": "bevelT",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bevel_b: None | CTBevel = field(
        default=None,
        metadata={
            "name": "bevelB",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    extrusion_clr: None | CTColor = field(
        default=None,
        metadata={
            "name": "extrusionClr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    contour_clr: None | CTColor = field(
        default=None,
        metadata={
            "name": "contourClr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    z: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    extrusion_h: None | int = field(
        default=None,
        metadata={
            "name": "extrusionH",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    contour_w: None | int = field(
        default=None,
        metadata={
            "name": "contourW",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    prst_material: None | STPresetMaterialType = field(
        default=None,
        metadata={
            "name": "prstMaterial",
            "type": "Attribute",
            "schema_default": STPresetMaterialType.WARM_MATTE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeStyle(Child):
    class Meta:
        name = "CT_ShapeStyle"

    ln_ref: None | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "name": "lnRef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    fill_ref: None | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "name": "fillRef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    effect_ref: None | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "name": "effectRef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    font_ref: None | CTFontReference = field(
        default=None,
        metadata={
            "name": "fontRef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableStyleTextStyle(Child):
    class Meta:
        name = "CT_TableStyleTextStyle"

    font_or_font_ref: None | FontCollection | CTFontReference = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "font",
                    "type": ForwardRef("FontCollection"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fontRef",
                    "type": ForwardRef("CTFontReference"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    content: (
        None
        | CTScRgbColor
        | CTSRgbColor
        | CTHslColor
        | CTSystemColor
        | CTSchemeColor
        | CTPresetColor
    ) = field(
        default=None,
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
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    b: None | STOnOffStyleType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STOnOffStyleType.DEF,
        },
    )
    i: None | STOnOffStyleType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STOnOffStyleType.DEF,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBlip(Child):
    class Meta:
        name = "CT_Blip"

    content: list[
        CTAlphaBiLevelEffect
        | CTAlphaCeilingEffect
        | CTAlphaFloorEffect
        | CTAlphaInverseEffect
        | CTAlphaModulateEffect
        | CTAlphaModulateFixedEffect
        | CTAlphaReplaceEffect
        | CTBiLevelEffect
        | CTBlurEffect
        | CTColorChangeEffect
        | CTColorReplaceEffect
        | CTDuotoneEffect
        | CTFillOverlayEffect
        | CTGrayscaleEffect
        | CTHSLEffect
        | CTLuminanceEffect
        | CTTintEffect
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "alphaBiLevel",
                    "type": ForwardRef("CTAlphaBiLevelEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaCeiling",
                    "type": ForwardRef("CTAlphaCeilingEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaFloor",
                    "type": ForwardRef("CTAlphaFloorEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaInv",
                    "type": ForwardRef("CTAlphaInverseEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("CTAlphaModulateEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaModFix",
                    "type": ForwardRef("CTAlphaModulateFixedEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaRepl",
                    "type": ForwardRef("CTAlphaReplaceEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "biLevel",
                    "type": ForwardRef("CTBiLevelEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blur",
                    "type": ForwardRef("CTBlurEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "clrChange",
                    "type": ForwardRef("CTColorChangeEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "clrRepl",
                    "type": ForwardRef("CTColorReplaceEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "duotone",
                    "type": ForwardRef("CTDuotoneEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fillOverlay",
                    "type": ForwardRef("CTFillOverlayEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grayscl",
                    "type": ForwardRef("CTGrayscaleEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hsl",
                    "type": ForwardRef("CTHSLEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("CTLuminanceEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "tint",
                    "type": ForwardRef("CTTintEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    embed: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    link: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    cstate: None | STBlipCompression = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STBlipCompression.NONE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorSchemeAndMapping(Child):
    class Meta:
        name = "CT_ColorSchemeAndMapping"

    clr_scheme: None | CTColorScheme = field(
        default=None,
        metadata={
            "name": "clrScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    clr_map: None | CTColorMapping = field(
        default=None,
        metadata={
            "name": "clrMap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGradientFillProperties(Child):
    class Meta:
        name = "CT_GradientFillProperties"

    gs_lst: None | CTGradientStopList = field(
        default=None,
        metadata={
            "name": "gsLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
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
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "path",
                    "type": ForwardRef("CTPathShadeProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    tile_rect: None | CTRelativeRect = field(
        default=None,
        metadata={
            "name": "tileRect",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    flip: None | STTileFlipMode = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    rot_with_shape: None | bool = field(
        default=None,
        metadata={
            "name": "rotWithShape",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlGraphicalObjectFrame(Child):
    class Meta:
        name = "CT_GvmlGraphicalObjectFrame"

    nv_graphic_frame_pr: None | CTGvmlGraphicFrameNonVisual = field(
        default=None,
        metadata={
            "name": "nvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    graphic: None | Graphic = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBodyProperties(Child):
    class Meta:
        name = "CT_TextBodyProperties"

    prst_tx_warp: None | CTPresetTextShape = field(
        default=None,
        metadata={
            "name": "prstTxWarp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    no_autofit_or_norm_autofit_or_sp_auto_fit: (
        None | CTTextNoAutofit | CTTextNormalAutofit | CTTextShapeAutofit
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noAutofit",
                    "type": ForwardRef("CTTextNoAutofit"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "normAutofit",
                    "type": ForwardRef("CTTextNormalAutofit"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "spAutoFit",
                    "type": ForwardRef("CTTextShapeAutofit"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
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
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    rot: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spc_first_last_para: None | bool = field(
        default=None,
        metadata={
            "name": "spcFirstLastPara",
            "type": "Attribute",
        },
    )
    vert_overflow: None | STTextVertOverflowType = field(
        default=None,
        metadata={
            "name": "vertOverflow",
            "type": "Attribute",
        },
    )
    horz_overflow: None | STTextHorzOverflowType = field(
        default=None,
        metadata={
            "name": "horzOverflow",
            "type": "Attribute",
        },
    )
    vert: None | STTextVerticalType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrap: None | STTextWrappingType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    l_ins: None | int = field(
        default=None,
        metadata={
            "name": "lIns",
            "type": "Attribute",
        },
    )
    t_ins: None | int = field(
        default=None,
        metadata={
            "name": "tIns",
            "type": "Attribute",
        },
    )
    r_ins: None | int = field(
        default=None,
        metadata={
            "name": "rIns",
            "type": "Attribute",
        },
    )
    b_ins: None | int = field(
        default=None,
        metadata={
            "name": "bIns",
            "type": "Attribute",
        },
    )
    num_col: None | int = field(
        default=None,
        metadata={
            "name": "numCol",
            "type": "Attribute",
            "min_inclusive": 1,
            "max_inclusive": 16,
        },
    )
    spc_col: None | int = field(
        default=None,
        metadata={
            "name": "spcCol",
            "type": "Attribute",
            "min_inclusive": 0,
        },
    )
    rtl_col: None | bool = field(
        default=None,
        metadata={
            "name": "rtlCol",
            "type": "Attribute",
        },
    )
    from_word_art: None | bool = field(
        default=None,
        metadata={
            "name": "fromWordArt",
            "type": "Attribute",
        },
    )
    anchor: None | STTextAnchoringType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    anchor_ctr: None | bool = field(
        default=None,
        metadata={
            "name": "anchorCtr",
            "type": "Attribute",
        },
    )
    force_aa: None | bool = field(
        default=None,
        metadata={
            "name": "forceAA",
            "type": "Attribute",
        },
    )
    upright: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    compat_ln_spc: None | bool = field(
        default=None,
        metadata={
            "name": "compatLnSpc",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBlipFillProperties(Child):
    class Meta:
        name = "CT_BlipFillProperties"

    blip: None | CTBlip = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    src_rect: None | CTRelativeRect = field(
        default=None,
        metadata={
            "name": "srcRect",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tile_or_stretch: None | CTTileInfoProperties | CTStretchInfoProperties = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tile",
                    "type": ForwardRef("CTTileInfoProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "stretch",
                    "type": ForwardRef("CTStretchInfoProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    dpi: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    rot_with_shape: None | bool = field(
        default=None,
        metadata={
            "name": "rotWithShape",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorSchemeList(Child):
    class Meta:
        name = "CT_ColorSchemeList"

    extra_clr_scheme: list[CTColorSchemeAndMapping] = field(
        default_factory=ChildList,
        metadata={
            "name": "extraClrScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineProperties(Child):
    class Meta:
        name = "CT_LineProperties"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTPatternFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    prst_dash_or_cust_dash: None | CTPresetLineDashProperties | CTDashStopList = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "prstDash",
                    "type": ForwardRef("CTPresetLineDashProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "custDash",
                    "type": ForwardRef("CTDashStopList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    round_or_bevel_or_miter: (
        None | CTLineJoinRound | CTLineJoinBevel | CTLineJoinMiterProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "round",
                    "type": ForwardRef("CTLineJoinRound"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "bevel",
                    "type": ForwardRef("CTLineJoinBevel"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "miter",
                    "type": ForwardRef("CTLineJoinMiterProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    head_end: None | CTLineEndProperties = field(
        default=None,
        metadata={
            "name": "headEnd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tail_end: None | CTLineEndProperties = field(
        default=None,
        metadata={
            "name": "tailEnd",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    w: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 20116800,
        },
    )
    cap: None | STLineCap = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cmpd: None | STCompoundLine = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    algn: None | STPenAlignment = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBlipBullet(Child):
    class Meta:
        name = "CT_TextBlipBullet"

    blip: None | CTBlip = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Blip(CTBlip):
    class Meta:
        name = "blip"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTBackgroundFillStyleList(Child):
    class Meta:
        name = "CT_BackgroundFillStyleList"

    content: list[
        CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFillEffect(Child):
    class Meta:
        name = "CT_FillEffect"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFillOverlayEffect(Child):
    class Meta:
        name = "CT_FillOverlayEffect"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    blend: None | STBlendMode = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFillProperties(Child):
    class Meta:
        name = "CT_FillProperties"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFillStyleList(Child):
    class Meta:
        name = "CT_FillStyleList"

    content: list[
        CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLineStyleList(Child):
    class Meta:
        name = "CT_LineStyleList"

    ln: list[CTLineProperties] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "min_occurs": 3,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableCellProperties(Child):
    class Meta:
        name = "CT_TableCellProperties"

    ln_l: None | CTLineProperties = field(
        default=None,
        metadata={
            "name": "lnL",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_r: None | CTLineProperties = field(
        default=None,
        metadata={
            "name": "lnR",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_t: None | CTLineProperties = field(
        default=None,
        metadata={
            "name": "lnT",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_b: None | CTLineProperties = field(
        default=None,
        metadata={
            "name": "lnB",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_tl_to_br: None | CTLineProperties = field(
        default=None,
        metadata={
            "name": "lnTlToBr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_bl_to_tr: None | CTLineProperties = field(
        default=None,
        metadata={
            "name": "lnBlToTr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    cell3_d: None | CTCell3D = field(
        default=None,
        metadata={
            "name": "cell3D",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    mar_l: None | int = field(
        default=None,
        metadata={
            "name": "marL",
            "type": "Attribute",
            "schema_default": "91440",
        },
    )
    mar_r: None | int = field(
        default=None,
        metadata={
            "name": "marR",
            "type": "Attribute",
            "schema_default": "91440",
        },
    )
    mar_t: None | int = field(
        default=None,
        metadata={
            "name": "marT",
            "type": "Attribute",
            "schema_default": "45720",
        },
    )
    mar_b: None | int = field(
        default=None,
        metadata={
            "name": "marB",
            "type": "Attribute",
            "schema_default": "45720",
        },
    )
    vert: None | STTextVerticalType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTextVerticalType.HORZ,
        },
    )
    anchor: None | STTextAnchoringType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTextAnchoringType.T,
        },
    )
    anchor_ctr: None | bool = field(
        default=None,
        metadata={
            "name": "anchorCtr",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    horz_overflow: None | STTextHorzOverflowType = field(
        default=None,
        metadata={
            "name": "horzOverflow",
            "type": "Attribute",
            "schema_default": STTextHorzOverflowType.CLIP,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextUnderlineFillGroupWrapper(Child):
    class Meta:
        name = "CT_TextUnderlineFillGroupWrapper"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTThemeableLineStyle(Child):
    class Meta:
        name = "CT_ThemeableLineStyle"

    ln_or_ln_ref: None | CTLineProperties | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "ln",
                    "type": ForwardRef("CTLineProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lnRef",
                    "type": ForwardRef("CTStyleMatrixReference"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEffectContainer(Child):
    class Meta:
        name = "CT_EffectContainer"

    content: list[
        CTEffectContainer
        | CTEffectReference
        | CTAlphaBiLevelEffect
        | CTAlphaCeilingEffect
        | CTAlphaFloorEffect
        | CTAlphaInverseEffect
        | CTAlphaModulateEffect
        | CTAlphaModulateFixedEffect
        | CTAlphaOutsetEffect
        | CTAlphaReplaceEffect
        | CTBiLevelEffect
        | CTBlendEffect
        | CTBlurEffect
        | CTColorChangeEffect
        | CTColorReplaceEffect
        | CTDuotoneEffect
        | CTFillEffect
        | CTFillOverlayEffect
        | CTGlowEffect
        | CTGrayscaleEffect
        | CTHSLEffect
        | CTInnerShadowEffect
        | CTLuminanceEffect
        | CTOuterShadowEffect
        | CTPresetShadowEffect
        | CTReflectionEffect
        | CTRelativeOffsetEffect
        | CTSoftEdgesEffect
        | CTTintEffect
        | CTTransformEffect
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cont",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effect",
                    "type": ForwardRef("CTEffectReference"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaBiLevel",
                    "type": ForwardRef("CTAlphaBiLevelEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaCeiling",
                    "type": ForwardRef("CTAlphaCeilingEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaFloor",
                    "type": ForwardRef("CTAlphaFloorEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaInv",
                    "type": ForwardRef("CTAlphaInverseEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("CTAlphaModulateEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaModFix",
                    "type": ForwardRef("CTAlphaModulateFixedEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOutset",
                    "type": ForwardRef("CTAlphaOutsetEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaRepl",
                    "type": ForwardRef("CTAlphaReplaceEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "biLevel",
                    "type": ForwardRef("CTBiLevelEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blend",
                    "type": ForwardRef("CTBlendEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blur",
                    "type": ForwardRef("CTBlurEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "clrChange",
                    "type": ForwardRef("CTColorChangeEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "clrRepl",
                    "type": ForwardRef("CTColorReplaceEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "duotone",
                    "type": ForwardRef("CTDuotoneEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("CTFillEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fillOverlay",
                    "type": ForwardRef("CTFillOverlayEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "glow",
                    "type": ForwardRef("CTGlowEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grayscl",
                    "type": ForwardRef("CTGrayscaleEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "hsl",
                    "type": ForwardRef("CTHSLEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "innerShdw",
                    "type": ForwardRef("CTInnerShadowEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("CTLuminanceEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "outerShdw",
                    "type": ForwardRef("CTOuterShadowEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "prstShdw",
                    "type": ForwardRef("CTPresetShadowEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "reflection",
                    "type": ForwardRef("CTReflectionEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "relOff",
                    "type": ForwardRef("CTRelativeOffsetEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "softEdge",
                    "type": ForwardRef("CTSoftEdgesEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "tint",
                    "type": ForwardRef("CTTintEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "xfrm",
                    "type": ForwardRef("CTTransformEffect"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    type_value: None | STEffectContainerType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": STEffectContainerType.SIB,
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEffectList(Child):
    class Meta:
        name = "CT_EffectList"

    blur: None | CTBlurEffect = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    fill_overlay: None | CTFillOverlayEffect = field(
        default=None,
        metadata={
            "name": "fillOverlay",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    glow: None | CTGlowEffect = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    inner_shdw: None | CTInnerShadowEffect = field(
        default=None,
        metadata={
            "name": "innerShdw",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    outer_shdw: None | CTOuterShadowEffect = field(
        default=None,
        metadata={
            "name": "outerShdw",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    prst_shdw: None | CTPresetShadowEffect = field(
        default=None,
        metadata={
            "name": "prstShdw",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    reflection: None | CTReflectionEffect = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    soft_edge: None | CTSoftEdgesEffect = field(
        default=None,
        metadata={
            "name": "softEdge",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableCellBorderStyle(Child):
    class Meta:
        name = "CT_TableCellBorderStyle"

    left: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    right: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    top: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bottom: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    inside_h: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "name": "insideH",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    inside_v: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "name": "insideV",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tl2br: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tr2bl: None | CTThemeableLineStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAlphaModulateEffect(Child):
    class Meta:
        name = "CT_AlphaModulateEffect"

    cont: None | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBackgroundFormatting(Child):
    class Meta:
        name = "CT_BackgroundFormatting"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBlendEffect(Child):
    class Meta:
        name = "CT_BlendEffect"

    cont: None | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    blend: None | STBlendMode = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEffectProperties(Child):
    class Meta:
        name = "CT_EffectProperties"

    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEffectStyleItem(Child):
    class Meta:
        name = "CT_EffectStyleItem"

    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sp3d: None | CTShape3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupShapeProperties(Child):
    class Meta:
        name = "CT_GroupShapeProperties"

    xfrm: None | CTGroupTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bw_mode: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwMode",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeProperties(Child):
    class Meta:
        name = "CT_ShapeProperties"

    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    cust_geom_or_prst_geom: None | CTCustomGeometry2D | CTPresetGeometry2D = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "custGeom",
                    "type": ForwardRef("CTCustomGeometry2D"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "prstGeom",
                    "type": ForwardRef("CTPresetGeometry2D"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    ln: None | CTLineProperties = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    scene3d: None | CTScene3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sp3d: None | CTShape3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bw_mode: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwMode",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableStyleCellStyle(Child):
    class Meta:
        name = "CT_TableStyleCellStyle"

    tc_bdr: None | CTTableCellBorderStyle = field(
        default=None,
        metadata={
            "name": "tcBdr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    fill_or_fill_ref: None | CTFillProperties | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "fill",
                    "type": ForwardRef("CTFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fillRef",
                    "type": ForwardRef("CTStyleMatrixReference"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    cell3_d: None | CTCell3D = field(
        default=None,
        metadata={
            "name": "cell3D",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextCharacterProperties(Child):
    class Meta:
        name = "CT_TextCharacterProperties"

    ln: None | CTLineProperties = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    highlight: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    u_ln_tx_or_u_ln: None | CTTextUnderlineLineFollowText | CTLineProperties = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "uLnTx",
                    "type": ForwardRef("CTTextUnderlineLineFollowText"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "uLn",
                    "type": ForwardRef("CTLineProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    u_fill_tx_or_u_fill: None | CTTextUnderlineFillFollowText | CTTextUnderlineFillGroupWrapper = (
        field(
            default=None,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "uFillTx",
                        "type": ForwardRef("CTTextUnderlineFillFollowText"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    },
                    {
                        "name": "uFill",
                        "type": ForwardRef("CTTextUnderlineFillGroupWrapper"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    },
                ),
            },
        )
    )
    latin: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ea: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    cs: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sym: None | TextFont = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    hlink_click: None | CTHyperlink = field(
        default=None,
        metadata={
            "name": "hlinkClick",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    hlink_mouse_over: None | CTHyperlink = field(
        default=None,
        metadata={
            "name": "hlinkMouseOver",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    kumimoji: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lang: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt_lang: None | str = field(
        default=None,
        metadata={
            "name": "altLang",
            "type": "Attribute",
        },
    )
    sz: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 100,
            "max_inclusive": 400000,
        },
    )
    b: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    i: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    u: None | STTextUnderlineType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strike: None | STTextStrikeType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    kern: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 400000,
        },
    )
    cap: None | STTextCapsType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spc: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -400000,
            "max_inclusive": 400000,
        },
    )
    normalize_h: None | bool = field(
        default=None,
        metadata={
            "name": "normalizeH",
            "type": "Attribute",
        },
    )
    baseline: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    no_proof: None | bool = field(
        default=None,
        metadata={
            "name": "noProof",
            "type": "Attribute",
        },
    )
    dirty: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    err: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    smt_clean: None | bool = field(
        default=None,
        metadata={
            "name": "smtClean",
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    smt_id: None | int = field(
        default=None,
        metadata={
            "name": "smtId",
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    bmk: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWholeE2OFormatting(Child):
    class Meta:
        name = "CT_WholeE2oFormatting"

    ln: None | CTLineProperties = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEffectStyleList(Child):
    class Meta:
        name = "CT_EffectStyleList"

    effect_style: list[CTEffectStyleItem] = field(
        default_factory=ChildList,
        metadata={
            "name": "effectStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "min_occurs": 3,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlConnector(Child):
    class Meta:
        name = "CT_GvmlConnector"

    nv_cxn_sp_pr: None | CTGvmlConnectorNonVisual = field(
        default=None,
        metadata={
            "name": "nvCxnSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlPicture(Child):
    class Meta:
        name = "CT_GvmlPicture"

    nv_pic_pr: None | CTGvmlPictureNonVisual = field(
        default=None,
        metadata={
            "name": "nvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    blip_fill: None | CTBlipFillProperties = field(
        default=None,
        metadata={
            "name": "blipFill",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRegularTextRun(Child):
    class Meta:
        name = "CT_RegularTextRun"

    r_pr: None | CTTextCharacterProperties = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    t: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableBackgroundStyle(Child):
    class Meta:
        name = "CT_TableBackgroundStyle"

    fill_or_fill_ref: None | CTFillProperties | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "fill",
                    "type": ForwardRef("CTFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fillRef",
                    "type": ForwardRef("CTStyleMatrixReference"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    effect_or_effect_ref: None | CTEffectProperties | CTStyleMatrixReference = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effect",
                    "type": ForwardRef("CTEffectProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectRef",
                    "type": ForwardRef("CTStyleMatrixReference"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTablePartStyle(Child):
    class Meta:
        name = "CT_TablePartStyle"

    tc_tx_style: None | CTTableStyleTextStyle = field(
        default=None,
        metadata={
            "name": "tcTxStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tc_style: None | CTTableStyleCellStyle = field(
        default=None,
        metadata={
            "name": "tcStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextLineBreak(Child):
    class Meta:
        name = "CT_TextLineBreak"

    r_pr: None | CTTextCharacterProperties = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextParagraphProperties(Child):
    class Meta:
        name = "CT_TextParagraphProperties"

    ln_spc: None | CTTextSpacing = field(
        default=None,
        metadata={
            "name": "lnSpc",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    spc_bef: None | CTTextSpacing = field(
        default=None,
        metadata={
            "name": "spcBef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    spc_aft: None | CTTextSpacing = field(
        default=None,
        metadata={
            "name": "spcAft",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bu_clr_tx_or_bu_clr: None | CTTextBulletColorFollowText | CTColor = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "buClrTx",
                    "type": ForwardRef("CTTextBulletColorFollowText"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buClr",
                    "type": ForwardRef("CTColor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    bu_sz_tx_or_bu_sz_pct_or_bu_sz_pts: (
        None | CTTextBulletSizeFollowText | CTTextBulletSizePercent | CTTextBulletSizePoint
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "buSzTx",
                    "type": ForwardRef("CTTextBulletSizeFollowText"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buSzPct",
                    "type": ForwardRef("CTTextBulletSizePercent"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buSzPts",
                    "type": ForwardRef("CTTextBulletSizePoint"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    bu_font_tx_or_bu_font: None | CTTextBulletTypefaceFollowText | TextFont = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "buFontTx",
                    "type": ForwardRef("CTTextBulletTypefaceFollowText"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buFont",
                    "type": ForwardRef("TextFont"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    content: (
        None | CTTextNoBullet | CTTextAutonumberBullet | CTTextCharBullet | CTTextBlipBullet
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "buNone",
                    "type": ForwardRef("CTTextNoBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buAutoNum",
                    "type": ForwardRef("CTTextAutonumberBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buChar",
                    "type": ForwardRef("CTTextCharBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buBlip",
                    "type": ForwardRef("CTTextBlipBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    tab_lst: None | CTTextTabStopList = field(
        default=None,
        metadata={
            "name": "tabLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    def_rpr: None | CTTextCharacterProperties = field(
        default=None,
        metadata={
            "name": "defRPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    mar_l: None | int = field(
        default=None,
        metadata={
            "name": "marL",
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 51206400,
        },
    )
    mar_r: None | int = field(
        default=None,
        metadata={
            "name": "marR",
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 51206400,
        },
    )
    lvl: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 8,
        },
    )
    indent: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -51206400,
            "max_inclusive": 51206400,
        },
    )
    algn: None | STTextAlignType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    def_tab_sz: None | int = field(
        default=None,
        metadata={
            "name": "defTabSz",
            "type": "Attribute",
        },
    )
    rtl: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    ea_ln_brk: None | bool = field(
        default=None,
        metadata={
            "name": "eaLnBrk",
            "type": "Attribute",
        },
    )
    font_algn: None | STTextFontAlignType = field(
        default=None,
        metadata={
            "name": "fontAlgn",
            "type": "Attribute",
        },
    )
    latin_ln_brk: None | bool = field(
        default=None,
        metadata={
            "name": "latinLnBrk",
            "type": "Attribute",
        },
    )
    hanging_punct: None | bool = field(
        default=None,
        metadata={
            "name": "hangingPunct",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleMatrix(Child):
    class Meta:
        name = "CT_StyleMatrix"

    fill_style_lst: None | CTFillStyleList = field(
        default=None,
        metadata={
            "name": "fillStyleLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_style_lst: None | CTLineStyleList = field(
        default=None,
        metadata={
            "name": "lnStyleLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    effect_style_lst: None | CTEffectStyleList = field(
        default=None,
        metadata={
            "name": "effectStyleLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    bg_fill_style_lst: None | CTBackgroundFillStyleList = field(
        default=None,
        metadata={
            "name": "bgFillStyleLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
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
class CTTableStyle(Child):
    class Meta:
        name = "CT_TableStyle"

    tbl_bg: None | CTTableBackgroundStyle = field(
        default=None,
        metadata={
            "name": "tblBg",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    whole_tbl: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "wholeTbl",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    band1_h: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "band1H",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    band2_h: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "band2H",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    band1_v: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "band1V",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    band2_v: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "band2V",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    last_col: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "lastCol",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    first_col: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "firstCol",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    last_row: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "lastRow",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    se_cell: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "seCell",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sw_cell: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "swCell",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    first_row: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "firstRow",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ne_cell: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "neCell",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    nw_cell: None | CTTablePartStyle = field(
        default=None,
        metadata={
            "name": "nwCell",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    style_id: None | str = field(
        default=None,
        metadata={
            "name": "styleId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    style_name: None | str = field(
        default=None,
        metadata={
            "name": "styleName",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextField(Child):
    class Meta:
        name = "CT_TextField"

    r_pr: None | CTTextCharacterProperties = field(
        default=None,
        metadata={
            "name": "rPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    t: None | str = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextListStyle(Child):
    class Meta:
        name = "CT_TextListStyle"

    def_ppr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "defPPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl1p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl1pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl2p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl2pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl3p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl3pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl4p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl4pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl5p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl5pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl6p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl6pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl7p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl7pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl8p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl8pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lvl9p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "lvl9pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class BaseStyles(Child):
    class Meta:
        name = "CT_BaseStyles"

    clr_scheme: None | CTColorScheme = field(
        default=None,
        metadata={
            "name": "clrScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    font_scheme: None | CtBaseStylesFontScheme = field(
        default=None,
        metadata={
            "name": "fontScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    fmt_scheme: None | CTStyleMatrix = field(
        default=None,
        metadata={
            "name": "fmtScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBaseStylesOverride(Child):
    class Meta:
        name = "CT_BaseStylesOverride"

    clr_scheme: None | CTColorScheme = field(
        default=None,
        metadata={
            "name": "clrScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    font_scheme: None | CtBaseStylesOverrideFontScheme = field(
        default=None,
        metadata={
            "name": "fontScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    fmt_scheme: None | CTStyleMatrix = field(
        default=None,
        metadata={
            "name": "fmtScheme",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDefaultShapeDefinition(Child):
    class Meta:
        name = "CT_DefaultShapeDefinition"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    body_pr: None | CTTextBodyProperties = field(
        default=None,
        metadata={
            "name": "bodyPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lst_style: None | CTTextListStyle = field(
        default=None,
        metadata={
            "name": "lstStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableProperties(Child):
    class Meta:
        name = "CT_TableProperties"

    content: (
        None
        | CTNoFillProperties
        | CTSolidColorFillProperties
        | CTGradientFillProperties
        | CTBlipFillProperties
        | CTPatternFillProperties
        | CTGroupFillProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "noFill",
                    "type": ForwardRef("CTNoFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "solidFill",
                    "type": ForwardRef("CTSolidColorFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "gradFill",
                    "type": ForwardRef("CTGradientFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blipFill",
                    "type": ForwardRef("CTBlipFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pattFill",
                    "type": ForwardRef("CTPatternFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpFill",
                    "type": ForwardRef("CTGroupFillProperties"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    effect_lst_or_effect_dag: None | CTEffectList | CTEffectContainer = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "effectLst",
                    "type": ForwardRef("CTEffectList"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "effectDag",
                    "type": ForwardRef("CTEffectContainer"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    table_style_or_table_style_id: None | CTTableStyle | str = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tableStyle",
                    "type": ForwardRef("CTTableStyle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "tableStyleId",
                    "type": str,
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    rtl: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    first_row: None | bool = field(
        default=None,
        metadata={
            "name": "firstRow",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    first_col: None | bool = field(
        default=None,
        metadata={
            "name": "firstCol",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    last_row: None | bool = field(
        default=None,
        metadata={
            "name": "lastRow",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    last_col: None | bool = field(
        default=None,
        metadata={
            "name": "lastCol",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    band_row: None | bool = field(
        default=None,
        metadata={
            "name": "bandRow",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    band_col: None | bool = field(
        default=None,
        metadata={
            "name": "bandCol",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableStyleList(Child):
    class Meta:
        name = "CT_TableStyleList"

    tbl_style: list[CTTableStyle] = field(
        default_factory=ChildList,
        metadata={
            "name": "tblStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    def_value: None | str = field(
        default=None,
        metadata={
            "name": "def",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextParagraph(Child):
    class Meta:
        name = "CT_TextParagraph"

    p_pr: None | CTTextParagraphProperties = field(
        default=None,
        metadata={
            "name": "pPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    r_or_br_or_fld: list[CTRegularTextRun | CTTextLineBreak | CTTextField] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "r",
                    "type": ForwardRef("CTRegularTextRun"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "br",
                    "type": ForwardRef("CTTextLineBreak"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "fld",
                    "type": ForwardRef("CTTextField"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    end_para_rpr: None | CTTextCharacterProperties = field(
        default=None,
        metadata={
            "name": "endParaRPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTClipboardStyleSheet(Child):
    class Meta:
        name = "CT_ClipboardStyleSheet"

    theme_elements: None | BaseStyles = field(
        default=None,
        metadata={
            "name": "themeElements",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    clr_map: None | CTColorMapping = field(
        default=None,
        metadata={
            "name": "clrMap",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTObjectStyleDefaults(Child):
    class Meta:
        name = "CT_ObjectStyleDefaults"

    sp_def: None | CTDefaultShapeDefinition = field(
        default=None,
        metadata={
            "name": "spDef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ln_def: None | CTDefaultShapeDefinition = field(
        default=None,
        metadata={
            "name": "lnDef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tx_def: None | CTDefaultShapeDefinition = field(
        default=None,
        metadata={
            "name": "txDef",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextBody(Child):
    class Meta:
        name = "CT_TextBody"

    body_pr: None | CTTextBodyProperties = field(
        default=None,
        metadata={
            "name": "bodyPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    lst_style: None | CTTextListStyle = field(
        default=None,
        metadata={
            "name": "lstStyle",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    p: list[CTTextParagraph] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class TblStyleLst(CTTableStyleList):
    class Meta:
        name = "tblStyleLst"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(slots=True, kw_only=True)
class ThemeOverride(CTBaseStylesOverride):
    class Meta:
        name = "themeOverride"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass(slots=True, kw_only=True)
class CTGvmlTextShape(Child):
    class Meta:
        name = "CT_GvmlTextShape"

    tx_body: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txBody",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    use_sp_rect_or_xfrm: None | CTGvmlUseShapeRectangle | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "useSpRect",
                    "type": ForwardRef("CTGvmlUseShapeRectangle"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "xfrm",
                    "type": ForwardRef("CTTransform2D"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableCell(Child):
    class Meta:
        name = "CT_TableCell"

    tx_body: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txBody",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tc_pr: None | CTTableCellProperties = field(
        default=None,
        metadata={
            "name": "tcPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    row_span: None | int = field(
        default=None,
        metadata={
            "name": "rowSpan",
            "type": "Attribute",
            "schema_default": "1",
        },
    )
    grid_span: None | int = field(
        default=None,
        metadata={
            "name": "gridSpan",
            "type": "Attribute",
            "schema_default": "1",
        },
    )
    h_merge: None | bool = field(
        default=None,
        metadata={
            "name": "hMerge",
            "type": "Attribute",
            "schema_default": "false",
        },
    )
    v_merge: None | bool = field(
        default=None,
        metadata={
            "name": "vMerge",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class Theme(Child):
    class Meta:
        name = "theme"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"

    theme_elements: None | BaseStyles = field(
        default=None,
        metadata={
            "name": "themeElements",
            "type": "Element",
        },
    )
    object_defaults: None | CTObjectStyleDefaults = field(
        default=None,
        metadata={
            "name": "objectDefaults",
            "type": "Element",
        },
    )
    extra_clr_scheme_lst: None | CTColorSchemeList = field(
        default=None,
        metadata={
            "name": "extraClrSchemeLst",
            "type": "Element",
        },
    )
    cust_clr_lst: None | CTCustomColorList = field(
        default=None,
        metadata={
            "name": "custClrLst",
            "type": "Element",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    other_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##other",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlShape(Child):
    class Meta:
        name = "CT_GvmlShape"

    nv_sp_pr: None | CTGvmlShapeNonVisual = field(
        default=None,
        metadata={
            "name": "nvSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tx_sp: None | CTGvmlTextShape = field(
        default=None,
        metadata={
            "name": "txSp",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTableRow(Child):
    class Meta:
        name = "CT_TableRow"

    tc: list[CTTableCell] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    h: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlGroupShape(Child):
    class Meta:
        name = "CT_GvmlGroupShape"

    nv_grp_sp_pr: None | CTGvmlGroupShapeNonVisual = field(
        default=None,
        metadata={
            "name": "nvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    grp_sp_pr: None | CTGroupShapeProperties = field(
        default=None,
        metadata={
            "name": "grpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    content: list[
        CTGvmlTextShape
        | CTGvmlShape
        | CTGvmlConnector
        | CTGvmlPicture
        | CTGvmlGraphicalObjectFrame
        | CTGvmlGroupShape
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "txSp",
                    "type": ForwardRef("CTGvmlTextShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sp",
                    "type": ForwardRef("CTGvmlShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "cxnSp",
                    "type": ForwardRef("CTGvmlConnector"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "pic",
                    "type": ForwardRef("CTGvmlPicture"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "graphicFrame",
                    "type": ForwardRef("CTGvmlGraphicalObjectFrame"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "grpSp",
                    "type": ForwardRef("CTGvmlGroupShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTable(Child):
    class Meta:
        name = "CT_Table"

    tbl_pr: None | CTTableProperties = field(
        default=None,
        metadata={
            "name": "tblPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tbl_grid: None | CTTableGrid = field(
        default=None,
        metadata={
            "name": "tblGrid",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    tr: list[CTTableRow] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class Tbl(CTTable):
    class Meta:
        name = "tbl"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.main import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.dml.main.el",
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
