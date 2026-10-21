from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class CTClipPath(Child):
    class Meta:
        name = "CT_ClipPath"

    v: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTEntry(Child):
    class Meta:
        name = "CT_Entry"

    new: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    old: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class STBWMode(Enum):
    COLOR = "color"
    AUTO = "auto"
    GRAY_SCALE = "grayScale"
    LIGHT_GRAYSCALE = "lightGrayscale"
    INVERSE_GRAY = "inverseGray"
    GRAY_OUTLINE = "grayOutline"
    HIGH_CONTRAST = "highContrast"
    BLACK = "black"
    WHITE = "white"
    HIDE = "hide"
    UNDRAWN = "undrawn"
    BLACK_TEXT_AND_LINES = "blackTextAndLines"


class STColorMode(Enum):
    AUTO = "auto"
    CUSTOM = "custom"


class STConnectType(Enum):
    NONE = "none"
    RECT = "rect"
    SEGMENTS = "segments"
    CUSTOM = "custom"


class STConnectorType(Enum):
    NONE = "none"
    STRAIGHT = "straight"
    ELBOW = "elbow"
    CURVED = "curved"


class STExtrusionPlane(Enum):
    XY = "XY"
    ZX = "ZX"
    YZ = "YZ"


class STExtrusionRender(Enum):
    SOLID = "solid"
    WIRE_FRAME = "wireFrame"
    BOUNDING_CUBE = "boundingCube"


class STExtrusionType(Enum):
    PERSPECTIVE = "perspective"
    PARALLEL = "parallel"


class STFillType(Enum):
    GRADIENT_CENTER = "gradientCenter"
    SOLID = "solid"
    PATTERN = "pattern"
    TILE = "tile"
    FRAME = "frame"
    GRADIENT_UNSCALED = "gradientUnscaled"
    GRADIENT_RADIAL = "gradientRadial"
    GRADIENT = "gradient"
    BACKGROUND = "background"


class STHow(Enum):
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class STHrAlign(Enum):
    LEFT = "left"
    RIGHT = "right"
    CENTER = "center"


class STInsetMode(Enum):
    AUTO = "auto"
    CUSTOM = "custom"


class STOLEDrawAspect(Enum):
    CONTENT = "Content"
    ICON = "Icon"


class STOLELinkType(Enum):
    PICTURE = "Picture"
    BITMAP = "Bitmap"
    ENHANCED_META_FILE = "EnhancedMetaFile"


class STOLEType(Enum):
    EMBED = "Embed"
    LINK = "Link"


class STOLEUpdateMode(Enum):
    ALWAYS = "Always"
    ON_CALL = "OnCall"


class STRType(Enum):
    ARC = "arc"
    CALLOUT = "callout"
    CONNECTOR = "connector"
    ALIGN = "align"


class STTrueFalse(Enum):
    T = "t"
    F = "f"
    TRUE = "true"
    FALSE = "false"


class StAngle(Enum):
    ANY = "any"
    VALUE_30 = "30"
    VALUE_45 = "45"
    VALUE_60 = "60"
    VALUE_90 = "90"
    AUTO = "auto"


class StScreenSize(Enum):
    VALUE_544_376 = "544,376"
    VALUE_640_480 = "640,480"
    VALUE_720_512 = "720,512"
    VALUE_800_600 = "800,600"
    VALUE_1024_768 = "1024,768"
    VALUE_1152_862 = "1152,862"


class StTrueFalseBlank(Enum):
    VALUE = ""
    T = "t"
    F = "f"
    TRUE = "true"
    FALSE = "false"


@dataclass(slots=True, kw_only=True)
class CTCallout(Child):
    class Meta:
        name = "CT_Callout"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    on: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    gap: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    angle: None | StAngle = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dropauto: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    drop: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    distance: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lengthspecified: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STTrueFalse.F,
        },
    )
    length: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    accentbar: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    textborder: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    minusx: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    minusy: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorMenu(Child):
    class Meta:
        name = "CT_ColorMenu"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fillcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    shadowcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    extrusioncolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorMru(Child):
    class Meta:
        name = "CT_ColorMru"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    colors: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTComplex(Child):
    class Meta:
        name = "CT_Complex"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTExtrusion(Child):
    class Meta:
        name = "CT_Extrusion"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    on: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    type_value: None | STExtrusionType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": STExtrusionType.PARALLEL,
        },
    )
    render: None | STExtrusionRender = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STExtrusionRender.SOLID,
        },
    )
    viewpointorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    viewpoint: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    plane: None | STExtrusionPlane = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STExtrusionPlane.XY,
        },
    )
    skewangle: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    skewamt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    foredepth: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    backdepth: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    orientation: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    orientationangle: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lockrotationcenter: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    autorotationcenter: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    rotationcenter: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    rotationangle: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    colormode: None | STColorMode = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    color: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    shininess: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    specularity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    diffusity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    metal: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    edge: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    facet: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightface: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    brightness: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightposition: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightlevel: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightharsh: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightposition2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightlevel2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lightharsh2: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFill(Child):
    class Meta:
        name = "CT_Fill"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    type_value: None | STFillType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTIdMap(Child):
    class Meta:
        name = "CT_IdMap"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    data: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInk(Child):
    class Meta:
        name = "CT_Ink"

    i: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "format": "base64",
        },
    )
    annotation: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLock(Child):
    class Meta:
        name = "CT_Lock"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    position: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    selection: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    grouping: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    ungrouping: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    rotation: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cropping: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    verticies: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    adjusthandles: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    text: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    aspectratio: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    shapetype: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOLEObject(Child):
    class Meta:
        name = "CT_OLEObject"

    link_type: None | STOLELinkType = field(
        default=None,
        metadata={
            "name": "LinkType",
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    locked_field: None | StTrueFalseBlank = field(
        default=None,
        metadata={
            "name": "LockedField",
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    field_codes: None | str = field(
        default=None,
        metadata={
            "name": "FieldCodes",
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    type_value: None | STOLEType = field(
        default=None,
        metadata={
            "name": "Type",
            "type": "Attribute",
        },
    )
    prog_id: None | str = field(
        default=None,
        metadata={
            "name": "ProgID",
            "type": "Attribute",
        },
    )
    shape_id: None | str = field(
        default=None,
        metadata={
            "name": "ShapeID",
            "type": "Attribute",
        },
    )
    draw_aspect: None | STOLEDrawAspect = field(
        default=None,
        metadata={
            "name": "DrawAspect",
            "type": "Attribute",
        },
    )
    object_id: None | str = field(
        default=None,
        metadata={
            "name": "ObjectID",
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
    update_mode: None | STOLEUpdateMode = field(
        default=None,
        metadata={
            "name": "UpdateMode",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTProxy(Child):
    class Meta:
        name = "CT_Proxy"

    start: None | StTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": StTrueFalseBlank.FALSE,
        },
    )
    end: None | StTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": StTrueFalseBlank.FALSE,
        },
    )
    idref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    connectloc: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRegroupTable(Child):
    class Meta:
        name = "CT_RegroupTable"

    entry: list[CTEntry] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelation(Child):
    class Meta:
        name = "CT_Relation"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    idsrc: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    iddest: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    idcntr: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSignatureLine(Child):
    class Meta:
        name = "CT_SignatureLine"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    issignatureline: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    provid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    signinginstructionsset: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    allowcomments: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    showsigndate: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    suggestedsigner: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    suggestedsigner2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    suggestedsigneremail: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    signinginstructions: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    addlxml: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    sigprovurl: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSkew(Child):
    class Meta:
        name = "CT_Skew"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    on: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    offset: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    origin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    matrix: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStrokeChild(Child):
    class Meta:
        name = "CT_StrokeChild"

    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    on: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    weight: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    color: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    color2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    linestyle: None | STStrokeLineStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    miterlimit: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    joinstyle: None | STStrokeJoinStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    endcap: None | STStrokeEndCap = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dashstyle: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    insetpen: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filltype: None | MainStfillType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    src: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    imageaspect: None | STImageAspect = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    imagesize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    imagealignshape: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    startarrow: None | STStrokeArrowType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    startarrowwidth: None | STStrokeArrowWidth = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    startarrowlength: None | STStrokeArrowLength = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    endarrow: None | STStrokeArrowType = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    endarrowwidth: None | STStrokeArrowWidth = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    endarrowlength: None | STStrokeArrowLength = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    althref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class Clippath(CTClipPath):
    class Meta:
        name = "clippath"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class CTR(Child):
    class Meta:
        name = "CT_R"

    proxy: list[CTProxy] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    type_value: None | STRType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    how: None | STHow = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    idref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelationTable(Child):
    class Meta:
        name = "CT_RelationTable"

    rel: list[CTRelation] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Oleobject(CTOLEObject):
    class Meta:
        name = "OLEObject"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Bottom(CTStrokeChild):
    class Meta:
        name = "bottom"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Callout(CTCallout):
    class Meta:
        name = "callout"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Column(CTStrokeChild):
    class Meta:
        name = "column"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Complex(CTComplex):
    class Meta:
        name = "complex"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Extrusion(CTExtrusion):
    class Meta:
        name = "extrusion"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Fill(CTFill):
    class Meta:
        name = "fill"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Ink(CTInk):
    class Meta:
        name = "ink"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Left(CTStrokeChild):
    class Meta:
        name = "left"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Lock(CTLock):
    class Meta:
        name = "lock"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Right(CTStrokeChild):
    class Meta:
        name = "right"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Signatureline(CTSignatureLine):
    class Meta:
        name = "signatureline"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Skew(CTSkew):
    class Meta:
        name = "skew"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Top(CTStrokeChild):
    class Meta:
        name = "top"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class CTDiagram(Child):
    class Meta:
        name = "CT_Diagram"

    relationtable: None | CTRelationTable = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    dgmstyle: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    autoformat: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    reverse: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    autolayout: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dgmscalex: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dgmscaley: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dgmfontsize: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    constrainbounds: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dgmbasetextscale: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRules(Child):
    class Meta:
        name = "CT_Rules"

    r: list[CTR] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeDefaults(Child):
    class Meta:
        name = "CT_ShapeDefaults"

    fill: None | MainFill = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    stroke: None | Stroke = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    textbox: None | Textbox = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    shadow: None | Shadow = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    skew: None | Skew = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    extrusion: None | Extrusion = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    callout: None | Callout = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    lock: None | Lock = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    colormru: None | CTColorMru = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    colormenu: None | CTColorMenu = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    spidmax: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    style: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fill_attribute: None | STTrueFalse = field(
        default=None,
        metadata={
            "name": "fill",
            "type": "Attribute",
        },
    )
    fillcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroke_attribute: None | STTrueFalse = field(
        default=None,
        metadata={
            "name": "stroke",
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    allowincell: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeLayout(Child):
    class Meta:
        name = "CT_ShapeLayout"

    idmap: None | CTIdMap = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regrouptable: None | CTRegroupTable = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    rules: None | CTRules = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ext: None | STExt = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Diagram(CTDiagram):
    class Meta:
        name = "diagram"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Shapedefaults(CTShapeDefaults):
    class Meta:
        name = "shapedefaults"
        namespace = "urn:schemas-microsoft-com:office:office"


@dataclass(slots=True, kw_only=True)
class Shapelayout(CTShapeLayout):
    class Meta:
        name = "shapelayout"
        namespace = "urn:schemas-microsoft-com:office:office"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.vml.main import (
    Fill as MainFill,
)
from docx4j_py.vml.main import (
    Shadow,
    STExt,
    STImageAspect,
    Stroke,
    STStrokeArrowLength,
    STStrokeArrowType,
    STStrokeArrowWidth,
    STStrokeEndCap,
    STStrokeJoinStyle,
    STStrokeLineStyle,
    Textbox,
)
from docx4j_py.vml.main import (
    STFillType as MainStfillType,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.vml.officedrawing import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.vml.officedrawing.el",
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
