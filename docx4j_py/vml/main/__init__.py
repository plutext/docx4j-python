from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList
from docx4j_py.vml.officedrawing import (
    STConnectorType,
    STHrAlign,
    STInsetMode,
)

__NAMESPACE__ = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class CTF(Child):
    class Meta:
        name = "CT_F"

    eqn: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class STEditAs(Enum):
    CANVAS = "canvas"
    ORGCHART = "orgchart"
    RADIAL = "radial"
    CYCLE = "cycle"
    STACKED = "stacked"
    VENN = "venn"
    BULLSEYE = "bullseye"


class STExt(Enum):
    VIEW = "view"
    EDIT = "edit"
    BACKWARD_COMPATIBLE = "backwardCompatible"


class STFillMethod(Enum):
    NONE = "none"
    LINEAR = "linear"
    SIGMA = "sigma"
    ANY = "any"
    LINEAR_SIGMA = "linear sigma"


class STFillType(Enum):
    SOLID = "solid"
    GRADIENT = "gradient"
    GRADIENT_RADIAL = "gradientRadial"
    TILE = "tile"
    PATTERN = "pattern"
    FRAME = "frame"


class STImageAspect(Enum):
    IGNORE = "ignore"
    AT_MOST = "atMost"
    AT_LEAST = "atLeast"


class STShadowType(Enum):
    SINGLE = "single"
    DOUBLE = "double"
    EMBOSS = "emboss"
    PERSPECTIVE = "perspective"


class STStrokeArrowLength(Enum):
    SHORT = "short"
    MEDIUM = "medium"
    LONG = "long"


class STStrokeArrowType(Enum):
    NONE = "none"
    BLOCK = "block"
    CLASSIC = "classic"
    OVAL = "oval"
    DIAMOND = "diamond"
    OPEN = "open"


class STStrokeArrowWidth(Enum):
    NARROW = "narrow"
    MEDIUM = "medium"
    WIDE = "wide"


class STStrokeEndCap(Enum):
    FLAT = "flat"
    SQUARE = "square"
    ROUND = "round"


class STStrokeJoinStyle(Enum):
    ROUND = "round"
    BEVEL = "bevel"
    MITER = "miter"


class STStrokeLineStyle(Enum):
    SINGLE = "single"
    THIN_THIN = "thinThin"
    THIN_THICK = "thinThick"
    THICK_THIN = "thickThin"
    THICK_BETWEEN_THIN = "thickBetweenThin"


class STTrueFalse(Enum):
    T = "t"
    F = "f"
    TRUE = "true"
    FALSE = "false"


class StTrueFalseBlank(Enum):
    T = "t"
    F = "f"
    TRUE = "true"
    FALSE = "false"
    VALUE = ""


@dataclass(slots=True, kw_only=True)
class CTFill(Child):
    class Meta:
        name = "CT_Fill"

    fill: None | OfficedrawingFill = field(
        default=None,
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
    type_value: None | STFillType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    on: None | STTrueFalse = field(
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
    opacity: None | str = field(
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
    src: None | str = field(
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
    size: None | str = field(
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
    position: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    aspect: None | STImageAspect = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    colors: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    angle: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alignshape: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    focus: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    focussize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    focusposition: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    method: None | STFillMethod = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    detectmouseclick: None | OfficedrawingSttrueFalse = field(
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
    opacity2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    recolor: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    rotate: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    schemas_openxmlformats_org_office_document_2006_relationships_id: None | str = field(
        default=None,
        metadata={
            "name": "id",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    relid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTFormulas(Child):
    class Meta:
        name = "CT_Formulas"

    f: list[CTF] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTH(Child):
    class Meta:
        name = "CT_H"

    position: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    polar: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    map: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    invx: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    invy: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    switch: None | StTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    xrange: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    yrange: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    radiusrange: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTImageData(Child):
    class Meta:
        name = "CT_ImageData"

    id: None | str = field(
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
    cropleft: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    croptop: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cropright: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cropbottom: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gain: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    blacklevel: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gamma: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    grayscale: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    bilevel: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    embosscolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    recolortarget: None | str = field(
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
    oleid: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    detectmouseclick: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    movie: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    relid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    schemas_openxmlformats_org_office_document_2006_relationships_id: None | str = field(
        default=None,
        metadata={
            "name": "id",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    pict: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    schemas_openxmlformats_org_office_document_2006_relationships_href: None | str = field(
        default=None,
        metadata={
            "name": "href",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPath(Child):
    class Meta:
        name = "CT_Path"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    v: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    limo: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    textboxrect: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fillok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    shadowok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    arrowok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gradientshapeok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    textpathok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    insetpenok: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    connecttype: None | STConnectType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectlocs: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectangles: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    extrusionok: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShadow(Child):
    class Meta:
        name = "CT_Shadow"

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
    type_value: None | STShadowType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    obscured: None | STTrueFalse = field(
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
    opacity: None | str = field(
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
    color2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    offset2: None | str = field(
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
class CTStroke(Child):
    class Meta:
        name = "CT_Stroke"

    left: None | Left = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    top: None | Top = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    right: None | Right = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bottom: None | Bottom = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    column: None | Column = field(
        default=None,
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
    filltype: None | STFillType = field(
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
    color2: None | str = field(
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
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    schemas_openxmlformats_org_office_document_2006_relationships_id: None | str = field(
        default=None,
        metadata={
            "name": "id",
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    insetpen: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    relid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextPath(Child):
    class Meta:
        name = "CT_TextPath"

    id: None | str = field(
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
    on: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fitshape: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    fitpath: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    trim: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    xscale: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    string: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextbox(Child):
    class Meta:
        name = "CT_Textbox"

    txbx_content_or_local_element: None | TxbxContent | object = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "txbxContent",
                    "type": ForwardRef("TxbxContent"),
                    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                },
                {
                    "wildcard": True,
                    "type": object,
                    "namespace": "##local",
                },
            ),
        },
    )
    id: None | str = field(
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
    inset: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    singleclick: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTHandles(Child):
    class Meta:
        name = "CT_Handles"

    h: list[CTH] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )


@dataclass(slots=True, kw_only=True)
class Fill(CTFill):
    class Meta:
        name = "fill"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Formulas(CTFormulas):
    class Meta:
        name = "formulas"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Imagedata(CTImageData):
    class Meta:
        name = "imagedata"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Path(CTPath):
    class Meta:
        name = "path"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Shadow(CTShadow):
    class Meta:
        name = "shadow"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Stroke(CTStroke):
    class Meta:
        name = "stroke"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Textbox(CTTextbox):
    class Meta:
        name = "textbox"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Textpath(CTTextPath):
    class Meta:
        name = "textpath"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class CTBackground(Child):
    class Meta:
        name = "CT_Background"

    fill: None | Fill = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "urn:schemas-microsoft-com:vml",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    targetscreensize: None | StScreenSize = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class Handles(CTHandles):
    class Meta:
        name = "handles"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class CTArc(Child):
    class Meta:
        name = "CT_Arc"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    start_angle: None | Decimal = field(
        default=None,
        metadata={
            "name": "startAngle",
            "type": "Attribute",
        },
    )
    end_angle: None | Decimal = field(
        default=None,
        metadata={
            "name": "endAngle",
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCurve(Child):
    class Meta:
        name = "CT_Curve"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    from_value: None | str = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Attribute",
        },
    )
    control1: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    control2: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    to: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTImage(Child):
    class Meta:
        name = "CT_Image"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    src: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cropleft: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    croptop: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cropright: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    cropbottom: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gain: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    blacklevel: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gamma: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    grayscale: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    bilevel: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTLine(Child):
    class Meta:
        name = "CT_Line"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    from_value: None | str = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Attribute",
        },
    )
    to: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOval(Child):
    class Meta:
        name = "CT_Oval"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPolyLine(Child):
    class Meta:
        name = "CT_PolyLine"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
        | Ink
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
                {
                    "name": "ink",
                    "type": ForwardRef("Ink"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    points: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRect(Child):
    class Meta:
        name = "CT_Rect"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRoundRect(Child):
    class Meta:
        name = "CT_RoundRect"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    arcsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShape(Child):
    class Meta:
        name = "CT_Shape"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
        | Ink
        | Iscomment
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
                {
                    "name": "ink",
                    "type": ForwardRef("Ink"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "iscomment",
                    "type": ForwardRef("Iscomment"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    adj: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    path_attribute: None | str = field(
        default=None,
        metadata={
            "name": "path",
            "type": "Attribute",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )
    equationxml: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapetype(Child):
    class Meta:
        name = "CT_Shapetype"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
            ),
        },
    )
    complex: None | Complex = field(
        default=None,
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
    style: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    chromakey: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    filled: None | STTrueFalse = field(
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
    opacity: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    stroked: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokecolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    strokeweight: None | str = field(
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
    spt: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    connectortype: None | STConnectorType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STConnectorType.STRAIGHT,
        },
    )
    bwmode: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwpure: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bwnormal: None | STBWMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    forcedash: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oleicon: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    ole: None | OfficedrawingStTrueFalseBlank = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    preferrelative: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    cliptowrap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    clip: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    adj: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    path_attribute: None | str = field(
        default=None,
        metadata={
            "name": "path",
            "type": "Attribute",
        },
    )
    master: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )


@dataclass(slots=True, kw_only=True)
class Background(CTBackground):
    class Meta:
        name = "background"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Arc(CTArc):
    class Meta:
        name = "arc"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Curve(CTCurve):
    class Meta:
        name = "curve"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Image(CTImage):
    class Meta:
        name = "image"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Line(CTLine):
    class Meta:
        name = "line"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Oval(CTOval):
    class Meta:
        name = "oval"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Polyline(CTPolyLine):
    class Meta:
        name = "polyline"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Rect(CTRect):
    class Meta:
        name = "rect"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Roundrect(CTRoundRect):
    class Meta:
        name = "roundrect"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Shape(CTShape):
    class Meta:
        name = "shape"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class Shapetype(CTShapetype):
    class Meta:
        name = "shapetype"
        namespace = "urn:schemas-microsoft-com:vml"


@dataclass(slots=True, kw_only=True)
class CTGroup(Child):
    class Meta:
        name = "CT_Group"

    content: list[
        Path
        | Formulas
        | Handles
        | Fill
        | Stroke
        | Shadow
        | Textbox
        | Textpath
        | Imagedata
        | Skew
        | Extrusion
        | Callout
        | Lock
        | Clippath
        | Signatureline
        | Wrap
        | Anchorlock
        | Bordertop
        | Borderbottom
        | Borderleft
        | Borderright
        | ClientData
        | Textdata
        | Group
        | Shape
        | Shapetype
        | Arc
        | Curve
        | Image
        | Line
        | Oval
        | Polyline
        | Rect
        | Roundrect
        | Diagram
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "path",
                    "type": ForwardRef("Path"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "formulas",
                    "type": ForwardRef("Formulas"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "handles",
                    "type": ForwardRef("Handles"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "fill",
                    "type": ForwardRef("Fill"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "stroke",
                    "type": ForwardRef("Stroke"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shadow",
                    "type": ForwardRef("Shadow"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textbox",
                    "type": ForwardRef("Textbox"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "textpath",
                    "type": ForwardRef("Textpath"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "imagedata",
                    "type": ForwardRef("Imagedata"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "skew",
                    "type": ForwardRef("Skew"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "extrusion",
                    "type": ForwardRef("Extrusion"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "callout",
                    "type": ForwardRef("Callout"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "lock",
                    "type": ForwardRef("Lock"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "clippath",
                    "type": ForwardRef("Clippath"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "signatureline",
                    "type": ForwardRef("Signatureline"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
                {
                    "name": "wrap",
                    "type": ForwardRef("Wrap"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "anchorlock",
                    "type": ForwardRef("Anchorlock"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "bordertop",
                    "type": ForwardRef("Bordertop"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderbottom",
                    "type": ForwardRef("Borderbottom"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderleft",
                    "type": ForwardRef("Borderleft"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "borderright",
                    "type": ForwardRef("Borderright"),
                    "namespace": "urn:schemas-microsoft-com:office:word",
                },
                {
                    "name": "ClientData",
                    "type": ForwardRef("ClientData"),
                    "namespace": "urn:schemas-microsoft-com:office:excel",
                },
                {
                    "name": "textdata",
                    "type": ForwardRef("Textdata"),
                    "namespace": "urn:schemas-microsoft-com:office:powerpoint",
                },
                {
                    "name": "group",
                    "type": ForwardRef("Group"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shape",
                    "type": ForwardRef("Shape"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "shapetype",
                    "type": ForwardRef("Shapetype"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "arc",
                    "type": ForwardRef("Arc"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "curve",
                    "type": ForwardRef("Curve"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "image",
                    "type": ForwardRef("Image"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "line",
                    "type": ForwardRef("Line"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "oval",
                    "type": ForwardRef("Oval"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "polyline",
                    "type": ForwardRef("Polyline"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "rect",
                    "type": ForwardRef("Rect"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "roundrect",
                    "type": ForwardRef("Roundrect"),
                    "namespace": "urn:schemas-microsoft-com:vml",
                },
                {
                    "name": "diagram",
                    "type": ForwardRef("Diagram"),
                    "namespace": "urn:schemas-microsoft-com:office:office",
                },
            ),
        },
    )
    id: None | str = field(
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
    href: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    target: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    class_value: None | str = field(
        default=None,
        metadata={
            "name": "class",
            "type": "Attribute",
        },
    )
    title: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    alt: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordsize: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    coordorigin: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    wrapcoords: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    print: None | STTrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    oned: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    regroupid: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    doubleclicknotify: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    button: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userhidden: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bullet: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hr: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrstd: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrnoshade: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hrpct: None | float = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    hralign: None | STHrAlign = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STHrAlign.LEFT,
        },
    )
    allowincell: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    allowoverlap: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    userdrawn: None | OfficedrawingSttrueFalse = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    bordertopcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderleftcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderbottomcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    borderrightcolor: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayout: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmnodekind: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    dgmlayoutmru: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    insetmode: None | STInsetMode = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "schema_default": STInsetMode.CUSTOM,
        },
    )
    anchor_id: None | str = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
        },
    )
    filled: None | STTrueFalse = field(
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
    editas: None | STEditAs = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    tableproperties: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    tablelimits: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
        },
    )
    gfxdata: None | bytes = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "urn:schemas-microsoft-com:office:office",
            "format": "base64",
        },
    )


@dataclass(slots=True, kw_only=True)
class Group(CTGroup):
    class Meta:
        name = "group"
        namespace = "urn:schemas-microsoft-com:vml"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.vml.officedrawing import (
    Bottom,
    Callout,
    Clippath,
    Column,
    Complex,
    Diagram,
    Extrusion,
    Ink,
    Left,
    Lock,
    Right,
    Signatureline,
    Skew,
    STBWMode,
    STConnectType,
    StScreenSize,
    Top,
)
from docx4j_py.vml.officedrawing import (
    Fill as OfficedrawingFill,
)
from docx4j_py.vml.officedrawing import (
    STTrueFalse as OfficedrawingSttrueFalse,
)
from docx4j_py.vml.officedrawing import (
    StTrueFalseBlank as OfficedrawingStTrueFalseBlank,
)
from docx4j_py.vml.presentation_drawing import (
    Iscomment,
    Textdata,
)
from docx4j_py.vml.spreadsheet_drawing import ClientData
from docx4j_py.vml.wordprocessing_drawing import (
    Anchorlock,
    Borderbottom,
    Borderleft,
    Borderright,
    Bordertop,
    Wrap,
)
from docx4j_py.wml import TxbxContent


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.vml.main import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.vml.main.el",
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
