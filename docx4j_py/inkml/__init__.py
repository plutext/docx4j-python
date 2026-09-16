from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import ForwardRef

from docx4j_xsdata.models.datatype import XmlDateTime

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://www.w3.org/2003/InkML"


@dataclass(slots=True, kw_only=True)
class AffineType(Child):
    class Meta:
        name = "affine.type"

    value: str = field(default="")
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )


@dataclass(slots=True, kw_only=True)
class AnnotationType(Child):
    class Meta:
        name = "annotation.type"

    value: str = field(default="")
    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    encoding: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    any_attributes: dict[str, str] = field(
        default_factory=dict,
        metadata={
            "type": "Attributes",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class AnnotationXMLType(Child):
    class Meta:
        name = "annotationXML.type"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
            "process_contents": "skip",
        },
    )
    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    encoding: None | str = field(
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


@dataclass(slots=True, kw_only=True)
class BindType(Child):
    class Meta:
        name = "bind.type"

    source: None | str = field(
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
    column: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    variable: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class BooleanStrType(Enum):
    F = "F"
    T = "T"


class BrushRasterOpType(Enum):
    NO_OPERATION = "noOperation"
    COPY_PEN = "copyPen"
    MASK_PEN = "maskPen"


class BrushTipType(Enum):
    ELLIPSE = "ellipse"
    RECTANGLE = "rectangle"
    DROP = "drop"


@dataclass(slots=True, kw_only=True)
class CTMatrix(Child):
    class Meta:
        name = "CT_Matrix"

    value: str = field(default="")
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )


@dataclass(slots=True, kw_only=True)
class LatencyType(Child):
    class Meta:
        name = "latency.type"

    value: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class StStandardLengthUnits(Enum):
    M = "m"
    CM = "cm"
    MM = "mm"
    IN = "in"
    PT = "pt"
    PC = "pc"
    EM = "em"
    EX = "ex"
    VALUE_1_M = "1/m"
    VALUE_1_CM = "1/cm"
    VALUE_1_MM = "1/mm"
    VALUE_1_IN = "1/in"
    VALUE_1_PT = "1/pt"
    VALUE_1_PC = "1/pc"
    VALUE_1_EM = "1/em"
    VALUE_1_EX = "1/ex"


class StStandardTimeUnits(Enum):
    S = "s"
    MS = "ms"
    VALUE_1_S = "1/s"
    VALUE_1_MS = "1/ms"


@dataclass(slots=True, kw_only=True)
class SampleRateType(Child):
    class Meta:
        name = "sampleRate.type"

    uniform: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    value: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


class StandardBrushPropertyNameType(Enum):
    WIDTH = "width"
    HEIGHT = "height"
    COLOR = "color"
    TRANSPARENCY = "transparency"
    TIP = "tip"
    RASTER_OP = "rasterOp"
    ANTI_ALIASED = "antiAliased"
    FIT_TO_CURVE = "fitToCurve"
    IGNORE_PRESSURE = "ignorePressure"


class StandardChannelNameType(Enum):
    X = "X"
    Y = "Y"
    Z = "Z"
    F = "F"
    S = "S"
    B1 = "B1"
    B2 = "B2"
    B3 = "B3"
    B4 = "B4"
    OTX = "OTx"
    OTY = "OTy"
    OA = "OA"
    OE = "OE"
    OR = "OR"
    C = "C"
    CR = "CR"
    CG = "CG"
    CB = "CB"
    CC = "CC"
    CM = "CM"
    CY = "CY"
    CK = "CK"
    A = "A"
    W = "W"
    BW = "BW"
    BH = "BH"
    T = "T"


class StandardChannelPropertyNameType(Enum):
    THRESHOLD = "threshold"
    RESOLUTION = "resolution"
    QUANTIZATION = "quantization"
    NOISE = "noise"
    ACCURACY = "accuracy"
    CROSS_COUPLING = "crossCoupling"
    SKEW = "skew"
    MIN_BANDWIDTH = "minBandwidth"
    PEAK_RATE = "peakRate"
    DISTORTION = "distortion"


@dataclass(slots=True, kw_only=True)
class TimestampType(Child):
    class Meta:
        name = "timestamp.type"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    time: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    timestamp_ref: None | str = field(
        default=None,
        metadata={
            "name": "timestampRef",
            "type": "Attribute",
        },
    )
    time_string: None | XmlDateTime = field(
        default=None,
        metadata={
            "name": "timeString",
            "type": "Attribute",
        },
    )
    time_offset: None | Decimal = field(
        default=None,
        metadata={
            "name": "timeOffset",
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class TraceViewType(Child):
    class Meta:
        name = "traceView.type"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    trace_data_ref: None | str = field(
        default=None,
        metadata={
            "name": "traceDataRef",
            "type": "Attribute",
        },
    )
    from_value: None | str = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Attribute",
            "pattern": r"-?[0-9]+(\s*:\s*-?[0-9]+)*",
        },
    )
    to: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"-?[0-9]+(\s*:\s*-?[0-9]+)*",
        },
    )


class ChannelTypeOrientation(Enum):
    VE = "+ve"
    VE_1 = "-ve"


class ChannelTypeType(Enum):
    INTEGER = "integer"
    DECIMAL = "decimal"
    DOUBLE = "double"
    BOOLEAN = "boolean"


class MappingTypeType(Enum):
    IDENTITY = "identity"
    PRODUCT = "product"
    TABLE = "table"
    AFFINE = "affine"
    MATHML = "mathml"
    UNKNOWN = "unknown"


class StandardUnitsType(Enum):
    M = "m"
    CM = "cm"
    MM = "mm"
    IN = "in"
    PT = "pt"
    PC = "pc"
    EM = "em"
    EX = "ex"
    S = "s"
    MS = "ms"
    KG = "kg"
    G = "g"
    MG = "mg"
    N = "N"
    DEG = "deg"
    RAD = "rad"
    PERCENT_SIGN = "%"
    DEV = "dev"
    VALUE_1_M = "1/m"
    VALUE_1_CM = "1/cm"
    VALUE_1_MM = "1/mm"
    VALUE_1_IN = "1/in"
    VALUE_1_PT = "1/pt"
    VALUE_1_PC = "1/pc"
    VALUE_1_EM = "1/em"
    VALUE_1_EX = "1/ex"
    VALUE_1_S = "1/s"
    VALUE_1_MS = "1/ms"
    VALUE_1_KG = "1/kg"
    VALUE_1_G = "1/g"
    VALUE_1_MG = "1/mg"
    VALUE_1_LB = "1/lb"
    VALUE_1_N = "1/N"
    VALUE_1_DEG = "1/deg"
    VALUE_1_RAD = "1/rad"
    VALUE_1_DEV = "1/dev"
    M_S_1 = "m/s"
    CM_S = "cm/s"
    MM_S = "mm/s"
    MM_DEG = "mm/deg"
    MM_RAD = "mm/rad"


class TableTypeApply(Enum):
    ABSOLUTE = "absolute"
    RELATIVE = "relative"


class TableTypeInterpolation(Enum):
    FLOOR = "floor"
    MIDDLE = "middle"
    CEILING = "ceiling"
    LINEAR = "linear"
    CUBIC = "cubic"


class TraceTypeContinuation(Enum):
    BEGIN = "begin"
    END = "end"
    MIDDLE = "middle"


class TraceTypeType(Enum):
    PEN_DOWN = "penDown"
    PEN_UP = "penUp"
    INDETERMINATE = "indeterminate"


@dataclass(slots=True, kw_only=True)
class ActiveAreaType(Child):
    class Meta:
        name = "activeArea.type"

    size: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    height: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    width: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    units: None | StandardUnitsType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class BrushPropertyType(Child):
    class Meta:
        name = "brushProperty.type"

    name: None | StandardBrushPropertyNameType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    value: None | BrushRasterOpType | BrushTipType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"#[0-9a-fA-F]{6}",
        },
    )
    units: None | StandardUnitsType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class SourcePropertyType(Child):
    class Meta:
        name = "sourceProperty.type"

    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    value: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    units: None | StandardUnitsType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class TableType(Child):
    class Meta:
        name = "table.type"

    value: str = field(default="")
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    apply: None | TableTypeApply = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": TableTypeApply.ABSOLUTE,
        },
    )
    interpolation: None | TableTypeInterpolation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": TableTypeInterpolation.LINEAR,
        },
    )


@dataclass(slots=True, kw_only=True)
class TraceType(Child):
    class Meta:
        name = "trace.type"

    value: str = field(default="")
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    type_value: None | TraceTypeType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": TraceTypeType.PEN_DOWN,
        },
    )
    continuation: None | TraceTypeContinuation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": TraceTypeContinuation.BEGIN,
        },
    )
    prior_ref: None | str = field(
        default=None,
        metadata={
            "name": "priorRef",
            "type": "Attribute",
        },
    )
    context_ref: None | str = field(
        default=None,
        metadata={
            "name": "contextRef",
            "type": "Attribute",
        },
    )
    brush_ref: None | str = field(
        default=None,
        metadata={
            "name": "brushRef",
            "type": "Attribute",
        },
    )
    duration: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    time_offset: None | Decimal = field(
        default=None,
        metadata={
            "name": "timeOffset",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class ChannelPropertiesTypeChannelProperty(Child):
    class Meta:
        global_type = False

    channel: None | StandardChannelNameType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    name: None | StandardChannelPropertyNameType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    value: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    units: None | StandardUnitsType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class TraceView(TraceViewType):
    class Meta:
        name = "traceView"
        namespace = "http://www.w3.org/2003/InkML"


@dataclass(slots=True, kw_only=True)
class BrushType(Child):
    class Meta:
        name = "brush.type"

    brush_property: list[BrushPropertyType] = field(
        default_factory=ChildList,
        metadata={
            "name": "brushProperty",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    annotation: list[AnnotationType] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    annotation_xml: list[AnnotationXMLType] = field(
        default_factory=ChildList,
        metadata={
            "name": "annotationXML",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    brush_ref: None | str = field(
        default=None,
        metadata={
            "name": "brushRef",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class ChannelPropertiesType(Child):
    class Meta:
        name = "channelProperties.type"

    channel_property: list[ChannelPropertiesTypeChannelProperty] = field(
        default_factory=ChildList,
        metadata={
            "name": "channelProperty",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )


@dataclass(slots=True, kw_only=True)
class MappingType(Child):
    class Meta:
        name = "mapping.type"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    type_value: None | MappingTypeType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": MappingTypeType.UNKNOWN,
        },
    )
    mapping_ref: None | str = field(
        default=None,
        metadata={
            "name": "mappingRef",
            "type": "Attribute",
        },
    )
    content: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
            "mixed": True,
            "choices": (
                {
                    "name": "bind",
                    "type": ForwardRef("BindType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "table",
                    "type": ForwardRef("TableType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "affine",
                    "type": ForwardRef("AffineType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "math",
                    "type": ForwardRef("Math"),
                    "namespace": "http://www.w3.org/1998/Math/MathML",
                },
                {
                    "name": "mapping",
                    "type": ForwardRef("MappingType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class TraceGroupType(Child):
    class Meta:
        name = "traceGroup.type"

    trace: list[TraceType] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "sequence": 1,
        },
    )
    trace_group: list[TraceGroupType] = field(
        default_factory=ChildList,
        metadata={
            "name": "traceGroup",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "sequence": 1,
        },
    )
    trace_view: list[TraceViewType] = field(
        default_factory=ChildList,
        metadata={
            "name": "traceView",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "sequence": 1,
        },
    )
    annotation: list[AnnotationType] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "sequence": 1,
        },
    )
    annotation_xml: list[AnnotationXMLType] = field(
        default_factory=ChildList,
        metadata={
            "name": "annotationXML",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "sequence": 1,
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    context_ref: None | str = field(
        default=None,
        metadata={
            "name": "contextRef",
            "type": "Attribute",
        },
    )
    brush_ref: None | str = field(
        default=None,
        metadata={
            "name": "brushRef",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Trace(TraceType):
    class Meta:
        name = "trace"
        namespace = "http://www.w3.org/2003/InkML"


@dataclass(slots=True, kw_only=True)
class CanvasTransformType(Child):
    class Meta:
        name = "canvasTransform.type"

    mapping: list[MappingType] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "min_occurs": 1,
            "max_occurs": 2,
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    invertible: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class ChannelType(Child):
    class Meta:
        name = "channel.type"

    mapping: None | MappingType = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    name: None | StandardChannelNameType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    type_value: None | ChannelTypeType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "schema_default": ChannelTypeType.DECIMAL,
        },
    )
    default: None | Decimal | BooleanStrType = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    min: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    max: None | Decimal = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    orientation: None | ChannelTypeOrientation = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": ChannelTypeOrientation.VE,
        },
    )
    respect_to: None | str = field(
        default=None,
        metadata={
            "name": "respectTo",
            "type": "Attribute",
        },
    )
    units: None | StandardUnitsType | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class IntermittentChannelsType(Child):
    class Meta:
        name = "intermittentChannels.type"

    channel: list[ChannelType] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
            "min_occurs": 1,
        },
    )


@dataclass(slots=True, kw_only=True)
class TraceFormatType(Child):
    class Meta:
        name = "traceFormat.type"

    channel: list[ChannelType] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    intermittent_channels: None | IntermittentChannelsType = field(
        default=None,
        metadata={
            "name": "intermittentChannels",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )


@dataclass(slots=True, kw_only=True)
class CanvasType(Child):
    class Meta:
        name = "canvas.type"

    trace_format: None | TraceFormatType = field(
        default=None,
        metadata={
            "name": "traceFormat",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    trace_format_ref: None | str = field(
        default=None,
        metadata={
            "name": "traceFormatRef",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class InkSourceType(Child):
    class Meta:
        name = "inkSource.type"

    trace_format: None | TraceFormatType = field(
        default=None,
        metadata={
            "name": "traceFormat",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    sample_rate: None | SampleRateType = field(
        default=None,
        metadata={
            "name": "sampleRate",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    latency: None | LatencyType = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    active_area: None | ActiveAreaType = field(
        default=None,
        metadata={
            "name": "activeArea",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    source_property: list[SourcePropertyType] = field(
        default_factory=ChildList,
        metadata={
            "name": "sourceProperty",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    channel_properties: None | ChannelPropertiesType = field(
        default=None,
        metadata={
            "name": "channelProperties",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    manufacturer: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    model: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    serial_no: None | str = field(
        default=None,
        metadata={
            "name": "serialNo",
            "type": "Attribute",
        },
    )
    specification_ref: None | str = field(
        default=None,
        metadata={
            "name": "specificationRef",
            "type": "Attribute",
        },
    )
    description: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class ContextType(Child):
    class Meta:
        name = "context.type"

    canvas: None | CanvasType = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    canvas_transform: None | CanvasTransformType = field(
        default=None,
        metadata={
            "name": "canvasTransform",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    trace_format: None | TraceFormatType = field(
        default=None,
        metadata={
            "name": "traceFormat",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    ink_source: None | InkSourceType = field(
        default=None,
        metadata={
            "name": "inkSource",
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    brush: None | BrushType = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    timestamp: None | TimestampType = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    context_ref: None | str = field(
        default=None,
        metadata={
            "name": "contextRef",
            "type": "Attribute",
        },
    )
    canvas_ref: None | str = field(
        default=None,
        metadata={
            "name": "canvasRef",
            "type": "Attribute",
        },
    )
    canvas_transform_ref: None | str = field(
        default=None,
        metadata={
            "name": "canvasTransformRef",
            "type": "Attribute",
        },
    )
    trace_format_ref: None | str = field(
        default=None,
        metadata={
            "name": "traceFormatRef",
            "type": "Attribute",
        },
    )
    ink_source_ref: None | str = field(
        default=None,
        metadata={
            "name": "inkSourceRef",
            "type": "Attribute",
        },
    )
    brush_ref: None | str = field(
        default=None,
        metadata={
            "name": "brushRef",
            "type": "Attribute",
        },
    )
    timestamp_ref: None | str = field(
        default=None,
        metadata={
            "name": "timestampRef",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class DefinitionsType(Child):
    class Meta:
        name = "definitions.type"

    content: list[
        BrushType
        | CanvasType
        | CanvasTransformType
        | ContextType
        | InkSourceType
        | MappingType
        | TimestampType
        | TraceType
        | TraceFormatType
        | TraceGroupType
        | TraceViewType
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "brush",
                    "type": ForwardRef("BrushType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "canvas",
                    "type": ForwardRef("CanvasType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "canvasTransform",
                    "type": ForwardRef("CanvasTransformType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "context",
                    "type": ForwardRef("ContextType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "inkSource",
                    "type": ForwardRef("InkSourceType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "mapping",
                    "type": ForwardRef("MappingType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "timestamp",
                    "type": ForwardRef("TimestampType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "trace",
                    "type": ForwardRef("TraceType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "traceFormat",
                    "type": ForwardRef("TraceFormatType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "traceGroup",
                    "type": ForwardRef("TraceGroupType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "traceView",
                    "type": ForwardRef("TraceViewType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class InkType(Child):
    class Meta:
        name = "ink.type"

    content: list[
        DefinitionsType
        | ContextType
        | TraceType
        | TraceGroupType
        | TraceViewType
        | AnnotationType
        | AnnotationXMLType
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "definitions",
                    "type": ForwardRef("DefinitionsType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "context",
                    "type": ForwardRef("ContextType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "trace",
                    "type": ForwardRef("TraceType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "traceGroup",
                    "type": ForwardRef("TraceGroupType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "traceView",
                    "type": ForwardRef("TraceViewType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "annotation",
                    "type": ForwardRef("AnnotationType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "annotationXML",
                    "type": ForwardRef("AnnotationXMLType"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
            ),
        },
    )
    document_id: None | str = field(
        default=None,
        metadata={
            "name": "documentID",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Definitions(DefinitionsType):
    class Meta:
        name = "definitions"
        namespace = "http://www.w3.org/2003/InkML"


@dataclass(slots=True, kw_only=True)
class Ink(InkType):
    class Meta:
        name = "ink"
        namespace = "http://www.w3.org/2003/InkML"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.inkml import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.inkml.el",
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
