from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2012/chartStyle"


class STColorStyleMethodEnum(Enum):
    CYCLE = "cycle"
    WITHIN_LINEAR = "withinLinear"
    ACROSS_LINEAR = "acrossLinear"
    WITHIN_LINEAR_REVERSED = "withinLinearReversed"
    ACROSS_LINEAR_REVERSED = "acrossLinearReversed"


class STMarkerStyle(Enum):
    CIRCLE = "circle"
    DASH = "dash"
    DIAMOND = "diamond"
    DOT = "dot"
    PLUS = "plus"
    SQUARE = "square"
    STAR = "star"
    TRIANGLE = "triangle"
    X = "x"


class STStyleColorEnum(Enum):
    AUTO = "auto"


class STStyleEntryModifierEnum(Enum):
    ALLOW_NO_FILL_OVERRIDE = "allowNoFillOverride"
    ALLOW_NO_LINE_OVERRIDE = "allowNoLineOverride"


class STStyleReferenceModifierEnum(Enum):
    IGNORE_CSTRANSFORMS = "ignoreCSTransforms"


@dataclass(slots=True, kw_only=True)
class CTColorStyleVariation(Child):
    class Meta:
        name = "CT_ColorStyleVariation"

    content: list[
        Tint2
        | Shade2
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha2
        | CTFixedPercentage
        | AlphaMod2
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod2
        | Sat2
        | SatOff2
        | SatMod2
        | Lum2
        | LumOff2
        | LumMod2
        | Red2
        | RedOff2
        | RedMod2
        | Green2
        | GreenOff2
        | GreenMod2
        | Blue2
        | BlueOff2
        | BlueMod2
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade2"),
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
                    "type": ForwardRef("Alpha2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod2"),
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
                    "type": ForwardRef("HueMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod2"),
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


@dataclass(slots=True, kw_only=True)
class CTMarkerLayout(Child):
    class Meta:
        name = "CT_MarkerLayout"

    symbol: None | STMarkerStyle = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    size: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 2,
            "max_inclusive": 72,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleColor(Child):
    class Meta:
        name = "CT_StyleColor"

    content: list[
        Tint2
        | Shade2
        | CTComplementTransform
        | CTInverseTransform
        | CTGrayscaleTransform
        | Alpha2
        | CTFixedPercentage
        | AlphaMod2
        | CTPositiveFixedAngle
        | CTAngle
        | HueMod2
        | Sat2
        | SatOff2
        | SatMod2
        | Lum2
        | LumOff2
        | LumMod2
        | Red2
        | RedOff2
        | RedMod2
        | Green2
        | GreenOff2
        | GreenMod2
        | Blue2
        | BlueOff2
        | BlueMod2
        | CTGammaTransform
        | CTInverseGammaTransform
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "tint",
                    "type": ForwardRef("Tint2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "shade",
                    "type": ForwardRef("Shade2"),
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
                    "type": ForwardRef("Alpha2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaOff",
                    "type": ForwardRef("CTFixedPercentage"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "alphaMod",
                    "type": ForwardRef("AlphaMod2"),
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
                    "type": ForwardRef("HueMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "sat",
                    "type": ForwardRef("Sat2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satOff",
                    "type": ForwardRef("SatOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "satMod",
                    "type": ForwardRef("SatMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lum",
                    "type": ForwardRef("Lum2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumOff",
                    "type": ForwardRef("LumOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "lumMod",
                    "type": ForwardRef("LumMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "red",
                    "type": ForwardRef("Red2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redOff",
                    "type": ForwardRef("RedOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "redMod",
                    "type": ForwardRef("RedMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "green",
                    "type": ForwardRef("Green2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenOff",
                    "type": ForwardRef("GreenOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "greenMod",
                    "type": ForwardRef("GreenMod2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blue",
                    "type": ForwardRef("Blue2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueOff",
                    "type": ForwardRef("BlueOff2"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "blueMod",
                    "type": ForwardRef("BlueMod2"),
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
    val: None | int | STStyleColorEnum | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTColorStyle(Child):
    class Meta:
        name = "CT_ColorStyle"

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
    variation: list[CTColorStyleVariation] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    meth: None | STColorStyleMethodEnum | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
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
    style_clr: None | CTStyleColor = field(
        default=None,
        metadata={
            "name": "styleClr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    idx: None | STFontCollectionIndex = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    mods: list[STStyleReferenceModifierEnum | str] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleReference(Child):
    class Meta:
        name = "CT_StyleReference"

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
    style_clr: None | CTStyleColor = field(
        default=None,
        metadata={
            "name": "styleClr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    idx: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    mods: list[STStyleReferenceModifierEnum | str] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyleEntry(Child):
    class Meta:
        name = "CT_StyleEntry"

    ln_ref: None | CTStyleReference = field(
        default=None,
        metadata={
            "name": "lnRef",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    line_width_scale: None | float = field(
        default=None,
        metadata={
            "name": "lineWidthScale",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    fill_ref: None | CTStyleReference = field(
        default=None,
        metadata={
            "name": "fillRef",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    effect_ref: None | CTStyleReference = field(
        default=None,
        metadata={
            "name": "effectRef",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    font_ref: None | CTFontReference = field(
        default=None,
        metadata={
            "name": "fontRef",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    def_rpr: None | CTTextCharacterProperties = field(
        default=None,
        metadata={
            "name": "defRPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    body_pr: None | CTTextBodyProperties = field(
        default=None,
        metadata={
            "name": "bodyPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    mods: list[STStyleEntryModifierEnum | str] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class ColorStyle(CTColorStyle):
    class Meta:
        name = "colorStyle"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chartStyle"


@dataclass(slots=True, kw_only=True)
class CTChartStyle(Child):
    class Meta:
        name = "CT_ChartStyle"

    axis_title: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "axisTitle",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    category_axis: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "categoryAxis",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    chart_area: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "chartArea",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_label: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataLabel",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_label_callout: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataLabelCallout",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_point: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataPoint",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_point3_d: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataPoint3D",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_point_line: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataPointLine",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_point_marker: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataPointMarker",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_point_marker_layout: None | CTMarkerLayout = field(
        default=None,
        metadata={
            "name": "dataPointMarkerLayout",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_point_wireframe: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataPointWireframe",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    data_table: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dataTable",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    down_bar: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "downBar",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    drop_line: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "dropLine",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    error_bar: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "errorBar",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    floor: None | CTStyleEntry = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    gridline_major: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "gridlineMajor",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    gridline_minor: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "gridlineMinor",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    hi_lo_line: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "hiLoLine",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    leader_line: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "leaderLine",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    legend: None | CTStyleEntry = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    plot_area: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "plotArea",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    plot_area3_d: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "plotArea3D",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    series_axis: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "seriesAxis",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    series_line: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "seriesLine",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    title: None | CTStyleEntry = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    trendline: None | CTStyleEntry = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    trendline_label: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "trendlineLabel",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    up_bar: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "upBar",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    value_axis: None | CTStyleEntry = field(
        default=None,
        metadata={
            "name": "valueAxis",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    wall: None | CTStyleEntry = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class ChartStyle(CTChartStyle):
    class Meta:
        name = "chartStyle"
        namespace = "http://schemas.microsoft.com/office/drawing/2012/chartStyle"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    Alpha2,
    AlphaMod2,
    Blue2,
    BlueMod2,
    BlueOff2,
    CTAngle,
    CTComplementTransform,
    CTFixedPercentage,
    CTGammaTransform,
    CTGrayscaleTransform,
    CTHslColor,
    CTInverseGammaTransform,
    CTInverseTransform,
    CTOfficeArtExtensionList,
    CTPositiveFixedAngle,
    CTPresetColor,
    CTSchemeColor,
    CTScRgbColor,
    CTShapeProperties,
    CTSRgbColor,
    CTSystemColor,
    CTTextBodyProperties,
    CTTextCharacterProperties,
    Green2,
    GreenMod2,
    GreenOff2,
    HueMod2,
    Lum2,
    LumMod2,
    LumOff2,
    Red2,
    RedMod2,
    RedOff2,
    Sat2,
    SatMod2,
    SatOff2,
    Shade2,
    STFontCollectionIndex,
    Tint2,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.chart_style_2012 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.oart.chart_style_2012.el",
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
