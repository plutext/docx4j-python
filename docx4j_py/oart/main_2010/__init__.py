from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList
from docx4j_py.dml.main import (
    CTEffectProperties as MainCteffectProperties,
)
from docx4j_py.dml.main import (
    CTFillProperties as MainCtfillProperties,
)
from docx4j_py.dml.main import (
    CTLineProperties as MainCtlineProperties,
)
from docx4j_py.dml.main import (
    CTScene3D as MainCtscene3D,
)
from docx4j_py.dml.main import (
    CTShape3D as MainCtshape3D,
)

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CTCameraTool(Child):
    class Meta:
        name = "CT_CameraTool"

    cell_range: None | str = field(
        default=None,
        metadata={
            "name": "cellRange",
            "type": "Attribute",
        },
    )
    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCompatExt(Child):
    class Meta:
        name = "CT_CompatExt"

    spid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTIsGvmlCanvas(Child):
    class Meta:
        name = "CT_IsGvmlCanvas"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectBackgroundRemovalBackgroundMark(Child):
    class Meta:
        name = "CT_PictureEffectBackgroundRemovalBackgroundMark"

    x1: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    y1: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    x2: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    y2: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectBackgroundRemovalForegroundMark(Child):
    class Meta:
        name = "CT_PictureEffectBackgroundRemovalForegroundMark"

    x1: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    y1: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    x2: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    y2: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectBlur(Child):
    class Meta:
        name = "CT_PictureEffectBlur"

    radius: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "10",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectBrightnessContrast(Child):
    class Meta:
        name = "CT_PictureEffectBrightnessContrast"

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
class CTPictureEffectCement(Child):
    class Meta:
        name = "CT_PictureEffectCement"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    crack_spacing: None | int = field(
        default=None,
        metadata={
            "name": "crackSpacing",
            "type": "Attribute",
            "schema_default": "24",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectChalkSketch(Child):
    class Meta:
        name = "CT_PictureEffectChalkSketch"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    pressure: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 4,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectColorTemperature(Child):
    class Meta:
        name = "CT_PictureEffectColorTemperature"

    color_temp: None | int = field(
        default=None,
        metadata={
            "name": "colorTemp",
            "type": "Attribute",
            "schema_default": "6500",
            "min_inclusive": 1500,
            "max_inclusive": 11500,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectCrisscrossEtching(Child):
    class Meta:
        name = "CT_PictureEffectCrisscrossEtching"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "75",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    pressure: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "30",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectCutout(Child):
    class Meta:
        name = "CT_PictureEffectCutout"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    number_of_shades: None | int = field(
        default=None,
        metadata={
            "name": "numberOfShades",
            "type": "Attribute",
            "schema_default": "2",
            "min_inclusive": 0,
            "max_inclusive": 6,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectFilmGrain(Child):
    class Meta:
        name = "CT_PictureEffectFilmGrain"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    grain_size: None | int = field(
        default=None,
        metadata={
            "name": "grainSize",
            "type": "Attribute",
            "schema_default": "40",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectGlass(Child):
    class Meta:
        name = "CT_PictureEffectGlass"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    scaling: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "34",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectGlowDiffused(Child):
    class Meta:
        name = "CT_PictureEffectGlowDiffused"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    intensity: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "5",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectGlowEdges(Child):
    class Meta:
        name = "CT_PictureEffectGlowEdges"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "15",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    smoothness: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "3",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectLightScreen(Child):
    class Meta:
        name = "CT_PictureEffectLightScreen"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    grid_size: None | int = field(
        default=None,
        metadata={
            "name": "gridSize",
            "type": "Attribute",
            "schema_default": "4",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectLineDrawing(Child):
    class Meta:
        name = "CT_PictureEffectLineDrawing"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "25",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    pencil_size: None | int = field(
        default=None,
        metadata={
            "name": "pencilSize",
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectMarker(Child):
    class Meta:
        name = "CT_PictureEffectMarker"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    size: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "97",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectMosiaicBubbles(Child):
    class Meta:
        name = "CT_PictureEffectMosiaicBubbles"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    pressure: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "14",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPaintBrush(Child):
    class Meta:
        name = "CT_PictureEffectPaintBrush"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    brush_size: None | int = field(
        default=None,
        metadata={
            "name": "brushSize",
            "type": "Attribute",
            "schema_default": "2",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPaintStrokes(Child):
    class Meta:
        name = "CT_PictureEffectPaintStrokes"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    intensity: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "5",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPastelsSmooth(Child):
    class Meta:
        name = "CT_PictureEffectPastelsSmooth"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    scaling: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "34",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPencilGrayscale(Child):
    class Meta:
        name = "CT_PictureEffectPencilGrayscale"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    pencil_size: None | int = field(
        default=None,
        metadata={
            "name": "pencilSize",
            "type": "Attribute",
            "schema_default": "27",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPencilSketch(Child):
    class Meta:
        name = "CT_PictureEffectPencilSketch"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    pressure: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "22",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPhotocopy(Child):
    class Meta:
        name = "CT_PictureEffectPhotocopy"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "30",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    detail: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "3",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectPlasticWrap(Child):
    class Meta:
        name = "CT_PictureEffectPlasticWrap"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    smoothness: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "5",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectSaturation(Child):
    class Meta:
        name = "CT_PictureEffectSaturation"

    sat: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "100",
            "min_inclusive": 0,
            "max_inclusive": 400000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectSharpenSoften(Child):
    class Meta:
        name = "CT_PictureEffectSharpenSoften"

    amount: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": -100000,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectTexturizer(Child):
    class Meta:
        name = "CT_PictureEffectTexturizer"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    scaling: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "34",
            "min_inclusive": 0,
            "max_inclusive": 100,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffectWatercolorSponge(Child):
    class Meta:
        name = "CT_PictureEffectWatercolorSponge"

    trans: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    brush_size: None | int = field(
        default=None,
        metadata={
            "name": "brushSize",
            "type": "Attribute",
            "schema_default": "2",
            "min_inclusive": 0,
            "max_inclusive": 10,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShadowObscured(Child):
    class Meta:
        name = "CT_ShadowObscured"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextMath(Child):
    class Meta:
        name = "CT_TextMath"


@dataclass(slots=True, kw_only=True)
class CTUseLocalDpi(Child):
    class Meta:
        name = "CT_UseLocalDpi"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTContentPartLocking(Child):
    class Meta:
        name = "CT_ContentPartLocking"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
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
class CTEffectProperties(MainCteffectProperties):
    class Meta:
        name = "hiddenEffects"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CTFillProperties(MainCtfillProperties):
    class Meta:
        name = "hiddenFill"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CTLineProperties(MainCtlineProperties):
    class Meta:
        name = "hiddenLine"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CTPictureEffectBackgroundRemoval(Child):
    class Meta:
        name = "CT_PictureEffectBackgroundRemoval"

    foreground_mark: list[CTPictureEffectBackgroundRemovalForegroundMark] = field(
        default_factory=ChildList,
        metadata={
            "name": "foregroundMark",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    background_mark: list[CTPictureEffectBackgroundRemovalBackgroundMark] = field(
        default_factory=ChildList,
        metadata={
            "name": "backgroundMark",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    t: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    b: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    l: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )
    r: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 100000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScene3D(MainCtscene3D):
    class Meta:
        name = "hiddenScene3d"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CTShape3D(MainCtshape3D):
    class Meta:
        name = "hiddenSp3d"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CameraTool(CTCameraTool):
    class Meta:
        name = "cameraTool"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CompatExt(CTCompatExt):
    class Meta:
        name = "compatExt"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class IsCanvas(CTIsGvmlCanvas):
    class Meta:
        name = "isCanvas"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class M(CTTextMath):
    class Meta:
        name = "m"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class ShadowObscured(CTShadowObscured):
    class Meta:
        name = "shadowObscured"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class UseLocalDpi(CTUseLocalDpi):
    class Meta:
        name = "useLocalDpi"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class CTNonVisualInkContentPartProperties(Child):
    class Meta:
        name = "CT_NonVisualInkContentPartProperties"

    cp_locks: None | CTContentPartLocking = field(
        default=None,
        metadata={
            "name": "cpLocks",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    is_comment: None | bool = field(
        default=None,
        metadata={
            "name": "isComment",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureEffect(Child):
    class Meta:
        name = "CT_PictureEffect"

    content: (
        None
        | CTPictureEffectBlur
        | CTPictureEffectCement
        | CTPictureEffectChalkSketch
        | CTPictureEffectCrisscrossEtching
        | CTPictureEffectCutout
        | CTPictureEffectFilmGrain
        | CTPictureEffectGlass
        | CTPictureEffectGlowDiffused
        | CTPictureEffectGlowEdges
        | CTPictureEffectLightScreen
        | CTPictureEffectLineDrawing
        | CTPictureEffectMarker
        | CTPictureEffectMosiaicBubbles
        | CTPictureEffectPaintStrokes
        | CTPictureEffectPaintBrush
        | CTPictureEffectPastelsSmooth
        | CTPictureEffectPencilGrayscale
        | CTPictureEffectPencilSketch
        | CTPictureEffectPhotocopy
        | CTPictureEffectPlasticWrap
        | CTPictureEffectTexturizer
        | CTPictureEffectWatercolorSponge
        | CTPictureEffectBackgroundRemoval
        | CTPictureEffectBrightnessContrast
        | CTPictureEffectColorTemperature
        | CTPictureEffectSaturation
        | CTPictureEffectSharpenSoften
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "artisticBlur",
                    "type": ForwardRef("CTPictureEffectBlur"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticCement",
                    "type": ForwardRef("CTPictureEffectCement"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticChalkSketch",
                    "type": ForwardRef("CTPictureEffectChalkSketch"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticCrisscrossEtching",
                    "type": ForwardRef("CTPictureEffectCrisscrossEtching"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticCutout",
                    "type": ForwardRef("CTPictureEffectCutout"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticFilmGrain",
                    "type": ForwardRef("CTPictureEffectFilmGrain"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticGlass",
                    "type": ForwardRef("CTPictureEffectGlass"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticGlowDiffused",
                    "type": ForwardRef("CTPictureEffectGlowDiffused"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticGlowEdges",
                    "type": ForwardRef("CTPictureEffectGlowEdges"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticLightScreen",
                    "type": ForwardRef("CTPictureEffectLightScreen"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticLineDrawing",
                    "type": ForwardRef("CTPictureEffectLineDrawing"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticMarker",
                    "type": ForwardRef("CTPictureEffectMarker"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticMosiaicBubbles",
                    "type": ForwardRef("CTPictureEffectMosiaicBubbles"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPaintStrokes",
                    "type": ForwardRef("CTPictureEffectPaintStrokes"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPaintBrush",
                    "type": ForwardRef("CTPictureEffectPaintBrush"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPastelsSmooth",
                    "type": ForwardRef("CTPictureEffectPastelsSmooth"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPencilGrayscale",
                    "type": ForwardRef("CTPictureEffectPencilGrayscale"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPencilSketch",
                    "type": ForwardRef("CTPictureEffectPencilSketch"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPhotocopy",
                    "type": ForwardRef("CTPictureEffectPhotocopy"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticPlasticWrap",
                    "type": ForwardRef("CTPictureEffectPlasticWrap"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticTexturizer",
                    "type": ForwardRef("CTPictureEffectTexturizer"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "artisticWatercolorSponge",
                    "type": ForwardRef("CTPictureEffectWatercolorSponge"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "backgroundRemoval",
                    "type": ForwardRef("CTPictureEffectBackgroundRemoval"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "brightnessContrast",
                    "type": ForwardRef("CTPictureEffectBrightnessContrast"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "colorTemperature",
                    "type": ForwardRef("CTPictureEffectColorTemperature"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "saturation",
                    "type": ForwardRef("CTPictureEffectSaturation"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
                {
                    "name": "sharpenSoften",
                    "type": ForwardRef("CTPictureEffectSharpenSoften"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
                },
            ),
        },
    )
    visible: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlContentPartNonVisual(Child):
    class Meta:
        name = "CT_GvmlContentPartNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    c_nv_content_part_pr: None | CTNonVisualInkContentPartProperties = field(
        default=None,
        metadata={
            "name": "cNvContentPartPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureLayer(Child):
    class Meta:
        name = "CT_PictureLayer"

    img_effect: list[CTPictureEffect] = field(
        default_factory=ChildList,
        metadata={
            "name": "imgEffect",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    embed: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGvmlContentPart(Child):
    class Meta:
        name = "CT_GvmlContentPart"

    nv_content_part_pr: None | CTGvmlContentPartNonVisual = field(
        default=None,
        metadata={
            "name": "nvContentPartPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
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
class CTPhoto(Child):
    class Meta:
        name = "CT_Photo"

    img_layer: None | CTPictureLayer = field(
        default=None,
        metadata={
            "name": "imgLayer",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2010/main",
        },
    )


@dataclass(slots=True, kw_only=True)
class ContentPart(CTGvmlContentPart):
    class Meta:
        name = "contentPart"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


@dataclass(slots=True, kw_only=True)
class ImgProps(CTPhoto):
    class Meta:
        name = "imgProps"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/main"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTNonVisualDrawingProps,
    CTOfficeArtExtensionList,
    CTTransform2D,
    STBlackWhiteMode,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.main_2010 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.main_2010.el",
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
