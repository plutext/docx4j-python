from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2017/model3d"


@dataclass(slots=True, kw_only=True)
class CTPositiveRatio(Child):
    class Meta:
        name = "CT_PositiveRatio"

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
class CTRotate3D(Child):
    class Meta:
        name = "CT_Rotate3D"

    ax: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    ay: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    az: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTUnknownLight(Child):
    class Meta:
        name = "CT_UnknownLight"


@dataclass(slots=True, kw_only=True)
class CTAmbientLight(Child):
    class Meta:
        name = "CT_AmbientLight"

    clr: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    illuminance: None | CTPositiveRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    enabled: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDirectionalLight(Child):
    class Meta:
        name = "CT_DirectionalLight"

    clr: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    illuminance: None | CTPositiveRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    pos: None | CTPoint3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    look_at: None | CTPoint3D = field(
        default=None,
        metadata={
            "name": "lookAt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    enabled: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    angular_rad: None | int = field(
        default=None,
        metadata={
            "name": "angularRad",
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 5400000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTModel3DRaster(Child):
    class Meta:
        name = "CT_Model3DRaster"

    blip: None | CTBlip = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    r_name: None | str = field(
        default=None,
        metadata={
            "name": "rName",
            "type": "Attribute",
        },
    )
    r_ver: None | str = field(
        default=None,
        metadata={
            "name": "rVer",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTObjectViewport(Child):
    class Meta:
        name = "CT_ObjectViewport"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    viewport_sz: None | int = field(
        default=None,
        metadata={
            "name": "viewportSz",
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOrthographicProjection(Child):
    class Meta:
        name = "CT_OrthographicProjection"

    sz: None | CTPositiveRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPerspectiveProjection(Child):
    class Meta:
        name = "CT_PerspectiveProjection"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
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


@dataclass(slots=True, kw_only=True)
class CTPointLight(Child):
    class Meta:
        name = "CT_PointLight"

    clr: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    intensity: None | CTPositiveRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    pos: None | CTPoint3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    enabled: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTScale3D(Child):
    class Meta:
        name = "CT_Scale3D"

    sx: None | CTRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    sy: None | CTRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    sz: None | CTRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSpotLight(Child):
    class Meta:
        name = "CT_SpotLight"

    clr: None | CTColor = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    intensity: None | CTPositiveRatio = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    pos: None | CTPoint3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    look_at: None | CTPoint3D = field(
        default=None,
        metadata={
            "name": "lookAt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    enabled: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    rad: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 27273042316900,
        },
    )
    spot_ang: None | int = field(
        default=None,
        metadata={
            "name": "spotAng",
            "type": "Attribute",
            "min_inclusive": 0,
            "max_inclusive": 10800000,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWindowViewport(Child):
    class Meta:
        name = "CT_WindowViewport"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTModel3DCamera(Child):
    class Meta:
        name = "CT_Model3DCamera"

    pos: None | CTPoint3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    up: None | CTVector3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    look_at: None | CTPoint3D = field(
        default=None,
        metadata={
            "name": "lookAt",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    orthographic_or_perspective: None | CTOrthographicProjection | CTPerspectiveProjection = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "orthographic",
                    "type": ForwardRef("CTOrthographicProjection"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
                {
                    "name": "perspective",
                    "type": ForwardRef("CTPerspectiveProjection"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTModel3DTransform(Child):
    class Meta:
        name = "CT_Model3DTransform"

    meter_per_model_unit: None | CTPositiveRatio = field(
        default=None,
        metadata={
            "name": "meterPerModelUnit",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    pre_trans: None | CTVector3D = field(
        default=None,
        metadata={
            "name": "preTrans",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    scale: None | CTScale3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    rot: None | CTRotate3D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    post_trans: None | CTVector3D = field(
        default=None,
        metadata={
            "name": "postTrans",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTModel3D(Child):
    class Meta:
        name = "CT_Model3D"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    camera: None | CTModel3DCamera = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    trans: None | CTModel3DTransform = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    attr_src_url: None | CTPictureAttributionSourceURL = field(
        default=None,
        metadata={
            "name": "attrSrcUrl",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    raster: None | CTModel3DRaster = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    obj_viewport_or_win_viewport: None | CTObjectViewport | CTWindowViewport = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "objViewport",
                    "type": ForwardRef("CTObjectViewport"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
                {
                    "name": "winViewport",
                    "type": ForwardRef("CTWindowViewport"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
            ),
        },
    )
    ambient_light: None | CTAmbientLight = field(
        default=None,
        metadata={
            "name": "ambientLight",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        },
    )
    content: list[CTPointLight | CTSpotLight | CTDirectionalLight | CTUnknownLight] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "ptLight",
                    "type": ForwardRef("CTPointLight"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
                {
                    "name": "spotLight",
                    "type": ForwardRef("CTSpotLight"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
                {
                    "name": "dirLight",
                    "type": ForwardRef("CTDirectionalLight"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
                {
                    "name": "unkLight",
                    "type": ForwardRef("CTUnknownLight"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2017/model3d",
                },
            ),
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


@dataclass(slots=True, kw_only=True)
class Model3D(CTModel3D):
    class Meta:
        name = "model3D"
        namespace = "http://schemas.microsoft.com/office/drawing/2017/model3d"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTBlip,
    CTColor,
    CTOfficeArtExtensionList,
    CTPoint3D,
    CTRatio,
    CTShapeProperties,
    CTVector3D,
)
from docx4j_py.oart.main_2016_11 import CTPictureAttributionSourceURL


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.model3d_2017 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.oart.model3d_2017.el",
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
