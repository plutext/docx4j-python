from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"


@dataclass(slots=True, kw_only=True)
class CTMarker(Child):
    class Meta:
        name = "CT_Marker"

    x: None | float = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
            "min_inclusive": 0.0,
            "max_inclusive": 1.0,
        },
    )
    y: None | float = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
            "min_inclusive": 0.0,
            "max_inclusive": 1.0,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConnectorNonVisual(Child):
    class Meta:
        name = "CT_ConnectorNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    c_nv_cxn_sp_pr: None | CTNonVisualConnectorProperties = field(
        default=None,
        metadata={
            "name": "cNvCxnSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicFrameNonVisual(Child):
    class Meta:
        name = "CT_GraphicFrameNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    c_nv_graphic_frame_pr: None | CTNonVisualGraphicFrameProperties = field(
        default=None,
        metadata={
            "name": "cNvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupShapeNonVisual(Child):
    class Meta:
        name = "CT_GroupShapeNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    c_nv_grp_sp_pr: None | CTNonVisualGroupDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureNonVisual(Child):
    class Meta:
        name = "CT_PictureNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    c_nv_pic_pr: None | CTNonVisualPictureProperties = field(
        default=None,
        metadata={
            "name": "cNvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeNonVisual(Child):
    class Meta:
        name = "CT_ShapeNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    c_nv_sp_pr: None | CTNonVisualDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConnector(Child):
    class Meta:
        name = "CT_Connector"

    nv_cxn_sp_pr: None | CTConnectorNonVisual = field(
        default=None,
        metadata={
            "name": "nvCxnSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    macro: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    f_published: None | bool = field(
        default=None,
        metadata={
            "name": "fPublished",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicFrame(Child):
    class Meta:
        name = "CT_GraphicFrame"

    nv_graphic_frame_pr: None | CTGraphicFrameNonVisual = field(
        default=None,
        metadata={
            "name": "nvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    graphic: None | Graphic = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    macro: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    f_published: None | bool = field(
        default=None,
        metadata={
            "name": "fPublished",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPicture(Child):
    class Meta:
        name = "CT_Picture"

    nv_pic_pr: None | CTPictureNonVisual = field(
        default=None,
        metadata={
            "name": "nvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    blip_fill: None | CTBlipFillProperties = field(
        default=None,
        metadata={
            "name": "blipFill",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    macro: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )
    f_published: None | bool = field(
        default=None,
        metadata={
            "name": "fPublished",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShape(Child):
    class Meta:
        name = "CT_Shape"

    nv_sp_pr: None | CTShapeNonVisual = field(
        default=None,
        metadata={
            "name": "nvSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    tx_body: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txBody",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    macro: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    textlink: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    f_locks_text: None | bool = field(
        default=None,
        metadata={
            "name": "fLocksText",
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    f_published: None | bool = field(
        default=None,
        metadata={
            "name": "fPublished",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupShape(Child):
    class Meta:
        name = "CT_GroupShape"

    nv_grp_sp_pr: None | CTGroupShapeNonVisual = field(
        default=None,
        metadata={
            "name": "nvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    grp_sp_pr: None | CTGroupShapeProperties = field(
        default=None,
        metadata={
            "name": "grpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    content: list[CTShape | CTGroupShape | CTGraphicFrame | CTConnector | CTPicture] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "sp",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "grpSp",
                    "type": ForwardRef("CTGroupShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "graphicFrame",
                    "type": ForwardRef("CTGraphicFrame"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "cxnSp",
                    "type": ForwardRef("CTConnector"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "pic",
                    "type": ForwardRef("CTPicture"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAbsSizeAnchor(Child):
    class Meta:
        name = "CT_AbsSizeAnchor"

    from_value: None | CTMarker = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    ext: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    content: None | CTShape | CTGroupShape | CTGraphicFrame | CTConnector | CTPicture = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "sp",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "grpSp",
                    "type": ForwardRef("CTGroupShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "graphicFrame",
                    "type": ForwardRef("CTGraphicFrame"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "cxnSp",
                    "type": ForwardRef("CTConnector"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "pic",
                    "type": ForwardRef("CTPicture"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTRelSizeAnchor(Child):
    class Meta:
        name = "CT_RelSizeAnchor"

    from_value: None | CTMarker = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    to: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
        },
    )
    content: None | CTShape | CTGroupShape | CTGraphicFrame | CTConnector | CTPicture = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "sp",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "grpSp",
                    "type": ForwardRef("CTGroupShape"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "graphicFrame",
                    "type": ForwardRef("CTGraphicFrame"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "cxnSp",
                    "type": ForwardRef("CTConnector"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "pic",
                    "type": ForwardRef("CTPicture"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDrawing(Child):
    class Meta:
        name = "CT_Drawing"

    rel_size_anchor_or_abs_size_anchor: list[CTRelSizeAnchor | CTAbsSizeAnchor] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "relSizeAnchor",
                    "type": ForwardRef("CTRelSizeAnchor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
                {
                    "name": "absSizeAnchor",
                    "type": ForwardRef("CTAbsSizeAnchor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                },
            ),
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTBlipFillProperties,
    CTGroupShapeProperties,
    CTNonVisualConnectorProperties,
    CTNonVisualDrawingProps,
    CTNonVisualDrawingShapeProps,
    CTNonVisualGraphicFrameProperties,
    CTNonVisualGroupDrawingShapeProps,
    CTNonVisualPictureProperties,
    CTPositiveSize2D,
    CTShapeProperties,
    CTShapeStyle,
    CTTextBody,
    CTTransform2D,
    Graphic,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.chart_drawing import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.dml.chart_drawing.el",
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
