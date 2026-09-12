from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


@dataclass(slots=True, kw_only=True)
class CTAnchorClientData(Child):
    class Meta:
        name = "CT_AnchorClientData"

    f_locks_with_sheet: None | bool = field(
        default=None,
        metadata={
            "name": "fLocksWithSheet",
            "type": "Attribute",
            "schema_default": "true",
        },
    )
    f_prints_with_sheet: None | bool = field(
        default=None,
        metadata={
            "name": "fPrintsWithSheet",
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTMarker(Child):
    class Meta:
        name = "CT_Marker"

    col: None | int = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "min_inclusive": 0,
        },
    )
    col_off: None | int = field(
        default=None,
        metadata={
            "name": "colOff",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    row: None | int = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "min_inclusive": 0,
        },
    )
    row_off: None | int = field(
        default=None,
        metadata={
            "name": "rowOff",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


class STEditAs(Enum):
    TWO_CELL = "twoCell"
    ONE_CELL = "oneCell"
    ABSOLUTE = "absolute"


@dataclass(slots=True, kw_only=True)
class CTConnectorNonVisual(Child):
    class Meta:
        name = "CT_ConnectorNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    c_nv_cxn_sp_pr: None | CTNonVisualConnectorProperties = field(
        default=None,
        metadata={
            "name": "cNvCxnSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicalObjectFrameNonVisual(Child):
    class Meta:
        name = "CT_GraphicalObjectFrameNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    c_nv_graphic_frame_pr: None | CTNonVisualGraphicFrameProperties = field(
        default=None,
        metadata={
            "name": "cNvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    c_nv_grp_sp_pr: None | CTNonVisualGroupDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    c_nv_pic_pr: None | CTNonVisualPictureProperties = field(
        default=None,
        metadata={
            "name": "cNvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    c_nv_sp_pr: None | CTNonVisualDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class From(CTMarker):
    class Meta:
        name = "from"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


@dataclass(slots=True, kw_only=True)
class To(CTMarker):
    class Meta:
        name = "to"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


@dataclass(slots=True, kw_only=True)
class CTConnector(Child):
    class Meta:
        name = "CT_Connector"

    nv_cxn_sp_pr: None | CTConnectorNonVisual = field(
        default=None,
        metadata={
            "name": "nvCxnSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
class CTGraphicalObjectFrame(Child):
    class Meta:
        name = "CT_GraphicalObjectFrame"

    nv_graphic_frame_pr: None | CTGraphicalObjectFrameNonVisual = field(
        default=None,
        metadata={
            "name": "nvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    blip_fill: None | CTBlipFillProperties = field(
        default=None,
        metadata={
            "name": "blipFill",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    tx_body: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txBody",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
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
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    grp_sp_pr: None | CTGroupShapeProperties = field(
        default=None,
        metadata={
            "name": "grpSpPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    content: list[CTShape | CTGroupShape | CTGraphicalObjectFrame | CTConnector | CTPicture] = (
        field(
            default_factory=ChildList,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "sp",
                        "type": ForwardRef("CTShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "grpSp",
                        "type": ForwardRef("CTGroupShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "graphicFrame",
                        "type": ForwardRef("CTGraphicalObjectFrame"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "cxnSp",
                        "type": ForwardRef("CTConnector"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "pic",
                        "type": ForwardRef("CTPicture"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                ),
            },
        )
    )


@dataclass(slots=True, kw_only=True)
class CTAbsoluteAnchor(Child):
    class Meta:
        name = "CT_AbsoluteAnchor"

    pos: None | CTPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    ext: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    content: None | CTShape | CTGroupShape | CTGraphicalObjectFrame | CTConnector | CTPicture = (
        field(
            default=None,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "sp",
                        "type": ForwardRef("CTShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "grpSp",
                        "type": ForwardRef("CTGroupShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "graphicFrame",
                        "type": ForwardRef("CTGraphicalObjectFrame"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "cxnSp",
                        "type": ForwardRef("CTConnector"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "pic",
                        "type": ForwardRef("CTPicture"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                ),
            },
        )
    )
    client_data: None | CTAnchorClientData = field(
        default=None,
        metadata={
            "name": "clientData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTOneCellAnchor(Child):
    class Meta:
        name = "CT_OneCellAnchor"

    from_value: None | CTMarker = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    ext: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    content: None | CTShape | CTGroupShape | CTGraphicalObjectFrame | CTConnector | CTPicture = (
        field(
            default=None,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "sp",
                        "type": ForwardRef("CTShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "grpSp",
                        "type": ForwardRef("CTGroupShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "graphicFrame",
                        "type": ForwardRef("CTGraphicalObjectFrame"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "cxnSp",
                        "type": ForwardRef("CTConnector"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "pic",
                        "type": ForwardRef("CTPicture"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                ),
            },
        )
    )
    client_data: None | CTAnchorClientData = field(
        default=None,
        metadata={
            "name": "clientData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTwoCellAnchor(Child):
    class Meta:
        name = "CT_TwoCellAnchor"

    from_value: None | CTMarker = field(
        default=None,
        metadata={
            "name": "from",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    to: None | CTMarker = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    content: None | CTShape | CTGroupShape | CTGraphicalObjectFrame | CTConnector | CTPicture = (
        field(
            default=None,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "sp",
                        "type": ForwardRef("CTShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "grpSp",
                        "type": ForwardRef("CTGroupShape"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "graphicFrame",
                        "type": ForwardRef("CTGraphicalObjectFrame"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "cxnSp",
                        "type": ForwardRef("CTConnector"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                    {
                        "name": "pic",
                        "type": ForwardRef("CTPicture"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                    },
                ),
            },
        )
    )
    client_data: None | CTAnchorClientData = field(
        default=None,
        metadata={
            "name": "clientData",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        },
    )
    edit_as: None | STEditAs = field(
        default=None,
        metadata={
            "name": "editAs",
            "type": "Attribute",
            "schema_default": STEditAs.TWO_CELL,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDrawing(Child):
    class Meta:
        name = "CT_Drawing"

    two_cell_anchor_or_one_cell_anchor_or_absolute_anchor: list[
        CTTwoCellAnchor | CTOneCellAnchor | CTAbsoluteAnchor
    ] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "twoCellAnchor",
                    "type": ForwardRef("CTTwoCellAnchor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                },
                {
                    "name": "oneCellAnchor",
                    "type": ForwardRef("CTOneCellAnchor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                },
                {
                    "name": "absoluteAnchor",
                    "type": ForwardRef("CTAbsoluteAnchor"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                },
            ),
        },
    )


@dataclass(slots=True, kw_only=True)
class WsDr(CTDrawing):
    class Meta:
        name = "wsDr"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


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
    CTPoint2D,
    CTPositiveSize2D,
    CTShapeProperties,
    CTShapeStyle,
    CTTextBody,
    CTTransform2D,
    Graphic,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.spreadsheet_drawing import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.dml.spreadsheet_drawing.el",
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
