from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2008/diagram"


@dataclass(slots=True, kw_only=True)
class CTDataModelExtBlock(Child):
    class Meta:
        name = "CT_DataModelExtBlock"

    rel_id: None | str = field(
        default=None,
        metadata={
            "name": "relId",
            "type": "Attribute",
        },
    )
    min_ver: None | str = field(
        default=None,
        metadata={
            "name": "minVer",
            "type": "Attribute",
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
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    c_nv_grp_sp_pr: None | CTNonVisualGroupDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
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
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    c_nv_sp_pr: None | CTNonVisualDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvSpPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class DataModelExt(CTDataModelExtBlock):
    class Meta:
        name = "dataModelExt"
        namespace = "http://schemas.microsoft.com/office/drawing/2008/diagram"


@dataclass(slots=True, kw_only=True)
class CTShape(Child):
    class Meta:
        name = "CT_Shape"

    nv_sp_pr: None | CTShapeNonVisual = field(
        default=None,
        metadata={
            "name": "nvSpPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    tx_body: None | CTTextBody = field(
        default=None,
        metadata={
            "name": "txBody",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    tx_xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "name": "txXfrm",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    model_id: None | int | str = field(
        default=None,
        metadata={
            "name": "modelId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
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
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    grp_sp_pr: None | CTGroupShapeProperties = field(
        default=None,
        metadata={
            "name": "grpSpPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )
    sp_or_grp_sp: list[CTShape | CTGroupShape] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "sp",
                    "type": ForwardRef("CTShape"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
                },
                {
                    "name": "grpSp",
                    "type": ForwardRef("CTGroupShape"),
                    "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDrawing(Child):
    class Meta:
        name = "CT_Drawing"

    sp_tree: None | CTGroupShape = field(
        default=None,
        metadata={
            "name": "spTree",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2008/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class Drawing(CTDrawing):
    class Meta:
        name = "drawing"
        namespace = "http://schemas.microsoft.com/office/drawing/2008/diagram"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTGroupShapeProperties,
    CTNonVisualDrawingProps,
    CTNonVisualDrawingShapeProps,
    CTNonVisualGroupDrawingShapeProps,
    CTOfficeArtExtensionList,
    CTShapeProperties,
    CTShapeStyle,
    CTTextBody,
    CTTransform2D,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.diagram_2008 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.oart.diagram_2008.el",
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
