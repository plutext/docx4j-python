from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"


@dataclass(slots=True, kw_only=True)
class CTGraphicFrame(Child):
    class Meta:
        name = "CT_GraphicFrame"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )
    c_nv_fr_pr: None | CTNonVisualGraphicFrameProperties = field(
        default=None,
        metadata={
            "name": "cNvFrPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )
    graphic: None | Graphic = field(
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
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWordprocessingGroup(Child):
    class Meta:
        name = "CT_WordprocessingGroup"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )
    c_nv_grp_sp_pr: None | CTNonVisualGroupDrawingShapeProps = field(
        default=None,
        metadata={
            "name": "cNvGrpSpPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )
    grp_sp_pr: None | CTGroupShapeProperties = field(
        default=None,
        metadata={
            "name": "grpSpPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )
    content: list[Wsp | CTWordprocessingGroup | CTGraphicFrame | Pic] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "wsp",
                    "type": ForwardRef("Wsp"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                },
                {
                    "name": "grpSp",
                    "type": ForwardRef("CTWordprocessingGroup"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                },
                {
                    "name": "graphicFrame",
                    "type": ForwardRef("CTGraphicFrame"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                },
                {
                    "name": "pic",
                    "type": ForwardRef("Pic"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        },
    )


@dataclass(slots=True, kw_only=True)
class Wgp(CTWordprocessingGroup):
    class Meta:
        name = "wgp"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTGroupShapeProperties,
    CTNonVisualDrawingProps,
    CTNonVisualGraphicFrameProperties,
    CTNonVisualGroupDrawingShapeProps,
    CTOfficeArtExtensionList,
    CTTransform2D,
    Graphic,
)
from docx4j_py.dml.picture import Pic
from docx4j_py.word.wps import Wsp


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.word.wpg import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.word.wpg.el",
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
