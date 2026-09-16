from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


@dataclass(slots=True, kw_only=True)
class CTEffectExtent(Child):
    class Meta:
        name = "CT_EffectExtent"

    l: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    t: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    r: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )
    b: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": -27273042329600,
            "max_inclusive": 27273042316900,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrapNone(Child):
    class Meta:
        name = "CT_WrapNone"


class STAlignH(Enum):
    LEFT = "left"
    RIGHT = "right"
    CENTER = "center"
    INSIDE = "inside"
    OUTSIDE = "outside"


class STAlignV(Enum):
    TOP = "top"
    BOTTOM = "bottom"
    CENTER = "center"
    INSIDE = "inside"
    OUTSIDE = "outside"


class STRelFromH(Enum):
    MARGIN = "margin"
    PAGE = "page"
    COLUMN = "column"
    CHARACTER = "character"
    LEFT_MARGIN = "leftMargin"
    RIGHT_MARGIN = "rightMargin"
    INSIDE_MARGIN = "insideMargin"
    OUTSIDE_MARGIN = "outsideMargin"


class STRelFromV(Enum):
    MARGIN = "margin"
    PAGE = "page"
    PARAGRAPH = "paragraph"
    LINE = "line"
    TOP_MARGIN = "topMargin"
    BOTTOM_MARGIN = "bottomMargin"
    INSIDE_MARGIN = "insideMargin"
    OUTSIDE_MARGIN = "outsideMargin"


class STWrapText(Enum):
    BOTH_SIDES = "bothSides"
    LEFT = "left"
    RIGHT = "right"
    LARGEST = "largest"


@dataclass(slots=True, kw_only=True)
class CTPosH(Child):
    class Meta:
        name = "CT_PosH"

    align_or_pos_offset: None | STAlignH | int = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "align",
                    "type": ForwardRef("STAlignH"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                },
                {
                    "name": "posOffset",
                    "type": int,
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                },
            ),
        },
    )
    relative_from: None | STRelFromH = field(
        default=None,
        metadata={
            "name": "relativeFrom",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPosV(Child):
    class Meta:
        name = "CT_PosV"

    align_or_pos_offset: None | STAlignV | int = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "align",
                    "type": ForwardRef("STAlignV"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                },
                {
                    "name": "posOffset",
                    "type": int,
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                },
            ),
        },
    )
    relative_from: None | STRelFromV = field(
        default=None,
        metadata={
            "name": "relativeFrom",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrapPath(Child):
    class Meta:
        name = "CT_WrapPath"

    start: None | CTPoint2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    line_to: list[CTPoint2D] = field(
        default_factory=ChildList,
        metadata={
            "name": "lineTo",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "min_occurs": 2,
        },
    )
    edited: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrapSquare(Child):
    class Meta:
        name = "CT_WrapSquare"

    effect_extent: None | CTEffectExtent = field(
        default=None,
        metadata={
            "name": "effectExtent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    wrap_text: None | STWrapText = field(
        default=None,
        metadata={
            "name": "wrapText",
            "type": "Attribute",
        },
    )
    dist_t: None | int = field(
        default=None,
        metadata={
            "name": "distT",
            "type": "Attribute",
        },
    )
    dist_b: None | int = field(
        default=None,
        metadata={
            "name": "distB",
            "type": "Attribute",
        },
    )
    dist_l: None | int = field(
        default=None,
        metadata={
            "name": "distL",
            "type": "Attribute",
        },
    )
    dist_r: None | int = field(
        default=None,
        metadata={
            "name": "distR",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrapTopBottom(Child):
    class Meta:
        name = "CT_WrapTopBottom"

    effect_extent: None | CTEffectExtent = field(
        default=None,
        metadata={
            "name": "effectExtent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    dist_t: None | int = field(
        default=None,
        metadata={
            "name": "distT",
            "type": "Attribute",
        },
    )
    dist_b: None | int = field(
        default=None,
        metadata={
            "name": "distB",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Inline1(Child):
    class Meta:
        name = "CT_Inline"

    extent: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    effect_extent: None | CTEffectExtent = field(
        default=None,
        metadata={
            "name": "effectExtent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    doc_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "docPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    c_nv_graphic_frame_pr: None | CTNonVisualGraphicFrameProperties = field(
        default=None,
        metadata={
            "name": "cNvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    graphic: None | Graphic = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    dist_t: None | int = field(
        default=None,
        metadata={
            "name": "distT",
            "type": "Attribute",
        },
    )
    dist_b: None | int = field(
        default=None,
        metadata={
            "name": "distB",
            "type": "Attribute",
        },
    )
    dist_l: None | int = field(
        default=None,
        metadata={
            "name": "distL",
            "type": "Attribute",
        },
    )
    dist_r: None | int = field(
        default=None,
        metadata={
            "name": "distR",
            "type": "Attribute",
        },
    )
    anchor_id: None | bytes = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "length": 4,
            "format": "base16",
        },
    )
    edit_id: None | bytes = field(
        default=None,
        metadata={
            "name": "editId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "length": 4,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrapThrough(Child):
    class Meta:
        name = "CT_WrapThrough"

    wrap_polygon: None | CTWrapPath = field(
        default=None,
        metadata={
            "name": "wrapPolygon",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    wrap_text: None | STWrapText = field(
        default=None,
        metadata={
            "name": "wrapText",
            "type": "Attribute",
        },
    )
    dist_l: None | int = field(
        default=None,
        metadata={
            "name": "distL",
            "type": "Attribute",
        },
    )
    dist_r: None | int = field(
        default=None,
        metadata={
            "name": "distR",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrapTight(Child):
    class Meta:
        name = "CT_WrapTight"

    wrap_polygon: None | CTWrapPath = field(
        default=None,
        metadata={
            "name": "wrapPolygon",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    wrap_text: None | STWrapText = field(
        default=None,
        metadata={
            "name": "wrapText",
            "type": "Attribute",
        },
    )
    dist_l: None | int = field(
        default=None,
        metadata={
            "name": "distL",
            "type": "Attribute",
        },
    )
    dist_r: None | int = field(
        default=None,
        metadata={
            "name": "distR",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Inline(Inline1):
    class Meta:
        name = "inline"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


@dataclass(slots=True, kw_only=True)
class Anchor1(Child):
    class Meta:
        name = "CT_Anchor"

    simple_pos: None | CTPoint2D = field(
        default=None,
        metadata={
            "name": "simplePos",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    position_h: None | CTPosH = field(
        default=None,
        metadata={
            "name": "positionH",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    position_v: None | CTPosV = field(
        default=None,
        metadata={
            "name": "positionV",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    extent: None | CTPositiveSize2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    effect_extent: None | CTEffectExtent = field(
        default=None,
        metadata={
            "name": "effectExtent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    content: None | CTWrapNone | CTWrapSquare | CTWrapTight | CTWrapThrough | CTWrapTopBottom = (
        field(
            default=None,
            metadata={
                "type": "Elements",
                "choices": (
                    {
                        "name": "wrapNone",
                        "type": ForwardRef("CTWrapNone"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    },
                    {
                        "name": "wrapSquare",
                        "type": ForwardRef("CTWrapSquare"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    },
                    {
                        "name": "wrapTight",
                        "type": ForwardRef("CTWrapTight"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    },
                    {
                        "name": "wrapThrough",
                        "type": ForwardRef("CTWrapThrough"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    },
                    {
                        "name": "wrapTopAndBottom",
                        "type": ForwardRef("CTWrapTopBottom"),
                        "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    },
                ),
            },
        )
    )
    doc_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "docPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    c_nv_graphic_frame_pr: None | CTNonVisualGraphicFrameProperties = field(
        default=None,
        metadata={
            "name": "cNvGraphicFramePr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        },
    )
    graphic: None | Graphic = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
        },
    )
    size_rel_h: None | SizeRelH = field(
        default=None,
        metadata={
            "name": "sizeRelH",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        },
    )
    size_rel_v: None | SizeRelV = field(
        default=None,
        metadata={
            "name": "sizeRelV",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        },
    )
    dist_t: None | int = field(
        default=None,
        metadata={
            "name": "distT",
            "type": "Attribute",
        },
    )
    dist_b: None | int = field(
        default=None,
        metadata={
            "name": "distB",
            "type": "Attribute",
        },
    )
    dist_l: None | int = field(
        default=None,
        metadata={
            "name": "distL",
            "type": "Attribute",
        },
    )
    dist_r: None | int = field(
        default=None,
        metadata={
            "name": "distR",
            "type": "Attribute",
        },
    )
    simple_pos_attribute: None | bool = field(
        default=None,
        metadata={
            "name": "simplePos",
            "type": "Attribute",
        },
    )
    relative_height: None | int = field(
        default=None,
        metadata={
            "name": "relativeHeight",
            "type": "Attribute",
        },
    )
    behind_doc: None | bool = field(
        default=None,
        metadata={
            "name": "behindDoc",
            "type": "Attribute",
        },
    )
    locked: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    layout_in_cell: None | bool = field(
        default=None,
        metadata={
            "name": "layoutInCell",
            "type": "Attribute",
        },
    )
    hidden: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    allow_overlap: None | bool = field(
        default=None,
        metadata={
            "name": "allowOverlap",
            "type": "Attribute",
        },
    )
    anchor_id: None | bytes = field(
        default=None,
        metadata={
            "name": "anchorId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "length": 4,
            "format": "base16",
        },
    )
    edit_id: None | bytes = field(
        default=None,
        metadata={
            "name": "editId",
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "length": 4,
            "format": "base16",
        },
    )


@dataclass(slots=True, kw_only=True)
class Anchor(Anchor1):
    class Meta:
        name = "anchor"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTNonVisualDrawingProps,
    CTNonVisualGraphicFrameProperties,
    CTPoint2D,
    CTPositiveSize2D,
    Graphic,
)
from docx4j_py.word.wp14 import (
    SizeRelH,
    SizeRelV,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.wordprocessing_drawing import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.dml.wordprocessing_drawing.el",
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
