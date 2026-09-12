from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"


class STSizeRelFromH(Enum):
    MARGIN = "margin"
    PAGE = "page"
    LEFT_MARGIN = "leftMargin"
    RIGHT_MARGIN = "rightMargin"
    INSIDE_MARGIN = "insideMargin"
    OUTSIDE_MARGIN = "outsideMargin"


class STSizeRelFromV(Enum):
    MARGIN = "margin"
    PAGE = "page"
    TOP_MARGIN = "topMargin"
    BOTTOM_MARGIN = "bottomMargin"
    INSIDE_MARGIN = "insideMargin"
    OUTSIDE_MARGIN = "outsideMargin"


@dataclass(slots=True, kw_only=True)
class PctPosHoffset(Child):
    class Meta:
        name = "pctPosHOffset"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class PctPosVoffset(Child):
    class Meta:
        name = "pctPosVOffset"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"

    value: None | int = field(default=None)


@dataclass(slots=True, kw_only=True)
class CTSizeRelH(Child):
    class Meta:
        name = "CT_SizeRelH"

    pct_width: None | int = field(
        default=None,
        metadata={
            "name": "pctWidth",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "min_inclusive": 0,
        },
    )
    relative_from: None | STSizeRelFromH = field(
        default=None,
        metadata={
            "name": "relativeFrom",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTSizeRelV(Child):
    class Meta:
        name = "CT_SizeRelV"

    pct_height: None | int = field(
        default=None,
        metadata={
            "name": "pctHeight",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            "min_inclusive": 0,
        },
    )
    relative_from: None | STSizeRelFromV = field(
        default=None,
        metadata={
            "name": "relativeFrom",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class SizeRelH(CTSizeRelH):
    class Meta:
        name = "sizeRelH"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"


@dataclass(slots=True, kw_only=True)
class SizeRelV(CTSizeRelV):
    class Meta:
        name = "sizeRelV"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.word.wp14 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.word.wp14.el",
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
