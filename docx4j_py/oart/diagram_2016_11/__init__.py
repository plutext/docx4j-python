from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2016/11/diagram"


class STSTorageType(Enum):
    SIB_TRANS = "sibTrans"
    PAR_TRANS = "parTrans"


@dataclass(slots=True, kw_only=True)
class CTDiagramAutoBullet(Child):
    class Meta:
        name = "CT_DiagramAutoBullet"

    content: (
        None | CTTextNoBullet | CTTextAutonumberBullet | CTTextCharBullet | CTTextBlipBullet
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "buNone",
                    "type": ForwardRef("CTTextNoBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buAutoNum",
                    "type": ForwardRef("CTTextAutonumberBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buChar",
                    "type": ForwardRef("CTTextCharBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
                {
                    "name": "buBlip",
                    "type": ForwardRef("CTTextBlipBullet"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/main",
                },
            ),
        },
    )
    prefix: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    lead_zeros: None | bool = field(
        default=None,
        metadata={
            "name": "leadZeros",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumberDiagramInfo(Child):
    class Meta:
        name = "CT_NumberDiagramInfo"

    bu_pr: None | CTDiagramAutoBullet = field(
        default=None,
        metadata={
            "name": "buPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2016/11/diagram",
        },
    )
    lvl: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    pt_type: None | STSTorageType = field(
        default=None,
        metadata={
            "name": "ptType",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTNumberDiagramInfoList(Child):
    class Meta:
        name = "CT_NumberDiagramInfoList"

    auto_bu_node_info: list[CTNumberDiagramInfo] = field(
        default_factory=ChildList,
        metadata={
            "name": "autoBuNodeInfo",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2016/11/diagram",
        },
    )


@dataclass(slots=True, kw_only=True)
class AutoBuNodeInfoLst(CTNumberDiagramInfoList):
    class Meta:
        name = "autoBuNodeInfoLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2016/11/diagram"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTTextAutonumberBullet,
    CTTextBlipBullet,
    CTTextCharBullet,
    CTTextNoBullet,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.diagram_2016_11 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.oart.diagram_2016_11.el",
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
