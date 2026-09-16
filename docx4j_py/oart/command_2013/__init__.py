from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_xsdata.models.datatype import XmlDateTime

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class CTConnectorMoniker(Child):
    class Meta:
        name = "CT_ConnectorMoniker"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    creation_id: None | str = field(
        default=None,
        metadata={
            "name": "creationId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTConnectorMonikerList(Child):
    class Meta:
        name = "CT_ConnectorMonikerList"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicFrameMoniker(Child):
    class Meta:
        name = "CT_GraphicFrameMoniker"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    creation_id: None | str = field(
        default=None,
        metadata={
            "name": "creationId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicFrameMonikerList(Child):
    class Meta:
        name = "CT_GraphicFrameMonikerList"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupShapeMoniker(Child):
    class Meta:
        name = "CT_GroupShapeMoniker"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    creation_id: None | str = field(
        default=None,
        metadata={
            "name": "creationId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupShapeMonikerList(Child):
    class Meta:
        name = "CT_GroupShapeMonikerList"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInkMoniker(Child):
    class Meta:
        name = "CT_InkMoniker"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    creation_id: None | str = field(
        default=None,
        metadata={
            "name": "creationId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInkMonikerList(Child):
    class Meta:
        name = "CT_InkMonikerList"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureMoniker(Child):
    class Meta:
        name = "CT_PictureMoniker"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    creation_id: None | str = field(
        default=None,
        metadata={
            "name": "creationId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureMonikerList(Child):
    class Meta:
        name = "CT_PictureMonikerList"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeMoniker(Child):
    class Meta:
        name = "CT_ShapeMoniker"

    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    creation_id: None | str = field(
        default=None,
        metadata={
            "name": "creationId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeMonikerList(Child):
    class Meta:
        name = "CT_ShapeMonikerList"

    any_element: list[object] = field(
        default_factory=ChildList,
        metadata={
            "type": "Wildcard",
            "namespace": "##any",
        },
    )


class STConnectorChangeBit(Enum):
    ADD = "add"
    DEL = "del"
    MOD = "mod"
    ORD = "ord"
    TOP_LVL = "topLvl"
    MOD_VIS = "modVis"
    REPL_ST = "replST"
    DEL_ST = "delST"
    REPL_ID = "replId"


class STGraphicFrameChangeBit(Enum):
    ADD = "add"
    DEL = "del"
    MOD = "mod"
    ORD = "ord"
    TOP_LVL = "topLvl"
    MOD_VIS = "modVis"
    REPL_ST = "replST"
    DEL_ST = "delST"
    REPL_ID = "replId"
    MOD_GRAPHIC = "modGraphic"


class STGroupShapeChangeBit(Enum):
    ADD = "add"
    DEL = "del"
    MOD = "mod"
    ORD = "ord"
    TOP_LVL = "topLvl"
    MOD_VIS = "modVis"
    REPL_ST = "replST"
    DEL_ST = "delST"
    REPL_ID = "replId"


class STInkChangeBit(Enum):
    ADD = "add"
    DEL = "del"
    MOD = "mod"
    ORD = "ord"
    TOP_LVL = "topLvl"
    MOD_VIS = "modVis"
    REPL_ST = "replST"
    DEL_ST = "delST"
    REPL_ID = "replId"
    RECO = "reco"
    MOD_STROKES = "modStrokes"


class STPictureChangeBit(Enum):
    ADD = "add"
    DEL = "del"
    MOD = "mod"
    ORD = "ord"
    TOP_LVL = "topLvl"
    MOD_VIS = "modVis"
    REPL_ST = "replST"
    DEL_ST = "delST"
    REPL_ID = "replId"
    MOD_CROP = "modCrop"


class STShapeChangeBit(Enum):
    ADD = "add"
    DEL = "del"
    MOD = "mod"
    ORD = "ord"
    TOP_LVL = "topLvl"
    MOD_VIS = "modVis"
    REPL_ST = "replST"
    DEL_ST = "delST"
    REPL_ID = "replId"
    MOD_CROP = "modCrop"


@dataclass(slots=True, kw_only=True)
class CTChangesData(Child):
    class Meta:
        name = "CT_ChangesData"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    name: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    user_id: None | str = field(
        default=None,
        metadata={
            "name": "userId",
            "type": "Attribute",
        },
    )
    provider_id: None | str = field(
        default=None,
        metadata={
            "name": "providerId",
            "type": "Attribute",
        },
    )
    cl_id: None | str = field(
        default=None,
        metadata={
            "name": "clId",
            "type": "Attribute",
        },
    )
    email: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    dt: None | XmlDateTime = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    v: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    act_id: None | int = field(
        default=None,
        metadata={
            "name": "actId",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CxnSpMkLst(CTConnectorMonikerList):
    class Meta:
        name = "cxnSpMkLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class GraphicFrameMkLst(CTGraphicFrameMonikerList):
    class Meta:
        name = "graphicFrameMkLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class GrpMkLst(CTGroupShapeMonikerList):
    class Meta:
        name = "grpMkLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class InkMkLst(CTInkMonikerList):
    class Meta:
        name = "inkMkLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class PicMkLst(CTPictureMonikerList):
    class Meta:
        name = "picMkLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class SpMkLst(CTShapeMonikerList):
    class Meta:
        name = "spMkLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2013/main/command"


@dataclass(slots=True, kw_only=True)
class CTConnectorChanges(Child):
    class Meta:
        name = "CT_ConnectorChanges"

    chg_data: None | CTChangesData = field(
        default=None,
        metadata={
            "name": "chgData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    cxn_sp_mk_lst: None | CTConnectorMonikerList = field(
        default=None,
        metadata={
            "name": "cxnSpMkLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    chg: list[STConnectorChangeBit] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGraphicFrameChanges(Child):
    class Meta:
        name = "CT_GraphicFrameChanges"

    chg_data: None | CTChangesData = field(
        default=None,
        metadata={
            "name": "chgData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    graphic_frame_mk_lst: None | CTGraphicFrameMonikerList = field(
        default=None,
        metadata={
            "name": "graphicFrameMkLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    chg: list[STGraphicFrameChangeBit] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTGroupShapeChanges(Child):
    class Meta:
        name = "CT_GroupShapeChanges"

    chg_data: None | CTChangesData = field(
        default=None,
        metadata={
            "name": "chgData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    grp_sp_mk_lst: None | CTGroupShapeMonikerList = field(
        default=None,
        metadata={
            "name": "grpSpMkLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    chg: list[STGroupShapeChangeBit] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInkChanges(Child):
    class Meta:
        name = "CT_InkChanges"

    chg_data: None | CTChangesData = field(
        default=None,
        metadata={
            "name": "chgData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ink_mk_lst: None | CTInkMonikerList = field(
        default=None,
        metadata={
            "name": "inkMkLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    chg: list[STInkChangeBit] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPictureChanges(Child):
    class Meta:
        name = "CT_PictureChanges"

    chg_data: None | CTChangesData = field(
        default=None,
        metadata={
            "name": "chgData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    pic_mk_lst: None | CTPictureMonikerList = field(
        default=None,
        metadata={
            "name": "picMkLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    chg: list[STPictureChangeBit] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTShapeChanges(Child):
    class Meta:
        name = "CT_ShapeChanges"

    chg_data: None | CTChangesData = field(
        default=None,
        metadata={
            "name": "chgData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    sp_mk_lst: None | CTShapeMonikerList = field(
        default=None,
        metadata={
            "name": "spMkLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2013/main/command",
        },
    )
    chg: list[STShapeChangeBit] = field(
        default_factory=list,
        metadata={
            "type": "Attribute",
            "tokens": True,
        },
    )


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import CTOfficeArtExtensionList


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.command_2013 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.oart.command_2013.el",
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
