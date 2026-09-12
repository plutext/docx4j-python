from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child
from docx4j_py.dml.main import STBlackWhiteMode

__NAMESPACE__ = "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing"


@dataclass(slots=True, kw_only=True)
class CTApplicationNonVisualDrawingProps(Child):
    class Meta:
        name = "CT_ApplicationNonVisualDrawingProps"

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
class CTContentPartNonVisual(Child):
    class Meta:
        name = "CT_ContentPartNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        },
    )
    c_nv_content_part_pr: None | CTNonVisualInkContentPartProperties = field(
        default=None,
        metadata={
            "name": "cNvContentPartPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTContentPart(Child):
    class Meta:
        name = "CT_ContentPart"

    nv_content_part_pr: None | CTContentPartNonVisual = field(
        default=None,
        metadata={
            "name": "nvContentPartPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        },
    )
    nv_pr: None | CTApplicationNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "nvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        },
    )
    xfrm: None | CTTransform2D = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        },
    )
    bw_mode: None | STBlackWhiteMode = field(
        default=None,
        metadata={
            "name": "bwMode",
            "type": "Attribute",
            "schema_default": STBlackWhiteMode.AUTO,
        },
    )


@dataclass(slots=True, kw_only=True)
class ContentPart(CTContentPart):
    class Meta:
        name = "contentPart"
        namespace = "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTNonVisualDrawingProps,
    CTOfficeArtExtensionList,
    CTTransform2D,
)
from docx4j_py.oart.main_2010 import CTNonVisualInkContentPartProperties


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.spreadsheet_drawing_2010 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.spreadsheet_drawing_2010.el",
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
