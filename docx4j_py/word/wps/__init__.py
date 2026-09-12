from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


@dataclass(slots=True, kw_only=True)
class CTLinkedTextboxInformation(Child):
    class Meta:
        name = "CT_LinkedTextboxInformation"

    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    seq: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTTextboxInfo(Child):
    class Meta:
        name = "CT_TextboxInfo"

    txbx_content: None | TxbxContent = field(
        default=None,
        metadata={
            "name": "txbxContent",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    id: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "0",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWordprocessingShape(Child):
    class Meta:
        name = "CT_WordprocessingShape"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    c_nv_sp_pr_or_c_nv_cn_pr: (
        None | CTNonVisualDrawingShapeProps | CTNonVisualConnectorProperties
    ) = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "cNvSpPr",
                    "type": ForwardRef("CTNonVisualDrawingShapeProps"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                },
                {
                    "name": "cNvCnPr",
                    "type": ForwardRef("CTNonVisualConnectorProperties"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                },
            ),
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    style: None | CTShapeStyle = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    txbx_or_linked_txbx: None | CTTextboxInfo | CTLinkedTextboxInformation = field(
        default=None,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "txbx",
                    "type": ForwardRef("CTTextboxInfo"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                },
                {
                    "name": "linkedTxbx",
                    "type": ForwardRef("CTLinkedTextboxInformation"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                },
            ),
        },
    )
    body_pr: None | CTTextBodyProperties = field(
        default=None,
        metadata={
            "name": "bodyPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        },
    )
    normal_east_asian_flow: None | bool = field(
        default=None,
        metadata={
            "name": "normalEastAsianFlow",
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class Wsp(CTWordprocessingShape):
    class Meta:
        name = "wsp"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTNonVisualConnectorProperties,
    CTNonVisualDrawingProps,
    CTNonVisualDrawingShapeProps,
    CTOfficeArtExtensionList,
    CTShapeProperties,
    CTShapeStyle,
    CTTextBodyProperties,
)
from docx4j_py.wml import TxbxContent


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.word.wps import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.word.wps.el",
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
