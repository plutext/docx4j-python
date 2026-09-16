from __future__ import annotations

from dataclasses import dataclass, field
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"


@dataclass(slots=True, kw_only=True)
class CTWordprocessingCanvas(Child):
    class Meta:
        name = "CT_WordprocessingCanvas"

    bg: None | CTBackgroundFormatting = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        },
    )
    whole: None | CTWholeE2OFormatting = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        },
    )
    content: list[Wsp | Pic | ContentPart | Wgp | CTGraphicFrame] = field(
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
                    "name": "pic",
                    "type": ForwardRef("Pic"),
                    "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
                },
                {
                    "name": "contentPart",
                    "type": ForwardRef("ContentPart"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordml",
                },
                {
                    "name": "wgp",
                    "type": ForwardRef("Wgp"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                },
                {
                    "name": "graphicFrame",
                    "type": ForwardRef("CTGraphicFrame"),
                    "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                },
            ),
        },
    )
    ext_lst: None | CTOfficeArtExtensionList = field(
        default=None,
        metadata={
            "name": "extLst",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        },
    )


@dataclass(slots=True, kw_only=True)
class Wpc(CTWordprocessingCanvas):
    class Meta:
        name = "wpc"
        namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTBackgroundFormatting,
    CTOfficeArtExtensionList,
    CTWholeE2OFormatting,
)
from docx4j_py.dml.picture import Pic
from docx4j_py.w14 import ContentPart
from docx4j_py.word.wpg import (
    CTGraphicFrame,
    Wgp,
)
from docx4j_py.word.wps import Wsp


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.word.wpc import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.word.wpc.el",
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
