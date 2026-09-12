from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.openxmlformats.org/drawingml/2006/picture"


@dataclass(slots=True, kw_only=True)
class CTPictureNonVisual(Child):
    class Meta:
        name = "CT_PictureNonVisual"

    c_nv_pr: None | CTNonVisualDrawingProps = field(
        default=None,
        metadata={
            "name": "cNvPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        },
    )
    c_nv_pic_pr: None | CTNonVisualPictureProperties = field(
        default=None,
        metadata={
            "name": "cNvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        },
    )


@dataclass(slots=True, kw_only=True)
class Pic1(Child):
    class Meta:
        name = "CT_Picture"

    nv_pic_pr: None | CTPictureNonVisual = field(
        default=None,
        metadata={
            "name": "nvPicPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        },
    )
    blip_fill: None | CTBlipFillProperties = field(
        default=None,
        metadata={
            "name": "blipFill",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        },
    )
    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        },
    )


@dataclass(slots=True, kw_only=True)
class Pic(Pic1):
    class Meta:
        name = "pic"
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import (
    CTBlipFillProperties,
    CTNonVisualDrawingProps,
    CTNonVisualPictureProperties,
    CTShapeProperties,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.dml.picture import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.dml.picture.el",
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
