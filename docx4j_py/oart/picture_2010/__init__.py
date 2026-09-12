from __future__ import annotations

from dataclasses import dataclass

from docx4j_py.dml.main import (
    CTOfficeArtExtensionList,
    CTShapeStyle,
)

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2010/picture"


@dataclass(slots=True, kw_only=True)
class ExtLst(CTOfficeArtExtensionList):
    class Meta:
        name = "extLst"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/picture"


@dataclass(slots=True, kw_only=True)
class Style(CTShapeStyle):
    class Meta:
        name = "style"
        namespace = "http://schemas.microsoft.com/office/drawing/2010/picture"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.picture_2010 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.picture_2010.el",
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
