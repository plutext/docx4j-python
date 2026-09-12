from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2014/main"


@dataclass(slots=True, kw_only=True)
class CTConnectableReferences(Child):
    class Meta:
        name = "CT_ConnectableReferences"

    st: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    end: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCreationId(Child):
    class Meta:
        name = "CT_CreationId"

    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTIdentifier(Child):
    class Meta:
        name = "CT_Identifier"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPredecessorDrawingElementReference(Child):
    class Meta:
        name = "CT_PredecessorDrawingElementReference"

    pred: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class ColId(CTIdentifier):
    class Meta:
        name = "colId"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/main"


@dataclass(slots=True, kw_only=True)
class CreationId(CTCreationId):
    class Meta:
        name = "creationId"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/main"


@dataclass(slots=True, kw_only=True)
class CxnDerefs(CTConnectableReferences):
    class Meta:
        name = "cxnDERefs"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/main"


@dataclass(slots=True, kw_only=True)
class PredDeref(CTPredecessorDrawingElementReference):
    class Meta:
        name = "predDERef"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/main"


@dataclass(slots=True, kw_only=True)
class RowId(CTIdentifier):
    class Meta:
        name = "rowId"
        namespace = "http://schemas.microsoft.com/office/drawing/2014/main"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.main_2014 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.main_2014.el",
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
