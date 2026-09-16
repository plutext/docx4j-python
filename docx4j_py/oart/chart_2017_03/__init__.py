from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2017/03/chart"


@dataclass(slots=True, kw_only=True)
class CTBooleanFalse(Child):
    class Meta:
        name = "CT_BooleanFalse"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/03/chart",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTDataDisplayOptions16(Child):
    class Meta:
        name = "CT_DataDisplayOptions16"

    disp_na_as_blank: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "dispNaAsBlank",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2017/03/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class DataDisplayOptions16(CTDataDisplayOptions16):
    class Meta:
        name = "dataDisplayOptions16"
        namespace = "http://schemas.microsoft.com/office/drawing/2017/03/chart"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.chart_2017_03 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.oart.chart_2017_03.el",
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
