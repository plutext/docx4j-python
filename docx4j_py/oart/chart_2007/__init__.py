from __future__ import annotations

from dataclasses import dataclass, field

from docx4j_py.child import Child

__NAMESPACE__ = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"


@dataclass(slots=True, kw_only=True)
class CTBooleanFalse(Child):
    class Meta:
        name = "CT_BooleanFalse"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "false",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTBooleanTrue(Child):
    class Meta:
        name = "CT_BooleanTrue"

    val: None | bool = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "true",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTStyle(Child):
    class Meta:
        name = "CT_Style"

    val: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
            "min_inclusive": 101,
            "max_inclusive": 148,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTInvertSolidFillFmt(Child):
    class Meta:
        name = "CT_InvertSolidFillFmt"

    sp_pr: None | CTShapeProperties = field(
        default=None,
        metadata={
            "name": "spPr",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTPivotOptions(Child):
    class Meta:
        name = "CT_PivotOptions"

    drop_zone_filter: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "dropZoneFilter",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        },
    )
    drop_zone_categories: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "dropZoneCategories",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        },
    )
    drop_zone_data: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "dropZoneData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        },
    )
    drop_zone_series: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "dropZoneSeries",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        },
    )
    drop_zones_visible: None | CTBooleanFalse = field(
        default=None,
        metadata={
            "name": "dropZonesVisible",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart",
        },
    )


@dataclass(slots=True, kw_only=True)
class Style(CTStyle):
    class Meta:
        name = "style"
        namespace = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"


@dataclass(slots=True, kw_only=True)
class InvertSolidFillFmt(CTInvertSolidFillFmt):
    class Meta:
        name = "invertSolidFillFmt"
        namespace = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"


@dataclass(slots=True, kw_only=True)
class PivotOptions(CTPivotOptions):
    class Meta:
        name = "pivotOptions"
        namespace = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.dml.main import CTShapeProperties


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.chart_2007 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.chart_2007.el",
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
