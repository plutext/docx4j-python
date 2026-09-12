from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import ForwardRef

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/office/powerpoint/2014/inkAction"


class STActionTypeReserved(Enum):
    ADD = "add"
    REMOVE = "remove"
    TRANSFORM = "transform"


class STDataNameReserved(Enum):
    STROKE = "stroke"
    PATH = "path"
    TARGET = "target"


class STPropertyNameReserved(Enum):
    DATA_TYPE = "dataType"
    STYLE = "style"


class STPropertyValueReserved(Enum):
    INK = "ink"
    POINT_ERASER = "pointEraser"
    STROKE_ERASER = "strokeEraser"
    INSTANT = "instant"


@dataclass(slots=True, kw_only=True)
class CTActionData(Child):
    class Meta:
        name = "CT_ActionData"

    transform: None | CTMatrix = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
        },
    )
    trace_or_trace_view: list[Trace | TraceView] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "trace",
                    "type": ForwardRef("Trace"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
                {
                    "name": "traceView",
                    "type": ForwardRef("TraceView"),
                    "namespace": "http://www.w3.org/2003/InkML",
                },
            ),
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    name: None | STDataNameReserved | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STDataNameReserved.STROKE,
        },
    )
    ref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": "",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTActionProperty(Child):
    class Meta:
        name = "CT_ActionProperty"

    name: None | STPropertyNameReserved | str = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    value: None | STPropertyValueReserved | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STPropertyValueReserved.INK,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTActionDataGroup(Child):
    class Meta:
        name = "CT_ActionDataGroup"

    action_data: list[CTActionData] = field(
        default_factory=ChildList,
        metadata={
            "name": "actionData",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
            "min_occurs": 1,
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    name: None | STDataNameReserved | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "schema_default": STDataNameReserved.STROKE,
        },
    )


@dataclass(slots=True, kw_only=True)
class CTAction(Child):
    class Meta:
        name = "CT_Action"

    property: list[CTActionProperty] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
        },
    )
    action_data_or_action_data_group: list[CTActionData | CTActionDataGroup] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "actionData",
                    "type": ForwardRef("CTActionData"),
                    "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
                },
                {
                    "name": "actionDataGroup",
                    "type": ForwardRef("CTActionDataGroup"),
                    "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
                },
            ),
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    type_value: None | STActionTypeReserved | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    start_time: None | Decimal = field(
        default=None,
        metadata={
            "name": "startTime",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTActionGroup(Child):
    class Meta:
        name = "CT_ActionGroup"

    action: list[CTAction] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
            "min_occurs": 1,
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    type_value: None | STActionTypeReserved | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    start_time: None | Decimal = field(
        default=None,
        metadata={
            "name": "startTime",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTActions(Child):
    class Meta:
        name = "CT_Actions"

    definitions: None | Definitions = field(
        default=None,
        metadata={
            "type": "Element",
            "namespace": "http://www.w3.org/2003/InkML",
        },
    )
    action_group_or_action: list[CTActionGroup | CTAction] = field(
        default_factory=ChildList,
        metadata={
            "type": "Elements",
            "choices": (
                {
                    "name": "actionGroup",
                    "type": ForwardRef("CTActionGroup"),
                    "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
                },
                {
                    "name": "action",
                    "type": ForwardRef("CTAction"),
                    "namespace": "http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
                },
            ),
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "namespace": "http://www.w3.org/XML/1998/namespace",
        },
    )
    length_unit: None | StStandardLengthUnits = field(
        default=None,
        metadata={
            "name": "lengthUnit",
            "type": "Attribute",
        },
    )
    time_unit: None | StStandardTimeUnits = field(
        default=None,
        metadata={
            "name": "timeUnit",
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Actions(CTActions):
    class Meta:
        name = "actions"
        namespace = "http://schemas.microsoft.com/office/powerpoint/2014/inkAction"


# Imported below the classes, not above them: these names are only needed
# once the classes exist, and importing them earlier would not be possible
# where two modules depend on each other.
from docx4j_py.inkml import (
    CTMatrix,
    Definitions,
    StStandardLengthUnits,
    StStandardTimeUnits,
    Trace,
    TraceView,
)


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.oart.ink_action_2014 import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.oart.ink_action_2014.el",
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
