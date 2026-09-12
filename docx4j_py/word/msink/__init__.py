from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_py.child import Child, ChildList

__NAMESPACE__ = "http://schemas.microsoft.com/ink/2010/main"


@dataclass(slots=True, kw_only=True)
class CTProperty(Child):
    class Meta:
        name = "CT_Property"

    value: bytes = field(
        default=b"",
        metadata={
            "format": "base16",
        },
    )
    type_value: None | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


class STDir(Enum):
    TO = "to"
    FROM = "from"
    WITH = "with"


class STKnownCtxNodeType(Enum):
    ROOT = "root"
    UNCLASSIFIED_INK = "unclassifiedInk"
    WRITING_REGION = "writingRegion"
    ANALYSIS_HINT = "analysisHint"
    OBJECT = "object"
    INK_DRAWING = "inkDrawing"
    IMAGE = "image"
    PARAGRAPH = "paragraph"
    LINE = "line"
    INK_BULLET = "inkBullet"
    INK_WORD = "inkWord"
    TEXT_WORD = "textWord"
    CUSTOM_RECOGNIZER = "customRecognizer"
    MATH_REGION = "mathRegion"
    MATH_EQUATION = "mathEquation"
    MATH_STRUCT = "mathStruct"
    MATH_SYMBOL = "mathSymbol"
    MATH_IDENTIFIER = "mathIdentifier"
    MATH_OPERATOR = "mathOperator"
    MATH_NUMBER = "mathNumber"
    NON_INK_DRAWING = "nonInkDrawing"
    GROUP_NODE = "groupNode"
    MIXED_DRAWING = "mixedDrawing"


class STKnownSemanticType(Enum):
    NONE = "none"
    UNDERLINE = "underline"
    STRIKETHROUGH = "strikethrough"
    HIGHLIGHT = "highlight"
    SCRATCH_OUT = "scratchOut"
    VERTICAL_RANGE = "verticalRange"
    CALLOUT = "callout"
    ENCLOSURE = "enclosure"
    COMMENT = "comment"
    CONTAINER = "container"
    CONNECTOR = "connector"


@dataclass(slots=True, kw_only=True)
class CTCtxLink(Child):
    class Meta:
        name = "CT_CtxLink"

    direction: None | STDir = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    ref: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTCtxNode(Child):
    class Meta:
        name = "CT_CtxNode"

    property: list[CTProperty] = field(
        default_factory=ChildList,
        metadata={
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/ink/2010/main",
        },
    )
    source_link: list[CTCtxLink] = field(
        default_factory=ChildList,
        metadata={
            "name": "sourceLink",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/ink/2010/main",
        },
    )
    destination_link: list[CTCtxLink] = field(
        default_factory=ChildList,
        metadata={
            "name": "destinationLink",
            "type": "Element",
            "namespace": "http://schemas.microsoft.com/ink/2010/main",
        },
    )
    id: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    type_value: None | STKnownCtxNodeType | str = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    rotated_bounding_box: list[str] = field(
        default_factory=list,
        metadata={
            "name": "rotatedBoundingBox",
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )
    alignment_level: None | int = field(
        default=None,
        metadata={
            "name": "alignmentLevel",
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    content_type: None | int = field(
        default=None,
        metadata={
            "name": "contentType",
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    ascender: list[str] = field(
        default_factory=lambda: [
            "0,0",
        ],
        metadata={
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )
    descender: list[str] = field(
        default_factory=lambda: [
            "0,0",
        ],
        metadata={
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )
    baseline: list[str] = field(
        default_factory=lambda: [
            "0,0",
        ],
        metadata={
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )
    midline: list[str] = field(
        default_factory=lambda: [
            "0,0",
        ],
        metadata={
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )
    custom_recognizer_id: None | str = field(
        default=None,
        metadata={
            "name": "customRecognizerId",
            "type": "Attribute",
            "pattern": r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
        },
    )
    math_ml: None | str = field(
        default=None,
        metadata={
            "name": "mathML",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    math_struct: None | str = field(
        default=None,
        metadata={
            "name": "mathStruct",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    math_symbol: None | str = field(
        default=None,
        metadata={
            "name": "mathSymbol",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    begin_modifier_type: None | str = field(
        default=None,
        metadata={
            "name": "beginModifierType",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    end_modifier_type: None | str = field(
        default=None,
        metadata={
            "name": "endModifierType",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    rotation_angle: None | int = field(
        default=None,
        metadata={
            "name": "rotationAngle",
            "type": "Attribute",
            "schema_default": "0",
        },
    )
    hot_points: list[str] = field(
        default_factory=list,
        metadata={
            "name": "hotPoints",
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )
    centroid: None | str = field(
        default=None,
        metadata={
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
        },
    )
    semantic_type: None | STKnownSemanticType | int = field(
        default=None,
        metadata={
            "name": "semanticType",
            "type": "Attribute",
            "schema_default": STKnownSemanticType.NONE,
        },
    )
    shape_name: None | str = field(
        default=None,
        metadata={
            "name": "shapeName",
            "type": "Attribute",
            "schema_default": "",
        },
    )
    shape_geometry: list[str] = field(
        default_factory=list,
        metadata={
            "name": "shapeGeometry",
            "type": "Attribute",
            "pattern": r"-?[0-9]+,-?[0-9]+",
            "tokens": True,
        },
    )


@dataclass(slots=True, kw_only=True)
class Context(CTCtxNode):
    class Meta:
        name = "context"
        namespace = "http://schemas.microsoft.com/ink/2010/main"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.word.msink import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "el": "docx4j_py.word.msink.el",
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
