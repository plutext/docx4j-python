from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from docx4j_py.child import Child

__NAMESPACE__ = "urn:schemas-microsoft-com:office:word"


@dataclass(slots=True, kw_only=True)
class CTAnchorLock(Child):
    class Meta:
        name = "CT_AnchorLock"


class STBorderShadow(Enum):
    T = "t"
    TRUE = "true"
    F = "f"
    FALSE = "false"


class STBorderType(Enum):
    NONE = "none"
    SINGLE = "single"
    THICK = "thick"
    DOUBLE = "double"
    HAIRLINE = "hairline"
    DOT = "dot"
    DASH = "dash"
    DOT_DASH = "dotDash"
    DASH_DOT_DOT = "dashDotDot"
    TRIPLE = "triple"
    THIN_THICK_SMALL = "thinThickSmall"
    THICK_THIN_SMALL = "thickThinSmall"
    THICK_BETWEEN_THIN_SMALL = "thickBetweenThinSmall"
    THIN_THICK = "thinThick"
    THICK_THIN = "thickThin"
    THICK_BETWEEN_THIN = "thickBetweenThin"
    THIN_THICK_LARGE = "thinThickLarge"
    THICK_THIN_LARGE = "thickThinLarge"
    THICK_BETWEEN_THIN_LARGE = "thickBetweenThinLarge"
    WAVE = "wave"
    DOUBLE_WAVE = "doubleWave"
    DASHED_SMALL = "dashedSmall"
    DASH_DOT_STROKED = "dashDotStroked"
    THREE_DEMBOSS = "threeDEmboss"
    THREE_DENGRAVE = "threeDEngrave"
    HTMLOUTSET = "HTMLOutset"
    HTMLINSET = "HTMLInset"


class STHorizontalAnchor(Enum):
    MARGIN = "margin"
    PAGE = "page"
    TEXT = "text"
    CHAR = "char"


class STVerticalAnchor(Enum):
    MARGIN = "margin"
    PAGE = "page"
    TEXT = "text"
    LINE = "line"


class STWrapSide(Enum):
    BOTH = "both"
    LEFT = "left"
    RIGHT = "right"
    LARGEST = "largest"


class STWrapType(Enum):
    TOP_AND_BOTTOM = "topAndBottom"
    SQUARE = "square"
    NONE = "none"
    TIGHT = "tight"
    THROUGH = "through"


@dataclass(slots=True, kw_only=True)
class CTBorder(Child):
    class Meta:
        name = "CT_Border"

    type_value: None | STBorderType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    width: None | int = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    shadow: None | STBorderShadow = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class CTWrap(Child):
    class Meta:
        name = "CT_Wrap"

    type_value: None | STWrapType = field(
        default=None,
        metadata={
            "name": "type",
            "type": "Attribute",
        },
    )
    side: None | STWrapSide = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    anchorx: None | STHorizontalAnchor = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )
    anchory: None | STVerticalAnchor = field(
        default=None,
        metadata={
            "type": "Attribute",
        },
    )


@dataclass(slots=True, kw_only=True)
class Anchorlock(CTAnchorLock):
    class Meta:
        name = "anchorlock"
        namespace = "urn:schemas-microsoft-com:office:word"


@dataclass(slots=True, kw_only=True)
class Borderbottom(CTBorder):
    class Meta:
        name = "borderbottom"
        namespace = "urn:schemas-microsoft-com:office:word"


@dataclass(slots=True, kw_only=True)
class Borderleft(CTBorder):
    class Meta:
        name = "borderleft"
        namespace = "urn:schemas-microsoft-com:office:word"


@dataclass(slots=True, kw_only=True)
class Borderright(CTBorder):
    class Meta:
        name = "borderright"
        namespace = "urn:schemas-microsoft-com:office:word"


@dataclass(slots=True, kw_only=True)
class Bordertop(CTBorder):
    class Meta:
        name = "bordertop"
        namespace = "urn:schemas-microsoft-com:office:word"


@dataclass(slots=True, kw_only=True)
class Wrap(CTWrap):
    class Meta:
        name = "wrap"
        namespace = "urn:schemas-microsoft-com:office:word"


# CR-001 Phase C: el, the fragment helpers and the text sugar, reached
# through this package as CR-001 section 6.2 writes them
# (``from docx4j_py.vml.wordprocessing_drawing import el, p, r, t, wml, to_xml, text_of, walk, find``).
# Lazily, because the hand-written modules import the classes above and an
# eager import here would be a cycle.
_PHASE_C: dict[str, str] = {
    "FragmentError": "docx4j_py.fragments",
    "deep_copy": "docx4j_py.child",
    "deep_copy_as": "docx4j_py.child",
    "el": "docx4j_py.vml.wordprocessing_drawing.el",
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
