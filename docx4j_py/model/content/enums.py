"""Office JS's values, as ``Literal`` types and as ``StrEnum``s beside them.

CR-003 section 3.1 and decided question 1: the *members* are ``snake_case``, the
*values* stay Office JS's strings, because they are data that cross into tool
arguments and JSON results. Each is a :class:`typing.Literal` in the signature
and a :class:`enum.StrEnum` for those who prefer a name, so both spellings work
and neither is second class::

    body.insert_paragraph("Hello", location="Start")
    body.insert_paragraph("Hello", location=InsertLocation.START)

``InsertLocation.START == "Start"`` is True, so a ``StrEnum`` member is accepted
anywhere the ``Literal`` is and serialises as the string.
"""

from __future__ import annotations

from enum import StrEnum
from typing import Literal

__all__ = [
    "ALIGNMENTS",
    "BREAK_TYPES",
    "CHANGE_TRACKING_MODES",
    "TRACKED_CHANGE_TYPES",
    "UNDERLINE_TYPES",
    "Alignment",
    "AlignmentValue",
    "BodyLocation",
    "BreakType",
    "BreakTypeValue",
    "ChangeTracking",
    "ChangeTrackingModeValue",
    "InsertLocation",
    "Location",
    "ParagraphLocation",
    "RangeLocation",
    "TextLocation",
    "TextView",
    "TrackedChangeType",
    "TrackedChangeTypeValue",
    "UnderlineType",
    "UnderlineValue",
]

# ---------------------------------------------------------------------------
# insert locations
# ---------------------------------------------------------------------------

#: Office JS ``Word.InsertLocation``, every value.
Location = Literal["Start", "End", "Before", "After", "Replace"]
#: Where a body takes a block: its start or its end.
BodyLocation = Literal["Start", "End"]
#: Where text goes in a paragraph or a body.
TextLocation = Literal["Start", "End", "Replace"]
#: Where a paragraph goes relative to another one.
ParagraphLocation = Literal["Before", "After"]
#: Office JS ``Word.RangeLocation``, what ``get_range`` takes.
RangeLocation = Literal["Whole", "Start", "End", "Content"]
#: The two views of a document with tracked changes (CR-003 section 3.8).
TextView = Literal["accepted", "original"]


class InsertLocation(StrEnum):
    """Office JS ``Word.InsertLocation``."""

    START = "Start"
    END = "End"
    BEFORE = "Before"
    AFTER = "After"
    REPLACE = "Replace"


# ---------------------------------------------------------------------------
# paragraph alignment
# ---------------------------------------------------------------------------

#: Office JS ``Word.Alignment``, the subset a paragraph reports. ``"Unknown"``
#: is what a paragraph with no ``w:jc`` reads (CR-003 section 4).
AlignmentValue = Literal["Unknown", "Left", "Centered", "Right", "Justified"]


class Alignment(StrEnum):
    """Office JS ``Word.Alignment``."""

    UNKNOWN = "Unknown"
    LEFT = "Left"
    CENTERED = "Centered"
    RIGHT = "Right"
    JUSTIFIED = "Justified"


#: Every alignment value, in Office JS's order.
ALIGNMENTS: tuple[str, ...] = tuple(a.value for a in Alignment)


# ---------------------------------------------------------------------------
# underline
# ---------------------------------------------------------------------------

#: Office JS ``Word.UnderlineType``. ``"None"`` removes ``w:u``; ``"Mixed"`` is
#: read for a ``w:u`` value Office JS has no name for.
UnderlineValue = Literal[
    "None",
    "Single",
    "Word",
    "Double",
    "Thick",
    "Dotted",
    "DottedHeavy",
    "DashLine",
    "DashLineHeavy",
    "DashLineLong",
    "DashLineLongHeavy",
    "DotDashLine",
    "DotDashLineHeavy",
    "TwoDotDashLine",
    "TwoDotDashLineHeavy",
    "Wave",
    "WaveHeavy",
    "WaveDouble",
    "Mixed",
]


class UnderlineType(StrEnum):
    """Office JS ``Word.UnderlineType``; the values :data:`UNDERLINE` maps."""

    NONE = "None"
    SINGLE = "Single"
    WORD = "Word"
    DOUBLE = "Double"
    THICK = "Thick"
    DOTTED = "Dotted"
    DOTTED_HEAVY = "DottedHeavy"
    DASH_LINE = "DashLine"
    DASH_LINE_HEAVY = "DashLineHeavy"
    DASH_LINE_LONG = "DashLineLong"
    DASH_LINE_LONG_HEAVY = "DashLineLongHeavy"
    DOT_DASH_LINE = "DotDashLine"
    DOT_DASH_LINE_HEAVY = "DotDashLineHeavy"
    TWO_DOT_DASH_LINE = "TwoDotDashLine"
    TWO_DOT_DASH_LINE_HEAVY = "TwoDotDashLineHeavy"
    WAVE = "Wave"
    WAVE_HEAVY = "WaveHeavy"
    WAVE_DOUBLE = "WaveDouble"
    MIXED = "Mixed"


#: Every underline value, in Office JS's order.
UNDERLINE_TYPES: tuple[str, ...] = tuple(u.value for u in UnderlineType)


# ---------------------------------------------------------------------------
# breaks
# ---------------------------------------------------------------------------

#: Office JS ``Word.BreakType``, the subset a ``w:br`` and a ``w:sectPr`` can
#: carry. The section breaks are Phase C's; ``"Page"`` and ``"Line"`` are here.
BreakTypeValue = Literal["Page", "Line", "Next", "SectionNext", "SectionContinuous"]


class BreakType(StrEnum):
    """Office JS ``Word.BreakType``."""

    PAGE = "Page"
    LINE = "Line"
    NEXT = "Next"
    SECTION_NEXT = "SectionNext"
    SECTION_CONTINUOUS = "SectionContinuous"


#: Every break type, in Office JS's order.
BREAK_TYPES: tuple[str, ...] = tuple(b.value for b in BreakType)


# ---------------------------------------------------------------------------
# change tracking (CR-003 section 3.8)
# ---------------------------------------------------------------------------

#: Office JS ``Word.ChangeTrackingMode``. ``"TrackMineOnly"`` is **stored as**
#: ``"TrackAll"``: ``w:trackRevisions`` is a flag and a file cannot tell the two
#: apart (CR-003 section 4).
ChangeTrackingModeValue = Literal["Off", "TrackAll", "TrackMineOnly"]


class ChangeTracking(StrEnum):
    """Office JS ``Word.ChangeTrackingMode``."""

    OFF = "Off"
    TRACK_ALL = "TrackAll"
    TRACK_MINE_ONLY = "TrackMineOnly"


#: Every mode, in Office JS's order.
CHANGE_TRACKING_MODES: tuple[str, ...] = tuple(m.value for m in ChangeTracking)


#: Office JS ``Word.TrackedChangeType``. ``"None"`` is in Office JS's list and is
#: never produced here: every piece of markup a ``TrackedChange`` is over is an
#: insertion, a deletion or a formatting change.
TrackedChangeTypeValue = Literal["Added", "Deleted", "Formatted", "None"]


class TrackedChangeType(StrEnum):
    """Office JS ``Word.TrackedChangeType``."""

    ADDED = "Added"
    DELETED = "Deleted"
    FORMATTED = "Formatted"
    NONE = "None"


#: Every tracked-change type, in Office JS's order.
TRACKED_CHANGE_TYPES: tuple[str, ...] = tuple(t.value for t in TrackedChangeType)
