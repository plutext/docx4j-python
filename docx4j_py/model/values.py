"""Values as Word 365 reads them: the one reader of each kind of value (CR-018).

The parse keeps a value as it was written; a reader here applies Word's rule to
it. That is the split CR-018 section 4.1a settles: the fork keeps what was
written, exactly and explicitly, and ``docx4j_py`` decides what it means. Under
``runtime.parser_config()``'s exact parse an on/off value that is not exactly
``true``, ``false``, ``1`` or ``0`` stays the string it was (``"off"``,
``" false"``, ``""``, ``"TRUE"``), so every reader of an on/off value must read
it here rather than ask whether it ``is not False``.

**ST_OnOff** (CR-018 section 1.1, rule 1, four rounds of probes in Word 365,
2026-10-24). A value is **off for exactly** ``false``, ``off`` or ``0``:
case-sensitive and untrimmed. It is **on for anything else**: absent (the
schema's true for an element's ``w:val``), ``true``, ``on``, ``1``, ``""``,
``" "``, ``TRUE``, ``FALSE``, ``OFF``, ``" false"``, ``"false "``, ``00``,
``no``, ``yes``. Word opens all of them without a repair prompt.

* :func:`on_off` reads one value, :func:`is_on` an on/off element such as
  ``w:b`` (absent is off, present is :func:`on_off` of its ``w:val``);
* :func:`xsd_boolean` reads ``xsd:boolean`` as XSD reads it, whitespace
  collapsed, for the namespaces whose booleans are not Word's ST_OnOff
  (DrawingML's ``c:`` and ``cx:``, ``adec:``).

Measured on ``b``, ``i``, ``caps``, ``smallCaps``, ``strike``, ``dstrike``,
``outline``, ``shadow``, ``emboss``, ``imprint`` and ``vanish``; every other
on/off value (``bCs``, ``w:keepNext``, ``w:tblHeader``, ``w:trackRevisions``,
the on/off attributes such as ``w:fldLock``, ``w14`` and ``w15``) is read by the
same rule by extension, as docx4j's ``OnOffAdapter`` reads all of them
(docx4j CR-060 phase 1, ``893658f68``).

:func:`on_off_normal_form` is what a save writes for a kept value (section 4.7,
:mod:`docx4j_py.openpackaging.normalise`).

:func:`effective` (CR-018 Phase C, section 4.4) is the one reader of a value the
file leaves out: the stated value, else the schema's default, typed. It answers
**absence**; the readers above answer **a present value's spelling**, so an
on/off attribute is ``on_off(effective(obj, "fld_lock"))``.

**The enumerations** (CR-018 Phase B, section 4.2; rule 2). :func:`in_enumeration`
reads an enumerated ``w:val`` as Word does: a member only for exactly a
member's value, case and spacing and all, so anything else is unset, as if
absent. ``w:u``'s six on/off words read as ``none`` (round six,
:data:`UNDERLINE_ALIASES`). :func:`underline`, :func:`highlight`,
:func:`vertical_alignment` and :func:`emphasis_mark` are it for the four run
properties Word measured.

**Colour** (section 4.3; rule 3, rounds two to seven). :func:`colour` reads an
ST_HexColor value: ``"RRGGBB"`` for six hex digits in any case (trimmed, one
leading ``#`` allowed) or one of the 140 classic HTML names
(:data:`NETSCAPE_COLOURS`), :data:`AUTO` for exactly ``auto``, and None for
anything else. What None means is the attribute's: as if absent in a text or
underline colour (:func:`states_colour`), automatic in a fill or a border, no
page colour in ``w:background``.

**Integers** (CR-021 section 4.1, docx4j CR-065 phase 0: four rounds, 65 cases, Word
365, 2026-10-26). :func:`integer` reads a WordprocessingML integer attribute
(twips, half-points, eighths of a point, percentages, counts, ids) as Word does:
an optional sign, digits, and optionally a dot and digits, either digit group
possibly empty, or the empty string, reads as its **integer part with its
sign**, a decimal truncated toward zero (``"720.5"`` 720, ``"-121.5"`` -121,
``"141.99999999999994"`` 141, ``"240."`` 240, ``".5"`` 0, ``"+240.5"`` 240);
``""`` reads 0; any other form (``"abc"``, ``"1e3"``, ``"0x10"``, ``"240,5"``,
``" 240 "``) None, unset, because Word refuses the document holding it.

**Measurements as code sets them** (CR-022, docx4j CR-066 phase 0: rounds five and
six, Word 365, 2026-10-26). Word keeps each integer attribute in a field of its own
width, 16 bits signed or unsigned for most, and reads a value past it wrapped, with
no error and no repair. :func:`measure_in_range` is the one check every setter that
writes an integer measurement goes through: it converts the caller's value
(:data:`POINTS`, a python-docx ``Length`` in :data:`EMU`, :data:`LINES`, or the
attribute's own unit, :data:`OWN`) as the setter always has, and refuses it with a
``MeasureRangeError`` (``value.out_of_range``) where Word would not read it as
written, by the attribute's row of the limits table (:func:`integer_limit`,
``resources/integer-limits.json``), and, where Office JS's setter stops sooner, at
Office JS's limit (a :class:`CallerLimit`). The integer the setter writes is what it
returns. An attribute the probe did not measure has no row and is not checked (Q3).
"""

from __future__ import annotations

import dataclasses
import json
import math
import re
import threading
from collections.abc import Callable
from decimal import ROUND_CEILING, ROUND_FLOOR, Decimal
from enum import Enum
from typing import Any

__all__ = [
    "AUDITED_NAMESPACES",
    "AUTO",
    "EMU",
    "INTEGER_LIMITS_RESOURCE",
    "LINES",
    "NETSCAPE_COLOURS",
    "OFF",
    "OWN",
    "POINTS",
    "UNDERLINE_ALIASES",
    "XSD_FALSE",
    "XSD_TRUE",
    "CallerLimit",
    "IntegerLimit",
    "Units",
    "colour",
    "effective",
    "emphasis_mark",
    "highlight",
    "in_enumeration",
    "integer",
    "integer_limit",
    "integer_limits",
    "is_on",
    "measure_in_range",
    "on_off",
    "on_off_normal_form",
    "states_colour",
    "theme_colour",
    "underline",
    "unmeasured_attributes",
    "vertical_alignment",
    "xsd_boolean",
]

#: Rule 1: the only spellings Word reads as off. Exactly these, case-sensitive, untrimmed.
OFF: frozenset[str] = frozenset({"false", "off", "0"})

#: ``xsd:boolean``'s four lexical forms, after XSD's whitespace collapse.
XSD_TRUE: frozenset[str] = frozenset({"true", "1"})
XSD_FALSE: frozenset[str] = frozenset({"false", "0"})


def on_off(value: Any, default: bool = True) -> bool:
    """An ST_OnOff value as Word 365 reads it.

    ``None`` (the attribute absent) is `default`: the schema's true for an
    element's ``w:val`` (``BooleanDefaultTrue``), and the caller's answer for
    an attribute whose absence means something else (``w:fldLock``, absent, is
    off). A ``bool`` is itself, which is what the parse binds for exactly
    ``true``, ``false``, ``1`` and ``0``. An enumeration member reads by its
    value (``w:tblLook``'s ``on``/``off``, ``w14``'s and ``w15``'s ST_OnOff,
    OMML's ``m:val``). Any other value is off only if it is exactly
    ``"false"``, ``"off"`` or ``"0"``.
    """
    if value is None:
        return default
    if value is True or value is False:
        return value
    if isinstance(value, Enum):
        value = value.value
        if value is True or value is False:
            return value
    return str(value) not in OFF


def is_on(element: Any) -> bool:
    """An on/off element such as ``w:b`` or ``w:vanish``: absent is off; present, :func:`on_off` of its ``w:val``.

    A bare ``<w:b/>`` is on: ``w:val``'s schema default is true. A wildcard
    element (one the model holds untyped) reads its ``w:val`` attribute the
    same way.
    """
    if element is None:
        return False
    attributes = getattr(element, "attributes", None)
    if isinstance(attributes, dict) and not hasattr(element, "val"):
        value = next((v for k, v in attributes.items() if k == "val" or k.endswith("}val")), None)
        return on_off(value)
    return on_off(getattr(element, "val", None))


def on_off_normal_form(value: Any, absent_is_on: bool) -> bool | None:
    """What a save writes for an ST_OnOff value the parse kept as its text (CR-018 section 4.7's on/off rows).

    None means the attribute is dropped: Word reads the value as on and the
    attribute's absence is on too (an element's ``w:val``, whose schema default
    is true), so the value is written **bare**, as Word's own save writes it.
    Otherwise the ``bool`` Word reads, which the serialiser writes ``1`` or
    ``0``: off is ``0`` with the element kept, since an off value can override a
    style and a save looks nothing up.
    """
    on = on_off(value)
    if on and absent_is_on:
        return None
    return on


def xsd_boolean(value: Any, default: bool | None = None) -> bool | None:
    """An ``xsd:boolean`` value as XSD reads it: whitespace collapsed, then ``true``/``1`` or ``false``/``0``.

    For the namespaces whose booleans are XSD's, not Word's ST_OnOff:
    DrawingML's charts (``c:``, ``cx:``) and the decorative flag (``adec:``).
    ``None`` is `default`; a ``bool`` is itself; a value that is none of the
    four literals once collapsed is `default` too, as a schema-valid reader
    treats an invalid value as absent.
    """
    if value is None:
        return default
    if value is True or value is False:
        return value
    if isinstance(value, Enum):
        value = value.value
    text = " ".join(str(value).split())
    if text in XSD_TRUE:
        return True
    if text in XSD_FALSE:
        return False
    return default


# ---------------------------------------------------------------------------
# the enumerations (CR-018 Phase B, section 4.2)
# ---------------------------------------------------------------------------

#: ``w:u w:val``'s six on/off words, each read as ``none`` (round six, docx4j's
#: ``UnderlineValueAdapter``, ``1d1b824a9``): exactly these, case-sensitive and
#: untrimmed. Over an underlining style each removes the underline, ``true`` as
#: much as ``false``, and Word saves ``w:val="none"``. ``True``, ``TRUE``, ``yes``
#: and ``no`` are outside the enumeration, so unset.
UNDERLINE_ALIASES: dict[str, str] = {word: "none" for word in ("true", "false", "on", "off", "1", "0")}


def in_enumeration(value: Any, enumeration: type[Enum], aliases: dict[str, str] | None = None) -> Any:
    """A member of `enumeration`, or None: Word's reading of an enumerated ``w:val`` (rule 2).

    A member is itself. A ``str`` is the member whose value it is **exactly**,
    case, spacing and all (``"SINGLE"``, ``"Yellow"``, ``" yellow"`` and
    ``"bogus"`` are none, and so are ``""`` and anything else the exact parse
    kept as its text); else the member `aliases` names for it (``w:u``'s
    :data:`UNDERLINE_ALIASES`). None, and anything else, is None: the value is
    unset, as if absent, so a style's value beneath shows (rounds two, five and
    six). Word repairs a missing or invalid ``highlight``, ``vertAlign`` or
    ``em``, dropping the element; an invalid ``w:u`` it opens without a prompt.

    A padded member (``" yellow"``) is read as outside the enumeration until a
    probe measures Word's whitespace there (CR-018 Q15): Word is case-exact for
    the enumerations and untrimmed for on/off.
    """
    if value is None:
        return None
    if isinstance(value, enumeration):
        return value
    if isinstance(value, Enum):  # a member of another enumeration: its value read as text
        value = value.value
    if not isinstance(value, str):
        return None
    try:
        return enumeration(value)
    except ValueError:
        pass
    if aliases is not None:
        alias = aliases.get(value)
        if alias is not None:
            return enumeration(alias)
    return None


_ENUMERATIONS: dict[str, type[Enum]] = {}


def _enumeration(name: str) -> type[Enum]:
    """A ``wml`` enumeration by name, imported on first use: this module imports no model class at import."""
    found = _ENUMERATIONS.get(name)
    if found is None:
        from docx4j_py import wml

        found = _ENUMERATIONS[name] = getattr(wml, name)
    return found


def underline(value: Any) -> Any:
    """``w:u/@w:val`` as Word reads it: an ``UnderlineEnumeration`` member, or None for unset.

    :func:`in_enumeration` with :data:`UNDERLINE_ALIASES`: the six on/off words
    are ``NONE``; anything else outside the enumeration is None.
    """
    return in_enumeration(value, _enumeration("UnderlineEnumeration"), UNDERLINE_ALIASES)


def highlight(value: Any) -> Any:
    """``w:highlight/@w:val`` as Word reads it: a ``HighlightVal`` member, or None for unset (case-exact)."""
    return in_enumeration(value, _enumeration("HighlightVal"))


def vertical_alignment(value: Any) -> Any:
    """``w:vertAlign/@w:val`` as Word reads it: an ``STVerticalAlignRun`` member, or None for unset."""
    return in_enumeration(value, _enumeration("STVerticalAlignRun"))


def emphasis_mark(value: Any) -> Any:
    """``w:em/@w:val`` as Word reads it: an ``STEm`` member, or None for unset."""
    return in_enumeration(value, _enumeration("STEm"))


def theme_colour(value: Any) -> Any:
    """A theme reference (``w:themeColor``, ``w:themeFill``; ``ST_ThemeColor``) as read: an ``STThemeColor`` member, or None.

    A value outside the enumeration (``Accent1``, kept as its text by the exact
    parse) names nothing, as docx4j reads it (JAXB unmarshals it as null), so
    it states no reference and a style's colour beneath shows. Word's reading
    of this attribute is unmeasured; the rule is CR-018's for every
    enumeration (CR-009 section 12.10 item 1).
    """
    return in_enumeration(value, _enumeration("STThemeColor"))


# ---------------------------------------------------------------------------
# colour (CR-018 Phase B, section 4.3)
# ---------------------------------------------------------------------------

#: What :func:`colour` answers for a stated automatic colour: exactly ``auto``.
AUTO = "auto"

#: The classic HTML (Netscape) colour names Word 365 reads in a colour, lower case, to
#: their ``RRGGBB``: 140 names, docx4j's ``ColorTables.HTML`` (``868a1797b``; since CR-061,
#: ``WordColorValues.HTML``), which is CSS 3's 147 keywords less ``grey``, its five other
#: ``grey`` spellings and ``lightgray`` (rule 3; CR-018 Q7, docx4j's D5). Round four probed
#: all seven, and Word dropped each; 41 of the 140 are measured, each saved as this hex.
NETSCAPE_COLOURS: dict[str, str] = {
    "aliceblue": "F0F8FF", "antiquewhite": "FAEBD7", "aqua": "00FFFF", "aquamarine": "7FFFD4", "azure": "F0FFFF",
    "beige": "F5F5DC", "bisque": "FFE4C4", "black": "000000", "blanchedalmond": "FFEBCD", "blue": "0000FF",
    "blueviolet": "8A2BE2", "brown": "A52A2A", "burlywood": "DEB887", "cadetblue": "5F9EA0",
    "chartreuse": "7FFF00", "chocolate": "D2691E", "coral": "FF7F50", "cornflowerblue": "6495ED",
    "cornsilk": "FFF8DC", "crimson": "DC143C", "cyan": "00FFFF", "darkblue": "00008B", "darkcyan": "008B8B",
    "darkgoldenrod": "B8860B", "darkgray": "A9A9A9", "darkgreen": "006400", "darkkhaki": "BDB76B",
    "darkmagenta": "8B008B", "darkolivegreen": "556B2F", "darkorange": "FF8C00", "darkorchid": "9932CC",
    "darkred": "8B0000", "darksalmon": "E9967A", "darkseagreen": "8FBC8F", "darkslateblue": "483D8B",
    "darkslategray": "2F4F4F", "darkturquoise": "00CED1", "darkviolet": "9400D3", "deeppink": "FF1493",
    "deepskyblue": "00BFFF", "dimgray": "696969", "dodgerblue": "1E90FF", "firebrick": "B22222",
    "floralwhite": "FFFAF0", "forestgreen": "228B22", "fuchsia": "FF00FF", "gainsboro": "DCDCDC",
    "ghostwhite": "F8F8FF", "gold": "FFD700", "goldenrod": "DAA520", "gray": "808080", "green": "008000",
    "greenyellow": "ADFF2F", "honeydew": "F0FFF0", "hotpink": "FF69B4", "indianred": "CD5C5C", "indigo": "4B0082",
    "ivory": "FFFFF0", "khaki": "F0E68C", "lavender": "E6E6FA", "lavenderblush": "FFF0F5", "lawngreen": "7CFC00",
    "lemonchiffon": "FFFACD", "lightblue": "ADD8E6", "lightcoral": "F08080", "lightcyan": "E0FFFF",
    "lightgoldenrodyellow": "FAFAD2", "lightgreen": "90EE90", "lightgrey": "D3D3D3", "lightpink": "FFB6C1",
    "lightsalmon": "FFA07A", "lightseagreen": "20B2AA", "lightskyblue": "87CEFA", "lightslategray": "778899",
    "lightsteelblue": "B0C4DE", "lightyellow": "FFFFE0", "lime": "00FF00", "limegreen": "32CD32",
    "linen": "FAF0E6", "magenta": "FF00FF", "maroon": "800000", "mediumaquamarine": "66CDAA",
    "mediumblue": "0000CD", "mediumorchid": "BA55D3", "mediumpurple": "9370DB", "mediumseagreen": "3CB371",
    "mediumslateblue": "7B68EE", "mediumspringgreen": "00FA9A", "mediumturquoise": "48D1CC",
    "mediumvioletred": "C71585", "midnightblue": "191970", "mintcream": "F5FFFA", "mistyrose": "FFE4E1",
    "moccasin": "FFE4B5", "navajowhite": "FFDEAD", "navy": "000080", "oldlace": "FDF5E6", "olive": "808000",
    "olivedrab": "6B8E23", "orange": "FFA500", "orangered": "FF4500", "orchid": "DA70D6",
    "palegoldenrod": "EEE8AA", "palegreen": "98FB98", "paleturquoise": "AFEEEE", "palevioletred": "DB7093",
    "papayawhip": "FFEFD5", "peachpuff": "FFDAB9", "peru": "CD853F", "pink": "FFC0CB", "plum": "DDA0DD",
    "powderblue": "B0E0E6", "purple": "800080", "red": "FF0000", "rosybrown": "BC8F8F", "royalblue": "4169E1",
    "saddlebrown": "8B4513", "salmon": "FA8072", "sandybrown": "F4A460", "seagreen": "2E8B57",
    "seashell": "FFF5EE", "sienna": "A0522D", "silver": "C0C0C0", "skyblue": "87CEEB", "slateblue": "6A5ACD",
    "slategray": "708090", "snow": "FFFAFA", "springgreen": "00FF7F", "steelblue": "4682B4", "tan": "D2B48C",
    "teal": "008080", "thistle": "D8BFD8", "tomato": "FF6347", "turquoise": "40E0D0", "violet": "EE82EE",
    "wheat": "F5DEB3", "white": "FFFFFF", "whitesmoke": "F5F5F5", "yellow": "FFFF00", "yellowgreen": "9ACD32",
}

_HEX_DIGITS = frozenset("0123456789abcdefABCDEF")
#: XML's whitespace, which is what Word trims from a colour (``" FF0000"``, ``"FF0000 "``).
_XML_SPACE = " \t\r\n"


def colour(value: Any) -> str | None:
    """An ST_HexColor value as Word 365 reads it: ``"RRGGBB"`` upper case, :data:`AUTO`, or None.

    * **exactly** ``auto`` (the ``STHexColorAuto`` member, or the text) is
      :data:`AUTO`, a stated automatic (round six: ``Auto``, ``AUTO``,
      ``" auto"``, ``"auto "`` and ``#auto`` are not);
    * otherwise the text (a parsed hexBinary's source spelling, which the fork's
      ``lexical_types`` keeps; a kept ``str``) is trimmed of whitespace, loses
      one leading ``#``, and is the colour if it is six hex digits in any case;
    * else a name in :data:`NETSCAPE_COLOURS`, in any case (``red``, ``DarkRed``);
    * anything else is None: ``F00``, ``0xFF0000``, five or seven digits,
      ``GG0000``, ``""``, a name outside the table (``grey``, ``rebeccapurple``,
      ``windowtext``, ``dark red``).

    **What None means is the attribute's** (rounds five, six and seven): as if
    absent for ``w:color/@w:val`` and ``w:u/@w:color``, so a style's colour
    beneath shows; automatic, and overriding a style's, for ``w:shd/@w:fill``,
    ``w:shd/@w:color`` and a border's ``w:color``; no page colour for
    ``w:background/@w:color``. A stated ``w:themeColor`` is CR-009 Phase B's and
    wins over ``w:val`` as Word reads it; this reads ``w:val`` only.
    """
    if value is None:
        return None
    if isinstance(value, Enum):
        value = value.value
    if isinstance(value, (bytes, bytearray)):
        lexical = getattr(value, "lexical", None)
        if lexical is None:  # bytes set in code, or the converter's own spelling: the value itself
            return value.hex().upper() if len(value) == 3 else None
        value = lexical
    text = str(value)
    if text == AUTO:
        return AUTO
    trimmed = text.strip(_XML_SPACE)
    digits = trimmed.removeprefix("#")
    if len(digits) == 6 and all(c in _HEX_DIGITS for c in digits):
        return digits.upper()
    return NETSCAPE_COLOURS.get(trimmed.lower())


def states_colour(value: Any) -> bool:
    """Whether a ``w:color/@w:val`` or ``w:u/@w:color`` value states a colour, as Word reads it.

    docx4j's ``WordColorValues.stated``: one :func:`colour` reads, exactly
    ``auto`` included. Any other value is as if absent, so the merge passes it
    over and a style's colour shows (rounds five and six). A fill's and a
    border's are read otherwise: there an invalid value is a stated automatic.
    """
    return colour(value) is not None


# ---------------------------------------------------------------------------
# integers (CR-021 section 4.1)
# ---------------------------------------------------------------------------

#: Word's lexical form for an integer attribute (docx4j CR-065 D3 (a), rounds two to four):
#: an optional sign, ASCII digits, optionally a dot and ASCII digits, either digit group
#: possibly empty. It matches ``""`` too, which Word reads as 0. ``[0-9]``, never ``\d``,
#: which matches every Unicode digit; :func:`re.fullmatch`, never ``$``, which matches
#: before a trailing newline; and no surrounding whitespace, which Word refuses (round four).
_WORD_INTEGER = re.compile(r"([+-]?)([0-9]*)(?:\.[0-9]*)?")


def integer(value: Any) -> int | None:
    """A WordprocessingML integer attribute as Word 365 reads it: an ``int``, or None for unset.

    * an ``int`` is itself (what the parse binds for an integer; a ``bool``,
      which is no integer attribute's value, is None);
    * a ``str`` (a value the parse kept as its text: a decimal, ``""``) in
      Word's form, :data:`_WORD_INTEGER`, reads as its integer part with its
      sign, the fraction dropped: truncation toward zero, done on the string,
      never through ``float`` (``float("2.9999999999999999")`` is 3.0, where
      the string's integer part is 2). ``"720.5"`` is 720, ``"-121.5"`` -121,
      ``"141.99999999999994"`` 141, ``"240."`` 240, ``"+240.5"`` 240, ``".5"``
      and ``"-0.5"`` 0;
    * ``""`` is 0, as Word reads it (SSRS's ``w:header=""``); so is a form
      whose digit groups are both empty (``"."``, ``"+"``, ``"-"``), as
      docx4j's test reads it (unmeasured in Word);
    * anything else is None, unset, as if absent: ``"abc"``, ``"1e3"``,
      ``"0x10"``, ``"240,5"``, a value with surrounding whitespace
      (``" 240 "``, ``" 240.5 "``), ``"1_000"``, other scripts' digits. Word
      refuses a document holding any of them; a save drops such a value
      (:mod:`docx4j_py.openpackaging.normalise`).

    Word's twips are 16 bits wide inside, and a value beyond them wraps
    (docx4j CR-065's aside, CR-022); that is not emulated (CR-021 Q7): the
    integer is the file's. A digit string longer than Python converts
    (``sys.get_int_max_str_digits()``, 4,300 by default) is None rather than
    an exception: no measurement covers it.
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if not isinstance(value, str):
        return None
    found = _WORD_INTEGER.fullmatch(value)
    if found is None:
        return None
    sign, digits = found.group(1), found.group(2)
    if not digits:
        return 0
    try:
        whole = int(digits)
    except ValueError:  # beyond Python's digit limit for a conversion (sys.get_int_max_str_digits)
        return None
    return -whole if sign == "-" else whole


# ---------------------------------------------------------------------------
# effective(): the value the file leaves out (CR-018 Phase C, section 4.4)
# ---------------------------------------------------------------------------

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

#: The namespaces whose schema defaults have been audited, and only these (CR-018 Q10):
#: ``w``, by docx4j's #697 and CR-059 and the refresh from ``46c727bf8`` (52 fields, every
#: one ``true`` or ``false``); ``w14``, whose three are docx4j's own patch of ST_OnOff
#: (``w14:CT_OnOff``, ``conflictMode`` and ``discardImageEditingData``, default true, which
#: docx4j's getters answer for a bare element; unmeasured in Word, which always writes
#: ``w14:val``); and ``w15``, which declares none of its own (its on/off elements take
#: ``w:val``). Unaudited defaults are how docx4j's getters went wrong (#697, #641, the
#: autospacing bug, CR-059's ``anchorLock`` regression, the FontEmbedder bug), so anywhere
#: else :func:`effective` raises until someone audits that namespace.
AUDITED_NAMESPACES: frozenset[str] = frozenset({W, W14, W15})

_ATTRIBUTES: dict[type, dict[str, Any]] = {}
_ATTRIBUTES_LOCK = threading.Lock()


def _attribute_vars(cls: type) -> dict[str, Any]:
    """A class's attribute vars by field name, built once per class."""
    found = _ATTRIBUTES.get(cls)
    if found is None:
        from docx4j_py.runtime import context

        found = {var.name: var for var in context().build(cls).get_attribute_vars()}
        with _ATTRIBUTES_LOCK:
            _ATTRIBUTES[cls] = found
    return found


def effective(obj: Any, name: str) -> Any:
    """The value `obj` states for attribute `name`, else the schema's default, typed; else None.

    The stated value wins whatever it is, a kept ``str`` included (``" false"``,
    ``"red"``): the readers above decide what it means. Absent, it is the
    fork's ``XmlVar.schema_default_value`` (stage 13, item 31), the schema
    default converted once as a parse would convert it: ``True`` for
    ``"true"``, an enumeration's member. The field itself never changes: it
    stays None when absent, and nothing but this reads the default.

    Raises:
        LookupError: `name` is not an attribute of `obj`'s class, or the
            attribute's namespace is not in :data:`AUDITED_NAMESPACES`
            (CR-018 Q10).
    """
    var = _attribute_vars(type(obj)).get(name)
    if var is None:
        raise LookupError(f"{type(obj).__name__} has no attribute field {name!r}")
    qname = var.qname or ""
    namespace = qname[1:].split("}", 1)[0] if qname.startswith("{") else ""
    if namespace not in AUDITED_NAMESPACES:
        raise LookupError(
            f"{type(obj).__name__}.{name}: the schema defaults of {namespace or 'no namespace'!r} are not audited "
            "(CR-018 Q10); read the field, which is None when the attribute is absent"
        )
    value = getattr(obj, name, None)
    if value is not None:
        return value
    return var.schema_default_value


# ---------------------------------------------------------------------------
# measurements as code sets them (CR-022)
# ---------------------------------------------------------------------------

#: The limits table (CR-022 D1): hand-written, in ``docx4j_py/resources/``.
INTEGER_LIMITS_RESOURCE = "integer-limits.json"

#: The unit of every attribute the table lists as unmeasured (``w:gridCol``, ``w:wBefore``,
#: ``w:wAfter``, ``w:legacy``, ``w:fitText``, ``w:trHeight`` under ``hRule="auto"``): twips.
_UNMEASURED_UNIT = "twips"

#: A row's context: the sibling attribute and the values the row applies under (None: absent).
Context = tuple[str, tuple[str | None, ...]]
#: An unmeasured attribute's entries: (its context or None, why the probe could not measure it).
Unmeasured = tuple[tuple[Context | None, str], ...]


@dataclasses.dataclass(frozen=True, slots=True)
class IntegerLimit:
    """One attribute's row of the limits table: the range Word 365 reads it faithfully in (CR-022 D1).

    Attributes:
        attribute: the path as the table spells it, ``w:ind/@w:left``, with a
            holder first where the element name alone binds a class whose
            attribute is not this one (``w:rPr/w:sz/@w:val``: ``w:frame/w:sz``
            is text).
        minimum: the least value Word reads as written, in :attr:`unit`.
        maximum: the greatest, inclusive.
        unit: the attribute's unit: ``twips``, ``half-points``, ``eighths of a
            point``, ``points``, ``fiftieths of a percent``, ``percent``,
            ``240ths of a line`` or ``page number``.
        storage: how Word stores it, as the probe found.
        context: (the sibling attribute the range depends on, the values this
            row applies under, None for that attribute absent), or None where
            the range depends on nothing: ``("w:lineRule", ("exact",))``.
        word_writes: a stored value Word's save spells otherwise, to that
            spelling (``w:before``'s 65535 is saved as ``"-1"``). Word keeps
            it, but the range leaves it out by choice, since another reader
            takes the spelling for its number (CR-022 section 10, decision 3).
        sentinel: a value Word writes for many (``w:col/@w:space``'s -1 for
            every negative), so not a measurement it keeps.
        cases: the probe's case ids the row rests on, by round
            (``tests/fixtures/integer-limits/``).
        analogy: where no case measured the row, why it holds (Q3: the other
            border edges).
        note: anything else the row needs said.
    """

    attribute: str
    minimum: int
    maximum: int
    unit: str
    storage: str
    context: Context | None = None
    word_writes: tuple[tuple[int, str], ...] = ()
    sentinel: int | None = None
    cases: tuple[tuple[str, tuple[str, ...]], ...] = ()
    analogy: str | None = None
    note: str | None = None

    @property
    def path(self) -> tuple[str, ...]:
        """The elements of :attr:`attribute`, outermost first: ``("w:rPr", "w:sz")``."""
        return tuple(self.attribute.split("/@", 1)[0].split("/"))

    @property
    def element(self) -> str:
        """The element that carries the attribute: ``w:sz``."""
        return self.path[-1]

    @property
    def holder(self) -> str | None:
        """The element the row's element must be under, where the row names one: ``w:rPr``."""
        return self.path[-2] if len(self.path) > 1 else None

    @property
    def name(self) -> str:
        """The attribute itself: ``w:val``."""
        return self.attribute.split("/@", 1)[1]

    def applies(self, context: str | None) -> bool:
        """Whether this row is the one for `context`, the sibling attribute's value (None: absent)."""
        return self.context is None or context in self.context[1]

    def admits(self, stored: int) -> bool:
        """Whether Word reads `stored` as written: :attr:`minimum` to :attr:`maximum`, inclusive."""
        return self.minimum <= stored <= self.maximum


_LIMITS: tuple[tuple[IntegerLimit, ...], dict[str, Unmeasured]] | None = None
_LIMITS_LOCK = threading.Lock()


def _load_limits() -> tuple[tuple[IntegerLimit, ...], dict[str, Unmeasured]]:
    """The table's rows, one per attribute path, and its unmeasured attributes; read once."""
    global _LIMITS
    if _LIMITS is not None:
        return _LIMITS
    from importlib.resources import files

    data = json.loads(files("docx4j_py.resources").joinpath(INTEGER_LIMITS_RESOURCE).read_text("utf-8"))

    def context_of(raw: dict[str, list[str | None]] | None) -> Context | None:
        if not raw:
            return None
        ((attribute, values),) = raw.items()
        return attribute, tuple(values)

    rows: list[IntegerLimit] = []
    for raw in data["rows"]:
        for attribute in raw["attributes"]:
            rows.append(
                IntegerLimit(
                    attribute=attribute,
                    minimum=int(raw["min"]),
                    maximum=int(raw["max"]),
                    unit=raw["unit"],
                    storage=raw["storage"],
                    context=context_of(raw.get("context")),
                    word_writes=tuple((int(k), v) for k, v in raw.get("word_writes", {}).items()),
                    sentinel=raw.get("sentinel"),
                    cases=tuple((name, tuple(ids)) for name, ids in raw.get("cases", {}).items()),
                    analogy=raw.get("analogy"),
                    note=raw.get("note"),
                )
            )
    unmeasured: dict[str, list[tuple[Context | None, str]]] = {}
    for raw in data.get("unmeasured", ()):
        unmeasured.setdefault(raw["attribute"], []).append((context_of(raw.get("context")), raw["why"]))
    loaded = (tuple(rows), {k: tuple(v) for k, v in unmeasured.items()})
    with _LIMITS_LOCK:
        _LIMITS = loaded
    return loaded


def integer_limits() -> tuple[IntegerLimit, ...]:
    """Every row of the limits table, one per attribute path and context (CR-022 D1)."""
    return _load_limits()[0]


def unmeasured_attributes() -> dict[str, Unmeasured]:
    """The attributes the probe could not measure, each with its context (or None) and why (CR-022 section 3.3)."""
    return _load_limits()[1]


def integer_limit(attribute: str, context: str | None = None) -> IntegerLimit | None:
    """The row for `attribute` under `context` (the sibling attribute's value; None: absent), or None.

    None where the probe could not measure the attribute, or not under that
    context (``w:gridCol/@w:w``; ``w:trHeight/@w:val`` under ``hRule="auto"``,
    which Word discards): such a value is not checked (CR-022 Q3).

    Raises:
        LookupError: the table neither has a row for `attribute` under
            `context` nor lists it as unmeasured: a setter naming an attribute
            the table does not know is a mistake in the setter.
    """
    rows, unmeasured = _load_limits()
    found = [row for row in rows if row.attribute == attribute]
    for row in found:
        if row.applies(context):
            return row
    for where, _why in unmeasured.get(attribute, ()):
        if where is None or context in where[1]:
            return None
    raise LookupError(
        f"the integer limits table has no row for {attribute}"
        + ("" if context is None else f" under {context!r}")
        + " (CR-022 D1; docx4j_py/resources/integer-limits.json)"
    )


def _decimal(value: Any) -> Decimal:
    """`value` as a Decimal, exactly for an int and by its shortest repr for a float."""
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(int(value))
    return Decimal(repr(float(value)))


def _number(value: Decimal) -> str:
    """``1,638.35``: thousands separated, no trailing zeros, no exponent."""
    text = format(value, ",f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return "0" if text in ("-0", "") else text


def _given(value: Any) -> str:
    """The caller's value as they gave it: ``400000`` for ``400000.0``, ``1638.35``, ``nan``."""
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        return repr(value)
    if isinstance(value, int):
        return str(int(value))
    number = float(value)
    if math.isfinite(number) and number == int(number):
        return str(int(number))
    return repr(number)


#: How many twips, half-points, eighths of a point and points make a point.
#: How many of each unit make one point. ``240ths of a line`` is the content API's line spacing under
#: ``w:lineRule="auto"``: Office JS's points are the line as a multiple of 12 pt ("in the Word UI, this value is
#: divided by 12"), and Word's own setter writes ``w:line`` as 20 x the points (core-ts check 71b; CR-024).
_PER_POINT: dict[str, int] = {"twips": 20, "half-points": 2, "eighths of a point": 8, "points": 1, "240ths of a line": 20}
#: How many EMU make one of each length unit (python-docx's ``Length``).
_EMU_PER: dict[str, Decimal] = {
    "twips": Decimal(635),
    "half-points": Decimal(6350),
    "eighths of a point": Decimal("1587.5"),
    "points": Decimal(12700),
}


@dataclasses.dataclass(frozen=True, slots=True)
class Units:
    """The units a caller gives a measurement in, and the conversion its setter has always made (CR-022 D2).

    `convert` takes the caller's value and the attribute's unit and gives the
    integer the setter writes, rounded as that setter rounded it before CR-022;
    `per` gives how many of the caller's units one of the attribute's makes,
    for the message's range; `secondary` an easier reading of a value, for the
    message (``500 cm`` for a ``Length``).
    """

    name: str
    convert: Callable[[Any, str], Any]
    per: Callable[[str], Decimal]
    secondary: Callable[[Any, str], str | None] = lambda _value, _unit: None


def _from_points(value: Any, unit: str) -> int:
    return round(value * _PER_POINT[unit])


def _from_emu(value: Any, unit: str) -> int:
    emu = int(value)
    if unit == "twips":  # python-docx's Length.twips
        return round(emu / 635.0)
    if unit == "half-points":  # round(Length.pt * 2), as the facade's Font.size always wrote it
        return round(emu / 12700.0 * 2)
    return round(emu / float(_EMU_PER[unit]))


def _emu_secondary(value: Any, unit: str) -> str | None:
    emu = Decimal(int(value))
    if unit == "half-points":
        return f"{_number(round(emu / Decimal(12700), 2))} pt"
    return f"{_number(round(emu / Decimal(360000), 3))} cm"


def _from_lines(value: Any, unit: str) -> int:
    if unit != "240ths of a line":
        raise KeyError(unit)
    # python-docx's line_spacing = 1.5: twips_of(Emu(int(value * Twips(240)))), Twips(240) being 152,400 EMU
    return round(int(value * 152400) / 635.0)


def _per_line(unit: str) -> Decimal:
    if unit != "240ths of a line":
        raise KeyError(unit)
    return Decimal(1) / Decimal(240)


def _own(value: Any, unit: str) -> Any:
    return value


#: A measurement in points, as the content API takes every one (Office JS's unit).
POINTS = Units("pt", _from_points, lambda unit: Decimal(1) / _PER_POINT[unit])
#: A python-docx ``Length``, a count of EMU (914,400 to the inch), as the facade takes every one.
EMU = Units("EMU", _from_emu, lambda unit: _EMU_PER[unit], _emu_secondary)
#: A line spacing in lines, python-docx's float: ``w:line`` under ``lineRule="auto"`` is 240ths of one.
LINES = Units("lines", _from_lines, _per_line)
#: The attribute's own unit, as the tree-layer builders take a width (twips).
OWN = Units("", _own, lambda unit: Decimal(1))


@dataclasses.dataclass(frozen=True, slots=True)
class CallerLimit:
    """A setter's own range, narrower than Word's storage (CR-022 Q1, CR-024).

    Word 365's own Office JS setters refuse some values Word would store as
    written (core-ts check 71b, CR-024 section 2): ``spaceBefore`` and
    ``spaceAfter`` take 0 to 1,584 pt (31,680 twips), ``lineSpacing`` 0.75 to
    1,584 pt (15 to 31,680 240ths of a line under ``w:lineRule="auto"``),
    ``setLevelIndents``' text indent -1,583.95 to 1,583.95 pt (31,679 twips
    either way) and ``font.size`` 1 to 1,638 pt after rounding to a half-point
    (2 to 3,276). The content API's setters of the same names refuse where
    Word's do. None at an end means the attribute's own range rules there.

    With `stored`, `low` and `high` are in the attribute's unit and bound the
    value as the setter writes it, as Word's own setters are bound (its font
    setter takes 0.9 pt, written as 1); otherwise they are in the caller's
    units and bound what the caller gave (a cell width's sign).
    """

    low: float | None = None
    high: float | None = None
    low_open: bool = False
    high_open: bool = False
    by: str = "Office JS"
    stored: bool = False

    def admits(self, value: Any) -> bool:
        """Whether `value` (the caller's, or with `stored` the value as written) is inside this limit."""
        if self.low is not None and (value < self.low or (self.low_open and value == self.low)):
            return False
        return not (self.high is not None and (value > self.high or (self.high_open and value == self.high)))


def _inward(value: Decimal, up: bool) -> Decimal:
    """`value` as shown: exact where it has three decimals or fewer, else rounded inward to three (``n/240``)."""
    if value == value.quantize(Decimal("0.001")):
        return value
    return value.quantize(Decimal("0.001"), rounding=ROUND_CEILING if up else ROUND_FLOOR)


def _range_text(low: Decimal, high: Decimal, low_open: bool, high_open: bool, unit: str) -> str:
    """``-1,638.4 to 1,638.35 pt``; ``0 to 1,638 pt, 1,638 excluded``: each bound shown inside the range."""
    lo, hi = _number(_inward(low, up=True)), _number(_inward(high, up=False))
    tail = f" {unit}" if unit else ""
    if low_open and high_open:
        return f"{lo} to {hi}{tail}, both ends excluded"
    if low_open:
        return f"{lo} to {hi}{tail}, {lo} excluded"
    if high_open:
        return f"{lo} to {hi}{tail}, {hi} excluded"
    return f"{lo} to {hi}{tail}"


def _refusal(
    row: IntegerLimit,
    value: Any,
    stored: Any,
    units: Units,
    label: str,
    limit: CallerLimit | None,
    finite: bool = True,
    hint: str | None = None,
) -> Exception:
    """The ``MeasureRangeError`` for `value`: the property, the value both ways, Word's range in the caller's units."""
    from docx4j_py.model.content.errors import MeasureRangeError

    per = units.per(row.unit)
    low, high = _decimal(row.minimum) * per, _decimal(row.maximum) * per
    low_open = high_open = False
    by_limit = False
    if limit is not None:
        scale = per if limit.stored else Decimal(1)
        if limit.low is not None:
            bound = _decimal(limit.low) * scale
            if bound > low or (bound == low and limit.low_open):
                low, low_open, by_limit = bound, limit.low_open, True
        if limit.high is not None:
            bound = _decimal(limit.high) * scale
            if bound < high or (bound == high and limit.high_open):
                high, high_open, by_limit = bound, limit.high_open, True
    unit_name = units.name or row.unit
    span = _range_text(low, high, low_open, high_open, unit_name)
    if units is EMU and row.unit in _EMU_PER:
        # the bounds read in cm (or pt), each rounded inward so that it is inside the range
        if row.unit == "half-points":
            scale, places, name = Decimal(12700), Decimal("0.01"), "pt"
        else:
            scale, places, name = Decimal(360000), Decimal("0.001"), "cm"
        inner_low = (low / scale).quantize(places, rounding=ROUND_CEILING)
        inner_high = (high / scale).quantize(places, rounding=ROUND_FLOOR)
        span = f"{span} ({_range_text(inner_low, inner_high, low_open, high_open, name)})"
    where = f", where {limit.by}'s setter stops" if by_limit else ""  # type: ignore[union-attr]
    shown = f"{label} {_given(value)}" + (f" {units.name}" if units.name else f" {row.unit}")
    if not finite:
        message = f"{shown} is not a finite number; Word's range is {span}{where}"
    else:
        both = []
        secondary = units.secondary(value, row.unit) if units.name else None
        if secondary is not None:
            both.append(secondary)
        if units.name:
            both.append(f"{_number(_decimal(stored))} {row.unit}")
        written = f" ({', '.join(both)})" if both else ""
        message = f"{shown}{written} is beyond Word's range, {span}{where}"
    error = MeasureRangeError(message, hint=f"pass a {label} from {span}" if hint is None else hint.format(span=span))
    error.attribute = row.attribute
    error.value = value
    error.stored = stored
    return error


def measure_in_range(
    attribute: str,
    value: Any,
    units: Units = POINTS,
    *,
    label: str,
    context: str | None = None,
    limit: CallerLimit | None = None,
    hint: str | None = None,
) -> Any:
    """`value` converted to `attribute`'s unit, where Word 365 reads it as written; else ``MeasureRangeError``.

    The one check every setter that writes an integer measurement goes
    through (CR-022 D2, D3), called **before** the setter changes anything,
    so that a refused call leaves the tree as it was and records no
    :class:`~docx4j_py.model.content.reports.ChangeReport`.

    Args:
        attribute: the attribute, as the limits table spells it
            (``"w:ind/@w:left"``, ``"w:rPr/w:sz/@w:val"``).
        value: what the caller gave, in `units`.
        units: :data:`POINTS` (the content API), :data:`EMU` (a python-docx
            ``Length``), :data:`LINES` (python-docx's line multiple) or
            :data:`OWN` (the attribute's own unit, the builders' twips). The
            conversion is the one the setter always made, rounding and all.
        label: the property, as the caller knows it (``"left indent"``),
            for the message.
        context: the value of the sibling attribute the range depends on,
            as the write leaves it: ``w:lineRule`` for ``w:spacing/@w:line``,
            ``w:hRule`` for ``w:trHeight/@w:val``, ``w:type`` for a table
            width; None for that attribute absent.
        limit: the setter's own narrower range, Word's own Office JS setter's
            (CR-022 Q1, CR-024), in the caller's units or with its `stored`
            on the value as written; it applies with the attribute's range.
        hint: the refusal's hint, where the default (``pass a <label> from
            <range>``) would not say what to do: a format string, ``{span}``
            the range as the message shows it. A rule setter's names the
            value already set and the rule (CR-022 section 10, decision 4).

    Returns:
        The value to write, in the attribute's unit. Not checked where the
        table lists the attribute (or this context) as unmeasured (Q3).

    Raises:
        MeasureRangeError: ``value.out_of_range``; the message names the
            property, the value in the caller's units and in the attribute's,
            and Word's range in the caller's units (Q6), and the hint the
            range. Also for a value that is not a finite number.
        LookupError: the table does not know `attribute` (a setter's mistake).
    """
    row = integer_limit(attribute, context)
    if row is None:  # unmeasured (Q3): converted as the setter always did, and not checked
        return units.convert(value, _UNMEASURED_UNIT)
    if isinstance(value, float) and not math.isfinite(value):
        raise _refusal(row, value, None, units, label, limit, finite=False, hint=hint)
    stored = units.convert(value, row.unit)
    if row.admits(stored) and (limit is None or limit.admits(stored if limit.stored else value)):
        return stored
    raise _refusal(row, value, stored, units, label, limit, hint=hint)

