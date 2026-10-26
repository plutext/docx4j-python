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
"""

from __future__ import annotations

import re
import threading
from enum import Enum
from typing import Any

__all__ = [
    "AUDITED_NAMESPACES",
    "AUTO",
    "NETSCAPE_COLOURS",
    "OFF",
    "UNDERLINE_ALIASES",
    "XSD_FALSE",
    "XSD_TRUE",
    "colour",
    "effective",
    "emphasis_mark",
    "highlight",
    "in_enumeration",
    "integer",
    "is_on",
    "on_off",
    "on_off_normal_form",
    "states_colour",
    "theme_colour",
    "underline",
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
