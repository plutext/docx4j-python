"""An equation's text as Office JS gives it (CR-013 Phase A, Q1).

The port of docx4j-core-ts's ``src/model/content/mathText.mts`` at ``9307bdb``
(its CR-012 phase 4; Word 365 checks 48 case 17 and 51). Office JS's
``paragraph.text`` reads an equation (``m:oMath``, ``m:oMathPara``) as the
characters Word draws, not the ones the file stores:

* the ``m:t`` text of its math runs in document order, a revision's runs as the
  view has them (a ``w:del`` or ``w:moveFrom`` out in the accepted view, a
  ``w:ins`` or ``w:moveTo`` out in the original one);
* each letter in the mathematical alphabet Word draws it in, by the run's
  ``m:sty`` and ``m:scr`` (:func:`math_letter`); a run under ``m:nor`` as it
  stands;
* a hyphen-minus as the minus sign U+2212 (check 51 case 6);
* each structure --- a fraction, a script, a radical, an n-ary operator, a
  delimiter, a function --- written ``,`` then its arguments joined by ``-``
  then ``.``: a fraction a over b is ``,𝑎-𝑏.``, sin x ``,sin-𝑥.`` (check 51
  case 12); an argument a structure hides (a radical's degree under
  ``m:degHide``) is left out.

This is not :mod:`docx4j_py.model.math` (OMML to LaTeX), which is the text
model's. :func:`math_text_of` is what :func:`~.text_model.segments_of` makes an
equation's one segment of.
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.model.values import on_off
from docx4j_py.traversal import element_name

__all__ = ["is_math", "math_letter", "math_text_of"]

_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: An equation at run-holder level.
MATH_NAMES = frozenset({f"{_M}oMath", f"{_M}oMathPara"})

_M_T = f"{_M}t"
_W_T = f"{_W}t"

#: The four revision holders and the view each one is hidden in.
_HIDDEN_IN = {
    f"{_W}del": "accepted",
    f"{_W}moveFrom": "accepted",
    f"{_W}ins": "original",
    f"{_W}moveTo": "original",
}


def is_math(element: Any) -> bool:
    """True for an equation at run-holder level: ``m:oMath`` or ``m:oMathPara``."""
    return element_name(element) in MATH_NAMES


def _math_classes() -> tuple[type, type, type, type]:
    from docx4j_py import math as m

    return m.CTOMath, m.CTOMathPara, m.CTOMathArg, m.CTR


def _is_math_object(value: Any) -> bool:
    return type(value).__module__ == "docx4j_py.math"


def _is_properties(value: Any) -> bool:
    """A math object's properties (``m:fPr``, ``m:radPr``, ``m:ctrlPr``, ...): no text."""
    if not _is_math_object(value):
        return False
    name = type(value).__name__
    return name.endswith(("Pr", "PR")) or name == "CTCtrlPr"


def _hidden(value: Any, view: str) -> bool:
    """A ``w:ins``, ``w:del`` or ``w:move*`` around math runs, which this view leaves out."""
    if type(value).__module__ != "docx4j_py.wml":
        return False
    return _HIDDEN_IN.get(element_name(value) or "") == view


def _fields(value: Any) -> list[Any]:
    """A dataclass's field values in schema order."""
    try:
        return [getattr(value, f.name, None) for f in dataclasses.fields(value)]
    except TypeError:
        return []


def math_text_of(math: Any, view: str = "accepted") -> str:
    """The text of an equation (``m:oMath``, ``m:oMathPara``) in a view, as Word draws it.

    The ``m:t`` of every math run in document order, a revision's runs as the
    view has them, each letter mapped as Word draws it (:func:`math_letter`),
    each structure ``,`` + its arguments joined by ``-`` + ``.``; a structure
    the view hides (its own ``m:ctrlPr``'s revision) left out whole, so that
    Word's form of an equation deleted whole reads ``""`` in the accepted view
    (CR-013 section 13.7, F2).
    """
    o_math, o_math_para, o_math_arg, math_run = _math_classes()
    containers = (o_math, o_math_para, o_math_arg)
    out: list[str] = []

    def run_text(run: Any) -> None:
        sty, scr, nor = _run_properties(run)

        def add(text: Any) -> None:
            out.append(_styled("" if text is None else str(text), sty, scr, nor))

        for item in getattr(run, "content", None) or ():
            if isinstance(item, str):
                continue
            name = element_name(item)
            if name in _HIDDEN_IN:
                # a w:ins or w:del inside the math run, around its m:t
                if _hidden(item, view):
                    continue
                for inner in getattr(item, "content", None) or ():
                    if element_name(inner) == _M_T:
                        add(getattr(inner, "value", None))
                continue
            if name == _M_T or name == _W_T:
                add(getattr(item, "value", None))

    def visit(value: Any, depth: int) -> None:
        if value is None or isinstance(value, (str, bytes, int, float, bool)) or depth > 64:
            return
        if isinstance(value, list):
            for item in value:
                visit(item, depth + 1)
            return
        if _hidden(value, view):
            return
        if isinstance(value, math_run):
            run_text(value)
            return
        if _is_properties(value):
            return
        if _is_math_object(value) and not isinstance(value, containers):
            # a structure the view hides, its own m:ctrlPr holding the revision, goes whole
            # (CR-013 section 13.7, F2; core-ts's hidesStructure at 906d5e3)
            if _hides_structure(value, view):
                return
            # a structure: ",", its arguments joined by "-", "." (check 51 case 12)
            out.append(",")
            for index, argument in enumerate(_arguments_of(value, o_math_arg)):
                if index > 0:
                    out.append("-")
                visit(argument, depth + 1)
            out.append(".")
            return
        for field in _fields(value):
            visit(field, depth + 1)

    visit(math, 0)
    return "".join(out)


def _hides_structure(structure: Any, view: str) -> bool:
    """True when a structure is inserted or deleted, in Word's form, and the view hides that revision.

    The revision is a ``w:ins`` or ``w:del`` in the ``m:ctrlPr`` of the
    structure's own properties (``m:sSupPr``, ``m:fPr``, ...); an argument's own
    ``m:ctrlPr`` marks no structure (core-ts's ``hidesStructure`` at
    ``906d5e3``, its CR-002 section 51).
    """
    for value in _fields(structure):
        if not _is_math_object(value) or not type(value).__name__.endswith("Pr"):
            continue
        ctrl_pr = getattr(value, "ctrl_pr", None)
        marker = getattr(ctrl_pr, "r_pr_or_ins_or_del", None) if ctrl_pr is not None else None
        if marker is not None and _hidden(marker, view):
            return True
    return False


def _arguments_of(structure: Any, o_math_arg: type) -> list[Any]:
    """A structure's arguments (``m:e``, ``m:num``, ``m:den``, ``m:sub``, ``m:sup``, ``m:deg``, ``m:lim``, ``m:fName``) in schema order.

    A matrix's rows' arguments in turn, and a radical's degree left out where
    ``m:degHide`` hides it.
    """
    out: list[Any] = []
    rad_pr = getattr(structure, "rad_pr", None)
    deg_hidden = _on(getattr(rad_pr, "deg_hide", None)) if rad_pr is not None else False

    def take(value: Any) -> None:
        if value is None or isinstance(value, (str, bytes, int, float, bool)):
            return
        if isinstance(value, list):
            for item in value:
                take(item)
            return
        if isinstance(value, o_math_arg):
            out.append(value)
            return
        if _is_properties(value):
            return
        if _is_math_object(value):
            for field in _fields(value):
                take(field)

    try:
        fields = dataclasses.fields(structure)
    except TypeError:
        return out
    for field in fields:
        if field.name == "deg" and deg_hidden:
            continue
        take(getattr(structure, field.name, None))
    return out


def _value(obj: Any) -> Any:
    val = getattr(obj, "val", None)
    return getattr(val, "value", val)


def _on(prop: Any) -> bool:
    """An ``m:`` on/off property: present and not exactly ``0``, ``false`` or ``off`` (CR-018's rule)."""
    return prop is not None and on_off(_value(prop))


def _run_properties(run: Any) -> tuple[str | None, str | None, bool]:
    """A math run's ``m:sty``, ``m:scr`` and ``m:nor``."""
    from docx4j_py import math as m

    r_pr = getattr(run, "r_pr", None)
    if r_pr is None:
        # Word's form of a deleted equation holds the run's m:rPr inside the w:del in it (CR-013 section 13.7, D2)
        r_pr = next(
            (
                getattr(item, "r_pr", None)
                for item in getattr(run, "content", None) or ()
                if element_name(item) in _HIDDEN_IN and getattr(item, "r_pr", None) is not None
            ),
            None,
        )
    if r_pr is None:
        return None, None, False
    sty = scr = None
    nor = False
    for item in getattr(r_pr, "nor_or_scr_or_sty", None) or ():
        if isinstance(item, m.CTStyle):
            sty = _value(item)
        elif isinstance(item, m.CTScript):
            scr = _value(item)
        elif isinstance(item, m.CTOnOff):
            nor = _on(item)
    return sty, scr, nor


def _styled(text: str, sty: str | None, scr: str | None, nor: bool) -> str:
    if nor:
        return text  # normal text: as the run's w:rPr draws it (check 51 case 11)
    # a minus sign (check 51 case 6)
    return "".join("−" if c == "-" else math_letter(c, sty, scr) for c in text)


# Unicode's Mathematical Alphanumeric Symbols: each alphabet's capital A and small a,
# its digit zero where it has digits, and its Greek block (capitals, nabla, smalls,
# then the variants).
_LATIN: dict[str, tuple[int, int]] = {
    "bold": (0x1D400, 0x1D41A),
    "italic": (0x1D434, 0x1D44E),
    "boldItalic": (0x1D468, 0x1D482),
    "script": (0x1D49C, 0x1D4B6),
    "boldScript": (0x1D4D0, 0x1D4EA),
    "fraktur": (0x1D504, 0x1D51E),
    "doubleStruck": (0x1D538, 0x1D552),
    "boldFraktur": (0x1D56C, 0x1D586),
    "sans": (0x1D5A0, 0x1D5BA),
    "sansBold": (0x1D5D4, 0x1D5EE),
    "sansItalic": (0x1D608, 0x1D622),
    "sansBoldItalic": (0x1D63C, 0x1D656),
    "monospace": (0x1D670, 0x1D68A),
}
_DIGITS: dict[str, int] = {
    "bold": 0x1D7CE,
    "doubleStruck": 0x1D7D8,
    "sans": 0x1D7E2,
    "sansBold": 0x1D7EC,
    "monospace": 0x1D7F6,
}
_GREEK: dict[str, int] = {
    "bold": 0x1D6A8,
    "italic": 0x1D6E2,
    "boldItalic": 0x1D71C,
    "sansBold": 0x1D756,
    "sansBoldItalic": 0x1D790,
}
#: The letters Unicode left where they were before the block: the BMP's letterlike symbols.
_HOLES: dict[str, dict[str, int]] = {
    "italic": {"h": 0x210E},
    "script": {
        "B": 0x212C,
        "E": 0x2130,
        "F": 0x2131,
        "H": 0x210B,
        "I": 0x2110,
        "L": 0x2112,
        "M": 0x2133,
        "R": 0x211B,
        "e": 0x212F,
        "g": 0x210A,
        "o": 0x2134,
    },
    "fraktur": {"C": 0x212D, "H": 0x210C, "I": 0x2111, "R": 0x211C, "Z": 0x2128},
    "doubleStruck": {
        "C": 0x2102,
        "H": 0x210D,
        "N": 0x2115,
        "P": 0x2119,
        "Q": 0x211A,
        "R": 0x211D,
        "Z": 0x2124,
    },
}
_GREEK_VARIANTS = {0x2202: 51, 0x3F5: 52, 0x3D1: 53, 0x3F0: 54, 0x3D5: 55, 0x3F1: 56, 0x3D6: 57}


def _greek_index(cp: int) -> int | None:
    """A Greek letter's place in an alphabet's Greek block of 58."""
    if 0x391 <= cp <= 0x3A1:
        return cp - 0x391  # Alpha to Rho
    if cp == 0x3F4:
        return 17  # capital theta symbol
    if 0x3A3 <= cp <= 0x3A9:
        return cp - 0x3A3 + 18  # Sigma to Omega
    if cp == 0x2207:
        return 25  # nabla
    if 0x3B1 <= cp <= 0x3C9:
        return cp - 0x3B1 + 26  # alpha to omega, final sigma among them
    return _GREEK_VARIANTS.get(cp)


def math_letter(c: str, sty: str | None, scr: str | None) -> str:
    """A character as Word draws it in a math run with this ``m:sty`` and ``m:scr``.

    ``m:sty`` is ``p``, ``b``, ``i`` or ``bi``, none being Word's default,
    italic; ``m:scr`` is ``roman`` (the default), ``script``, ``fraktur``,
    ``double-struck``, ``sans-serif`` or ``monospace``. As check 51 measured:
    letters, Greek capitals, small letters and nabla alike, in the alphabet of
    the style and script; digits upright in the script's family (bold for bold
    or bold italic, sans-serif under sans-serif, none of their own in script or
    fraktur); a character no alphabet has is itself.
    """
    cp = ord(c)
    bold = sty in ("b", "bi")
    italic = sty is None or sty in ("i", "bi")
    alphabet: str | None
    digits: str | None = None
    script = scr or "roman"
    if script == "roman":
        alphabet = ("boldItalic" if italic else "bold") if bold else ("italic" if italic else None)
        digits = "bold" if bold else None
    elif script == "script":
        alphabet = "boldScript" if bold else "script"
    elif script == "fraktur":
        alphabet = "boldFraktur" if bold else "fraktur"
    elif script == "double-struck":
        alphabet = "doubleStruck"
        digits = "doubleStruck"
    elif script == "sans-serif":
        if bold:
            alphabet = "sansBoldItalic" if italic else "sansBold"
        else:
            alphabet = "sansItalic" if italic else "sans"
        digits = "sansBold" if bold else "sans"
    elif script == "monospace":
        alphabet = "monospace"
        digits = "monospace"
    else:
        alphabet = None
    if 0x30 <= cp <= 0x39:
        zero = None if digits is None else _DIGITS.get(digits)
        return c if zero is None else chr(zero + cp - 0x30)
    if alphabet is None:
        return c
    hole = _HOLES.get(alphabet, {}).get(c)
    if hole is not None:
        return chr(hole)
    latin = _LATIN[alphabet]
    if 0x41 <= cp <= 0x5A:
        return chr(latin[0] + cp - 0x41)
    if 0x61 <= cp <= 0x7A:
        return chr(latin[1] + cp - 0x61)
    g = _greek_index(cp)
    if g is None:
        return c
    base = _GREEK.get(alphabet)
    return c if base is None else chr(base + g)
