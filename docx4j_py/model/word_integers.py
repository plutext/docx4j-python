"""What Word 365 reads for an integer attribute's value: the wrap, per row of the limits table (CR-020 Phase B).

Word keeps each integer attribute in a field of its binary format's width, 16 bits
for most, and reads a value past it **wrapped**, sometimes then clamped, with no
error and no repair (docx4j CR-066 rounds 5 and 6, Word 365, 2026-10-26; CR-022).
CR-022's table (``resources/integer-limits.json``, :func:`docx4j_py.model.values.integer_limit`)
says the range Word keeps each attribute in as written; it carries no model of
the wrap, and its ``storage`` is prose (CR-022 section 10.10 item 5). This module
is that model, for CR-020's out-of-range report (section 4.7):

* :data:`STORAGE`, one :class:`Storage` per attribute path of the table: the
  width, the sign and the special rules, as the objects package's
  ``WORD_INTEGER_STORAGE`` gives them (``docx4j-generated-objects-ts``,
  ``src/helpers/wml.mts``, its CR-016), which docx4j-core-ts's report reads;
* :attr:`Storage.token`, the ``storage`` string of the report, **exactly** the
  objects package's ``storageOf`` (``int16``, ``uint16, then 144..31681``,
  ``int16, the sign the rule``, ...), so that one search finds every port's
  report;
* :func:`reading`, what Word reads: the objects package's ``wordIntegerReading``,
  step for step, in Python's spelling (a value, the rule Word reads with it
  where the sign carries one, and the attribute where ``w:firstLine`` and
  ``w:hanging``'s shared field moved it);
* :func:`kept_range`, the objects package's ``rangeOf`` (core-ts's ``keptRange``):
  ``tests/test_validator.py`` holds it to every row's range, as read.

Nothing here is emulated into the tree (CR-021 Q7): the value stays as written,
and the report says what Word reads.
"""

from __future__ import annotations

import dataclasses
from typing import Any

__all__ = [
    "STORAGE",
    "Reading",
    "Storage",
    "kept_range",
    "read_range",
    "reading",
    "storage_of",
    "wrap",
]


@dataclasses.dataclass(frozen=True, slots=True)
class Storage:
    """How Word stores one integer attribute: the objects package's ``WordIntegerStorage`` for a measured entry.

    Attributes:
        bits: the field's width: 5, 8, 16 or 32.
        signed: whether the field is signed.
        minimum: after the wrap, Word clamps to at least this (a tab stop's
            -31,680, a page's 144).
        maximum: and to at most this (a tab stop's 31,680, a page width's
            31,681).
        sign_rule: the sibling attribute whose rule the stored sign carries,
            ``"lineRule"`` or ``"hRule"`` (an exact line or row height is stored
            negated).
        max_written_as: how Word's save spells the field's largest value: -1
            for 65,535 in ``w:before``, ``w:after`` and the five unsigned page
            margins.
        pair: ``("firstLine", "hanging")``: two attributes sharing one signed
            field, the first stored as itself and the second negated.
        height_flag: ``w:framePr/@w:h``: the low 15 bits are the height, and
            bit 15 clears ``hRule="exact"``.
        negative: what a negative becomes after the wrap: ``"computed"`` (a
            column width Word computes) or -1 (a column's space).
    """

    bits: int
    signed: bool
    minimum: int | None = None
    maximum: int | None = None
    sign_rule: str | None = None
    max_written_as: int | None = None
    pair: tuple[str, str] | None = None
    height_flag: bool = False
    negative: str | int | None = None

    @property
    def token(self) -> str:
        """The report's ``storage``: the objects package's ``storageOf``, character for character."""
        width = f"{'int' if self.signed else 'uint'}{self.bits}" if self.bits in (16, 32) else f"{self.bits} bits"
        if self.pair is not None:
            return f"{width}, shared by {self.pair[0]} (+) and {self.pair[1]} (-)"
        if self.height_flag:
            return "15 bits, bit 15 clearing exact"
        if self.negative is not None:
            return f"{width}, a negative {'computed' if self.negative == 'computed' else 'written -1'}"
        if self.sign_rule is not None:
            return f"{width}, the sign the rule"
        if self.minimum is not None and self.maximum is not None:
            if self.minimum == -self.maximum:
                return f"{width}, then ±{self.maximum}"
            return f"{width}, then {self.minimum}..{self.maximum}"
        if self.minimum is not None:
            return f"{width}, then at least {self.minimum}"
        return width


_INT16 = Storage(16, True)
_UINT16 = Storage(16, False)
_UINT16_MINUS_ONE = Storage(16, False, max_written_as=-1)
_FIRST_LINE_HANGING = Storage(16, True, pair=("firstLine", "hanging"))
_BORDER_SZ = Storage(8, False)
_BORDER_SPACE = Storage(5, False)

_EDGES = ("top", "left", "bottom", "right", "start", "end", "between", "bar", "insideH", "insideV", "tl2br", "tr2bl")

#: Each attribute path of the limits table (:attr:`~docx4j_py.model.values.IntegerLimit.attribute`) to Word's storage of it,
#: as the objects package's ``WORD_INTEGER_STORAGE`` has it: ``CTBorder`` for every border edge and ``w:bdr``,
#: ``TblWidth`` for the table widths and margins, ``HpsMeasure`` for ``w:sz``, ``w:szCs`` and ``w:kern``,
#: ``CTTwipsMeasure`` for the two settings. One entry per path; a path's rows by context share it, the rule
#: being the reading's.
STORAGE: dict[str, Storage] = {
    "w:ind/@w:left": _INT16,
    "w:ind/@w:start": _INT16,
    "w:ind/@w:right": _INT16,
    "w:ind/@w:end": _INT16,
    "w:ind/@w:firstLine": _FIRST_LINE_HANGING,
    "w:ind/@w:hanging": _FIRST_LINE_HANGING,
    "w:spacing/@w:before": _UINT16_MINUS_ONE,
    "w:spacing/@w:after": _UINT16_MINUS_ONE,
    "w:spacing/@w:line": Storage(16, True, sign_rule="lineRule"),
    "w:spacing/@w:val": _INT16,
    "w:position/@w:val": _INT16,
    "w:rPr/w:sz/@w:val": _UINT16,
    "w:szCs/@w:val": _UINT16,
    "w:kern/@w:val": _UINT16,
    "w:w/@w:val": _UINT16,
    "w:tblW/@w:w": _INT16,
    "w:tcW/@w:w": _INT16,
    "w:tblInd/@w:w": _INT16,
    "w:tblCellSpacing/@w:w": _INT16,
    "w:tblCellMar/w:left/@w:w": _INT16,
    "w:tcMar/w:top/@w:w": _INT16,
    "w:trHeight/@w:val": Storage(16, True, sign_rule="hRule"),
    "w:tab/@w:pos": Storage(16, True, minimum=-31680, maximum=31680),
    "w:pgSz/@w:w": Storage(16, False, minimum=144, maximum=31681),
    "w:pgSz/@w:h": Storage(16, False, minimum=144),
    "w:pgMar/@w:top": _INT16,
    "w:pgMar/@w:bottom": _INT16,
    "w:pgMar/@w:left": _UINT16_MINUS_ONE,
    "w:pgMar/@w:right": _UINT16_MINUS_ONE,
    "w:pgMar/@w:header": _UINT16_MINUS_ONE,
    "w:pgMar/@w:footer": _UINT16_MINUS_ONE,
    "w:pgMar/@w:gutter": _UINT16_MINUS_ONE,
    "w:pgNumType/@w:start": Storage(32, True),
    "w:cols/@w:space": _INT16,
    "w:col/@w:w": Storage(16, True, negative="computed"),
    "w:col/@w:space": Storage(16, True, negative=-1),
    "w:lnNumType/@w:distance": _INT16,
    "w:framePr/@w:w": _INT16,
    "w:framePr/@w:h": Storage(16, False, height_flag=True, sign_rule="hRule"),
    "w:framePr/@w:hSpace": _INT16,
    "w:framePr/@w:vSpace": _INT16,
    "w:framePr/@w:x": _INT16,
    "w:framePr/@w:y": _INT16,
    "w:tblpPr/@w:leftFromText": _INT16,
    "w:tblpPr/@w:rightFromText": _INT16,
    "w:tblpPr/@w:topFromText": _INT16,
    "w:tblpPr/@w:bottomFromText": _INT16,
    "w:tblpPr/@w:tblpX": _INT16,
    "w:tblpPr/@w:tblpY": _INT16,
    "w:defaultTabStop/@w:val": _UINT16,
    "w:hyphenationZone/@w:val": _UINT16,
    "w:bdr/@w:sz": _BORDER_SZ,
    "w:bdr/@w:space": _BORDER_SPACE,
    **{f"w:{edge}/@w:sz": _BORDER_SZ for edge in _EDGES},
    **{f"w:{edge}/@w:space": _BORDER_SPACE for edge in _EDGES},
}


def storage_of(attribute: str) -> Storage:
    """The storage of an attribute path of the limits table (``"w:ind/@w:left"``).

    Raises:
        KeyError: the path is not one of the table's.
    """
    return STORAGE[attribute]


@dataclasses.dataclass(frozen=True, slots=True)
class Reading:
    """What Word 365 reads for a value: the objects package's ``WordIntegerReading``, less its ``written``.

    Attributes:
        value: the value Word reads, wrapped and clamped as it stores it; None
            where Word computes one (a negative column width).
        rule: for ``w:spacing/@w:line``, ``w:trHeight`` and ``w:framePr/@w:h``,
            the rule Word reads with it (``"auto"``, ``"exact"``, ``"atLeast"``);
            else None.
        attribute: for ``w:ind``'s ``firstLine`` and ``hanging``, which share
            one field, the one Word reads it as (``"hanging"``); else None.
        kept: whether Word reads the value, and its rule, as written.
        storage: how Word stores it (:attr:`Storage.token`; a line under
            ``auto`` is plain ``int16``).
    """

    value: int | None
    rule: str | None
    attribute: str | None
    kept: bool
    storage: str


def wrap(value: int, bits: int, signed: bool) -> int:
    """`value` modulo ``2**bits``, as a signed or unsigned field of that width holds it."""
    modulus = 1 << bits
    unsigned = value % modulus
    return unsigned - modulus if signed and unsigned >= modulus >> 1 else unsigned


def reading(storage: Storage, attribute: str, value: int, rule: str | None = None) -> Reading:
    """What Word 365 reads for `value` in an attribute stored as `storage`: ``wordIntegerReading``.

    Args:
        storage: the attribute's :class:`Storage`.
        attribute: the attribute's local name (``"hanging"``), for the shared
            first-line and hanging field.
        value: the integer, as Word's lexical reading gives it
            (:func:`docx4j_py.model.values.integer`).
        rule: the ``lineRule`` or ``hRule`` written beside a line, a row's
            height or a frame's height (None: absent, which is ``auto`` for a
            line and ``atLeast`` for a row).

    The rules, as the objects package measured them: the value modulo its
    width, then clamped; under ``exact`` and ``atLeast`` a line or a row height
    keeps its rule in the sign of what is stored (``exact`` is stored negated:
    ``exact 65535`` reads as ``atLeast 1``), under ``auto`` a line is plain
    int16 and a row height is ignored; ``firstLine`` and ``hanging`` are one
    signed field (``firstLine`` 32768 reads as ``hanging`` 32768); a frame's
    height keeps 15 bits and bit 15 clears ``exact``; a negative column width
    is one Word computes, a negative column space -1.
    """
    token = storage.token
    if storage.pair is not None:
        positive, negative = storage.pair
        stored = wrap(-value if attribute == negative else value, 16, True)
        read = abs(stored)
        as_attribute = negative if stored < 0 else positive
        return Reading(read, None, as_attribute, read == value and (as_attribute == attribute or value == 0), token)
    if storage.height_flag:
        stored = wrap(value, 16, False)
        read = stored & 0x7FFF
        written = rule or "auto"
        read_rule = "auto" if stored & 0x8000 else written
        return Reading(read, read_rule, None, read == value and read_rule == written, token)
    if storage.sign_rule == "hRule" and rule == "auto":
        return Reading(value, "auto", None, True, f"{token}; ignored under auto")
    if storage.negative is not None:
        stored = wrap(value, storage.bits, storage.signed)
        if stored >= 0:
            return Reading(stored, None, None, stored == value, token)
        read = None if storage.negative == "computed" else storage.negative
        return Reading(read, None, None, read == value, token)
    if storage.sign_rule is not None:
        written = rule or ("auto" if storage.sign_rule == "lineRule" else "atLeast")
        if storage.sign_rule == "lineRule" and written == "auto":
            read = wrap(value, 16, True)
            return Reading(read, "auto", None, read == value, "int16")
        stored = wrap(-value if written == "exact" else value, 16, True)
        read = abs(stored)
        read_rule = "exact" if stored < 0 else "atLeast"
        kept = read == value and (read_rule == written or (written == "auto" and read_rule == "atLeast"))
        return Reading(read, read_rule, None, kept, token)
    read = wrap(value, storage.bits, storage.signed)
    if storage.minimum is not None:
        read = max(storage.minimum, read)
    if storage.maximum is not None:
        read = min(storage.maximum, read)
    largest = (1 << storage.bits) - 1
    kept = read == value or (storage.max_written_as is not None and value == storage.max_written_as and read == largest)
    return Reading(read, None, None, kept, token)


def kept_range(storage: Storage, attribute: str, rule: str | None = None) -> tuple[int, int] | None:
    """The values Word keeps as written: the objects package's ``rangeOf`` (core-ts's ``keptRange``).

    A signed field ``-2**(bits-1)`` to ``2**(bits-1) - 1``, an unsigned one ``0`` to
    ``2**bits - 1``; ``firstLine`` 0 to 32,767 and ``hanging`` 0 to 32,768; a
    line or a row height under ``exact`` 1 to 32,768, under ``atLeast`` 0 to
    32,767, a line under ``auto`` the plain signed field, a row height under
    ``auto`` None (ignored); a frame's height 0 to 32,767; then the clamps, and
    a column's width or space 0 and up. 65,535 where Word's save writes it -1
    is in the range: Word keeps it.
    """
    half = 1 << (storage.bits - 1)
    if storage.height_flag:
        low, high = 0, half - 1
    elif storage.pair is not None:
        low, high = (0, half) if attribute == storage.pair[1] else (0, half - 1)
    elif storage.sign_rule is not None:
        read = rule or ("atLeast" if storage.sign_rule == "hRule" else "auto")
        if storage.sign_rule == "hRule" and read == "auto":
            return None
        if read == "exact":
            low, high = 1, half
        elif read == "atLeast":
            low, high = 0, half - 1
        else:
            low, high = -half, half - 1
    elif storage.signed:
        low, high = -half, half - 1
    else:
        low, high = 0, (1 << storage.bits) - 1
    if storage.negative is not None:
        low = max(low, 0)
    if storage.minimum is not None:
        low = max(low, storage.minimum)
    if storage.maximum is not None:
        high = min(high, storage.maximum)
    return low, high


def read_range(row: Any) -> tuple[int, int]:
    """A limits-table row's range as Word **reads**: its own, with the values ``word_writes`` names put back.

    CR-022's rows stop at 65,534 where Word's save spells 65,535 as ``"-1"``
    (the write side's choice, its section 10 decision 3); Word keeps 65,535
    there, so the read side's range runs to it (CR-020 section 4.7).
    """
    high = max([row.maximum, *(stored for stored, _spelling in row.word_writes)])
    return row.minimum, high
