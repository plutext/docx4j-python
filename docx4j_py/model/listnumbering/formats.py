"""The number formats a list label is written in. docx4j ``NumberFormatter``.

CR-002 section 6.2 named eight: ``decimal``, ``decimalZero``, ``lowerLetter``,
``upperLetter``, ``lowerRoman``, ``upperRoman``, ``bullet`` and ``none``. They
are the ones Word's own list gallery writes and the ones the twenty exotic
formats of docx4j's ``NumberFormatter`` (Chinese, Hebrew, ordinal text, the
enclosed circles) all fall back from: an unknown ``w:numFmt``, or a value the
format cannot express, gives the **decimal** label and one ``logging`` warning
per format, as docx4j's ``warnOnce`` does. Nothing raises: a document does not
stop being readable because a level counts in Thai letters.

:func:`format_value` is docx4j's ``getCurrentValueFormatted``; :func:`register`
adds a format, which is how the rest would arrive later.
"""

from __future__ import annotations

import logging
from collections.abc import Callable

__all__ = [
    "BULLET_LABEL",
    "FORMATS",
    "format_value",
    "register",
    "warned",
]

log = logging.getLogger(__name__)

#: What docx4j's formatter returns for ``w:numFmt="bullet"``. A bullet level's
#: ``w:lvlText`` is the glyph itself and carries no ``%n``, so this is only ever
#: reached by a level text that asks for a bullet level's counter.
BULLET_LABEL = "*"

_ROMAN: tuple[tuple[int, str], ...] = (
    (1000, "m"),
    (900, "cm"),
    (500, "d"),
    (400, "cd"),
    (100, "c"),
    (90, "xc"),
    (50, "l"),
    (40, "xl"),
    (10, "x"),
    (9, "ix"),
    (5, "v"),
    (4, "iv"),
    (1, "i"),
)

#: The formats warned about already: one warning per format, whatever the value
#: and however many documents (docx4j's static ``WARNED`` set).
warned: set[str] = set()


def _roman(value: int) -> str:
    """1 to 3999 as a lower-case Roman numeral; ``ValueError`` outside that."""
    if value < 1 or value >= 4000:
        raise ValueError("roman numerals run from 1 to 3999")
    out: list[str] = []
    for number, letters in _ROMAN:
        while value >= number:
            value -= number
            out.append(letters)
    return "".join(out)


def _alphabet(letters: str) -> Callable[[int], str]:
    """docx4j ``NumberFormatAlphabet``: a, b, ... z, aa, bb, ... after the end."""

    def format_it(value: int) -> str:
        if value < 1:
            raise ValueError(f"no letter for {value}")
        size = len(letters)
        letter = letters[(value - 1) % size]
        return letter * ((value - 1) // size + 1)

    return format_it


def _decimal_zero(value: int) -> str:
    return f"0{value}" if value < 10 else str(value)


#: ``w:numFmt`` -> the label for a count. docx4j's ``REGISTRY``, the eight of
#: CR-002 section 6.2.
FORMATS: dict[str, Callable[[int], str]] = {
    "decimal": str,
    "decimalHalfWidth": str,
    "decimalZero": _decimal_zero,
    "lowerLetter": _alphabet("abcdefghijklmnopqrstuvwxyz"),
    "upperLetter": _alphabet("ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
    "lowerRoman": _roman,
    "upperRoman": lambda value: _roman(value).upper(),
    "none": lambda _value: "",
    "bullet": lambda _value: BULLET_LABEL,
}


def register(num_fmt: str, formatter: Callable[[int], str]) -> None:
    """Add or replace the formatter for a ``w:numFmt``. docx4j ``register``."""
    FORMATS[num_fmt] = formatter


def format_value(num_fmt: str | None, value: int, where: str | None = None) -> str:
    """A count as its label, the decimal one where the format cannot say it.

    Args:
        num_fmt: the level's ``w:numFmt``; ``None`` means decimal.
        value: the count.
        where: ``"numId 3 ilvl 0"``, named in the one-time warning.
    """
    if num_fmt is None:
        return str(value)
    formatter = FORMATS.get(str(num_fmt))
    if formatter is None:
        _warn_once(str(num_fmt), f"no formatter for w:numFmt {num_fmt}; decimal labels used", where)
        return str(value)
    try:
        return formatter(value)
    except ValueError as error:
        _warn_once(
            str(num_fmt),
            f"w:numFmt {num_fmt} cannot express {value} ({error}); decimal label used",
            where,
        )
        return str(value)


def _warn_once(num_fmt: str, message: str, where: str | None) -> None:
    """docx4j's ``warnOnce``: the first time at WARNING, afterwards at DEBUG.

    ``logging``, never ``warnings.warn``: ``pyproject.toml`` turns a
    ``DeprecationWarning`` into an error and a document with an exotic format is
    not a defect in the caller's code.
    """
    place = "" if where is None else f" (first at {where})"
    if num_fmt not in warned:
        warned.add(num_fmt)
        log.warning("%s%s; not reported again for this format", message, place)
    else:
        log.debug("%s%s", message, place)
