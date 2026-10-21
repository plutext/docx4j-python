"""A number as Excel formats it, by its format code: a measured subset. CR-011 section 6, risk 5.

What docx4j's chart table reads through ``org.xlsx4j.org.apache.poi.ss.usermodel.DataFormatter``
(Apache POI's, in the US locale): ``formatRawCellContents(value, index, code,
date1904)``. Not a general number-format engine:

* ``General`` (and ``@``): POI's ``ExcelGeneralNumberFormat`` --- an integer as
  it is, a magnitude of 1E11 or more (or 1E-10 or less) in ``0.#####E0``, and
  any other number rounded half up to ten significant digits and written with
  at most ten decimals;
* a date or time format (POI's ``DateUtil.isADateFormat``: only date letters,
  separators and an AM/PM marker): the serial date in the 1900 (or 1904) date
  system, written by POI's conversion of the code to a ``SimpleDateFormat``
  pattern (``d-mmm``, ``yyyy-mm-dd``, ``h:mm AM/PM``, ``mmmmm``'s first letter);
* any other code with a ``0`` or ``#``: POI's cleaning of the code (colours,
  locales, ``_x`` and ``*x``, backslashes and quotes taken out) and Java's
  ``DecimalFormat`` of what is left, rounding half up: integer and fraction
  digits, grouping, ``%`` and ``‰``, a negative section, scientific notation,
  literal prefix and suffix.

One departure: POI completes an exponent's ``+`` wherever its result has an
``E``, so a suffix such as ``EUR`` reads ``E+UR``; here only a scientific
format's result is completed.

What it does not do raises :class:`UnsupportedFormat` (elapsed time ``[h]``,
fractions ``# ?/?``, a third section, engineering notation), and the chart
table then writes the number as the file holds it, as docx4j's does where its
formatter throws.
"""

from __future__ import annotations

import datetime
import math
import re
from decimal import ROUND_HALF_UP, Context, Decimal, InvalidOperation

from docx4j_py.model.java_text import (
    DecimalFormat,
    UnsupportedFormat,
    simple_date_format,
)

__all__ = ["UnsupportedFormat", "format_raw_cell_contents", "is_a_date_format"]




# -- General ----------------------------------------------------------------------

_TO_10_SF = Context(prec=10, rounding=ROUND_HALF_UP)


def _general(value: float) -> str:
    """POI's ``ExcelGeneralNumberFormat``."""
    if math.isinf(value) or math.isnan(value):
        raise UnsupportedFormat("not a finite number")
    magnitude = abs(value)
    if magnitude >= 1e11 or (0 < magnitude <= 1e-10):
        return DecimalFormat("0.#####E0").format(Decimal(repr(value)))
    if math.floor(value) == value or magnitude >= 1e10:
        return DecimalFormat("#").format(Decimal(repr(value)))
    rounded = float(_TO_10_SF.plus(Decimal(value)))
    return DecimalFormat("#.##########").format(Decimal(repr(rounded)))


# -- dates (POI's DateUtil and createDateFormat) ----------------------------------------

_DATE_PTRN1 = re.compile(r"^\[\$\-.*?\]")
_DATE_PTRN2 = re.compile(r"^\[[a-zA-Z]+\]")
_DATE_PTRN3A = re.compile(r"[yYmMdDhHsS]")
_DATE_PTRN3B = re.compile(r"^[\[\]yYmMdDhHsS\-T/,. :\"\\]+0*[ampAMP/]*$")
_DATE_PTRN4 = re.compile(r"^\[([hH]+|[mM]+|[sS]+)\]")
_INTERNAL_DATE_FORMATS = frozenset({0x0E, 0x0F, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x2D, 0x2E, 0x2F})


def is_a_date_format(index: int, code: str | None) -> bool:
    """POI ``DateUtil.isADateFormat``."""
    if index in _INTERNAL_DATE_FORMATS:
        return True
    if not code:
        return False
    out: list[str] = []
    i = 0
    while i < len(code):
        c = code[i]
        if i < len(code) - 1:
            nc = code[i + 1]
            if c == "\\" and nc in "-,. \\":
                i += 1
                continue
            if c == ";" and nc == "@":
                i += 2
                continue
        out.append(c)
        i += 1
    fs = "".join(out)
    if _DATE_PTRN4.fullmatch(fs):
        return True
    fs = _DATE_PTRN1.sub("", fs)
    fs = _DATE_PTRN2.sub("", fs)
    separator = fs.find(";")
    if 0 < separator < len(fs) - 1:
        fs = fs[:separator]
    if not _DATE_PTRN3A.search(fs):
        return False
    return _DATE_PTRN3B.fullmatch(fs) is not None


_AM_PM = re.compile(r"((A|P)[M/P]*)", re.IGNORECASE)
_DAYS_AS_TEXT = re.compile(r"([d]{3,})", re.IGNORECASE)
def _java_date_pattern(code: str) -> str:
    """POI's ``createDateFormat``: an Excel date code as a ``SimpleDateFormat`` pattern."""
    s = code
    s = s.replace("\\-", "-").replace("\\,", ",").replace("\\.", ".").replace("\\ ", " ").replace("\\/", "/")
    s = s.replace(";@", "").replace('"/"', "/").replace('""', "'").replace("\\T", "'T'")
    has_am_pm = False
    while _AM_PM.search(s):
        s = _AM_PM.sub("@", s)
        has_am_pm = True
    s = s.replace("@", "a")
    found = _DAYS_AS_TEXT.search(s)
    if found:
        s = _DAYS_AS_TEXT.sub(found.group(0).upper().replace("D", "E"), s)
    out: list[str] = []
    m_is_month = True
    ms: list[int] = []
    j = 0
    while j < len(s):
        c = s[j]
        if c == "'":
            out.append(c)
            j += 1
            while j < len(s):
                out.append(s[j])
                if s[j] == "'":
                    break
                j += 1
        elif c == "[":
            raise UnsupportedFormat(f"{code!r}: elapsed time")
        elif c in "hH":
            m_is_month = False
            out.append("h" if has_am_pm else "H")
        elif c in "mM":
            if m_is_month:
                out.append("M")
                ms.append(len(out) - 1)
            else:
                out.append("m")
        elif c in "sS":
            out.append("s")
            for index in ms:
                if out[index] == "M":
                    out[index] = "m"
            m_is_month = True
            ms.clear()
        elif c.isalpha():
            m_is_month = True
            ms.clear()
            out.append("y" if c in "yY" else "d" if c in "dD" else c)
        else:
            if c.isspace():
                ms.clear()
            out.append(c)
        j += 1
    return "".join(out)


def _excel_date(value: float, date1904: bool) -> datetime.datetime:
    """POI ``DateUtil.getJavaDate``: the serial date in the 1900 or 1904 date system."""
    whole = math.floor(value)
    millis = int((value - whole) * 86_400_000 + 0.5)
    if date1904:
        start, adjust = 1904, 1
    else:
        start, adjust = 1900, 0 if whole < 61 else -1
    # a serial date is a wall-clock date: no time zone, as POI's local calendar has none in effect
    return datetime.datetime(start, 1, 1) + datetime.timedelta(days=whole + adjust - 1, milliseconds=millis)  # noqa: DTZ001


def _date(value: float, code: str, date1904: bool) -> str:
    pattern = _java_date_pattern(code)
    first_letter = "MMMMM" in pattern
    if first_letter:
        pattern = pattern.replace("MMMMM", "\ue001MMMM\ue002")
    text = simple_date_format(pattern, _excel_date(value, date1904))
    if first_letter:
        text = re.sub("\ue001(.)[^\ue002]*\ue002", r"\1", text)
    return text


# -- numbers (POI's createNumberFormat) ---------------------------------------------------

_COLOUR = re.compile(r"\[(BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW|COLOR\s*\d)\]", re.IGNORECASE)
_LOCALE = re.compile(r"(\[\$[^-\]]*-[0-9A-Z]+\])")
_NUMBER = re.compile(r"[0#]+")


def _clean_for_number(code: str) -> str:
    """POI ``cleanFormatForNumber``: ``_x`` and ``*x`` taken out, then backslashes, quotes and an exponent's ``+``."""
    chars = list(code)
    i = 0
    while i < len(chars):
        c = chars[i]
        if c in "_*" and not (i > 0 and chars[i - 1] == "\\"):
            if i < len(chars) - 1:
                del chars[i + 1]
            del chars[i]
            continue
        i += 1
    i = 0
    while i < len(chars):
        c = chars[i]
        if c in '\\"' or (c == "+" and i > 0 and chars[i - 1] == "E"):
            del chars[i]
            continue
        i += 1
    return "".join(chars)


def _number_text(value: float) -> Decimal:
    """POI ``NumberToTextConverter.toText``'s digits: the number to Excel's fifteen significant digits."""
    try:
        return Decimal(f"{value:.15g}")
    except InvalidOperation as error:
        raise UnsupportedFormat(str(error)) from error


def format_raw_cell_contents(value: float, index: int, code: str | None, date1904: bool = False) -> str:
    """A number as Excel shows it in a format code. POI ``DataFormatter.formatRawCellContents``, a subset.

    Raises:
        UnsupportedFormat: a code outside the subset.
    """
    code = code or "General"
    date = is_a_date_format(index, code) and value >= 0
    if code.lower() == "general" or code == "@":
        return _plus(_general(float(_number_text(value))))
    cleaned = _COLOUR.sub("", code)
    for found in _LOCALE.findall(cleaned):
        symbol = found[found.index("$") + 1 : found.index("-")]
        cleaned = cleaned.replace(found, symbol)
    if not cleaned.strip() or cleaned.lower() == "general":
        return _plus(_general(float(_number_text(value))))
    if date:
        return _date(value, cleaned, date1904)
    if "#/" in cleaned or "?/" in cleaned:
        raise UnsupportedFormat(f"{code!r}: a fraction")
    if not _NUMBER.search(cleaned):
        raise UnsupportedFormat(f"{code!r}: no digits")
    try:
        number_format = DecimalFormat(_clean_for_number(cleaned))
    except UnsupportedFormat:
        if cleaned.count(";") > 1:
            return _plus(_general(float(_number_text(value))))  # Java refuses a third section; POI's General
        raise
    result = number_format.format(_number_text(value))
    return _plus(result) if number_format.scientific else result


def _plus(result: str) -> str:
    """POI completes scientific notation with the ``+`` Java leaves out.

    Only where the number is in scientific notation: POI's own test is any
    ``E`` in the result, which turns ``5.00 EUR`` into ``5.00 E+UR``.
    """
    if "E" in result and "E-" not in result:
        return result.replace("E", "E+", 1)
    return result
