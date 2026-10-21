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

__all__ = ["UnsupportedFormat", "format_raw_cell_contents", "is_a_date_format"]


class UnsupportedFormat(ValueError):
    """A format code outside this subset."""


# -- General ----------------------------------------------------------------------

_TO_10_SF = Context(prec=10, rounding=ROUND_HALF_UP)


def _general(value: float) -> str:
    """POI's ``ExcelGeneralNumberFormat``."""
    if math.isinf(value) or math.isnan(value):
        raise UnsupportedFormat("not a finite number")
    magnitude = abs(value)
    if magnitude >= 1e11 or (0 < magnitude <= 1e-10):
        return _DecimalFormat("0.#####E0").format(Decimal(repr(value)))
    if math.floor(value) == value or magnitude >= 1e10:
        return _DecimalFormat("#").format(Decimal(repr(value)))
    rounded = float(_TO_10_SF.plus(Decimal(value)))
    return _DecimalFormat("#.##########").format(Decimal(repr(rounded)))


# -- Java's DecimalFormat, the parts a chart's codes use ------------------------------


def _affix(text: str) -> tuple[str, int]:
    """A prefix or suffix: its text, quotes resolved, and its multiplier (``%`` 100, ``‰`` 1000)."""
    out: list[str] = []
    multiplier = 1
    i = 0
    while i < len(text):
        c = text[i]
        if c == "'":
            if i + 1 < len(text) and text[i + 1] == "'":
                out.append("'")
                i += 2
                continue
            end = text.find("'", i + 1)
            if end < 0:
                raise UnsupportedFormat(f"unclosed quote in {text!r}")
            out.append(text[i + 1 : end])
            i = end + 1
            continue
        if c == "%":
            multiplier = 100
        elif c == "‰":
            multiplier = 1000
        elif c == "¤":
            c = "$"
        out.append(c)
        i += 1
    return "".join(out), multiplier


def _split_unquoted(pattern: str, separator: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    quoted = False
    for c in pattern:
        if c == "'":
            quoted = not quoted
        if c == separator and not quoted:
            parts.append("".join(current))
            current = []
            continue
        current.append(c)
    parts.append("".join(current))
    return parts


class _DecimalFormat:
    """A ``java.text.DecimalFormat`` pattern in the US locale, rounding half up (POI's Excel style)."""

    __slots__ = (
        "decimal_always",
        "exponent_digits",
        "grouping",
        "max_frac",
        "min_frac",
        "min_int",
        "multiplier",
        "neg_prefix",
        "neg_suffix",
        "prefix",
        "scientific",
        "suffix",
    )

    def __init__(self, pattern: str) -> None:
        sections = _split_unquoted(pattern, ";")
        if len(sections) > 2:
            raise UnsupportedFormat(f"{pattern!r}: more than two sections")
        prefix, number, suffix = self._parts(sections[0])
        self.prefix, m1 = _affix(prefix)
        self.suffix, m2 = _affix(suffix)
        self.multiplier = max(m1, m2)
        self._number(number)
        if len(sections) == 2:
            neg_prefix, _number, neg_suffix = self._parts(sections[1])
            self.neg_prefix, _ = _affix(neg_prefix)
            self.neg_suffix, _ = _affix(neg_suffix)
        else:
            self.neg_prefix = "-" + self.prefix
            self.neg_suffix = self.suffix

    @staticmethod
    def _parts(section: str) -> tuple[str, str, str]:
        """A section's prefix, number and suffix, quotes respected."""
        quoted = False
        start = end = -1
        for i, c in enumerate(section):
            if c == "'":
                quoted = not quoted
                continue
            if quoted:
                continue
            if c in "#0,.E" or (c.isdigit() and start >= 0):
                if c == "E" and start < 0:
                    continue
                if start < 0:
                    start = i
                end = i + 1
            elif start >= 0:
                break
        if start < 0:
            raise UnsupportedFormat(f"{section!r}: no digits")
        return section[:start], section[start:end], section[end:]

    def _number(self, number: str) -> None:
        mantissa, e, exponent = number.partition("E")
        self.scientific = bool(e)
        if self.scientific:
            if not exponent or set(exponent) - {"0"}:
                raise UnsupportedFormat(f"{number!r}: exponent")
            self.exponent_digits = len(exponent)
        else:
            self.exponent_digits = 0
        integer, dot, fraction = mantissa.partition(".")
        digit_left = zero = digit_right = 0
        decimal_pos = -1
        grouping_count = -1
        for c in integer + (dot and "."):
            if c == "#":
                if zero == 0:
                    digit_left += 1
                else:
                    digit_right += 1
                if grouping_count >= 0:
                    grouping_count += 1
            elif c == "0":
                if digit_right > 0:
                    raise UnsupportedFormat(f"{number!r}: 0 after #")
                zero += 1
                if grouping_count >= 0:
                    grouping_count += 1
            elif c == ",":
                grouping_count = 0
            elif c == ".":
                decimal_pos = digit_left + zero + digit_right
        for c in fraction:
            if c == "#":
                if zero == 0:
                    digit_left += 1
                else:
                    digit_right += 1
            elif c == "0":
                if digit_right > 0:
                    raise UnsupportedFormat(f"{number!r}: 0 after #")
                zero += 1
            else:
                raise UnsupportedFormat(f"{number!r}: {c!r} in the fraction")
        # Java's "###.###", "###." and ".###"
        if zero == 0 and digit_left > 0 and decimal_pos >= 0:
            n = decimal_pos if decimal_pos > 0 else 1
            digit_right = digit_left - n
            digit_left = n - 1
            zero = 1
        total = digit_left + zero + digit_right
        effective = decimal_pos if decimal_pos >= 0 else total
        self.min_int = effective - digit_left
        self.max_frac = total - decimal_pos if decimal_pos >= 0 else 0
        self.min_frac = digit_left + zero - decimal_pos if decimal_pos >= 0 else 0
        self.grouping = grouping_count if grouping_count > 0 else 0
        if self.scientific and (digit_left > 0 or self.min_int < 1):
            raise UnsupportedFormat(f"{number!r}: engineering notation")
        # Java's decimalSeparatorAlwaysShown
        self.decimal_always = decimal_pos == 0 or (decimal_pos >= 0 and decimal_pos == total)

    def format(self, value: Decimal) -> str:
        negative = value.is_signed()
        value = abs(value) * self.multiplier
        if self.scientific:
            body = self._scientific(value)
        else:
            body = self._fixed(value)
        return (self.neg_prefix if negative else self.prefix) + body + (self.neg_suffix if negative else self.suffix)

    def _digits(self, value: Decimal, places: int) -> tuple[str, str]:
        quantum = Decimal(1).scaleb(-places)
        rounded = value.quantize(quantum, rounding=ROUND_HALF_UP)
        text = f"{rounded:f}"
        integer, _, fraction = text.partition(".")
        return integer.lstrip("0"), fraction

    def _fixed(self, value: Decimal) -> str:
        integer, fraction = self._digits(value, self.max_frac)
        fraction = fraction.rstrip("0")
        if len(fraction) < self.min_frac:
            fraction += "0" * (self.min_frac - len(fraction))
        if len(integer) < self.min_int:
            integer = "0" * (self.min_int - len(integer)) + integer
        if not integer and not fraction:
            integer = "0"  # Java prints a zero rather than no digits at all
        if self.grouping and integer:
            groups = []
            while len(integer) > self.grouping:
                groups.insert(0, integer[-self.grouping :])
                integer = integer[: -self.grouping]
            groups.insert(0, integer)
            integer = ",".join(groups)
        return integer + ("." + fraction if fraction or self.decimal_always else "")

    def _scientific(self, value: Decimal) -> str:
        if value == 0:
            exponent = 0
            mantissa = Decimal(0)
        else:
            exponent = value.adjusted() - (self.min_int - 1)
            mantissa = value.scaleb(-exponent)
            integer, _ = self._digits(mantissa, self.max_frac)
            if len(integer) > self.min_int:  # rounding carried a digit
                exponent += 1
                mantissa = value.scaleb(-exponent)
        integer, fraction = self._digits(mantissa, self.max_frac)
        fraction = fraction.rstrip("0")
        if len(fraction) < self.min_frac:
            fraction += "0" * (self.min_frac - len(fraction))
        integer = integer.rjust(self.min_int, "0")
        digits = str(abs(exponent)).rjust(self.exponent_digits, "0")
        return integer + ("." + fraction if fraction else "") + "E" + ("-" if exponent < 0 else "") + digits


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
_MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December")  # fmt: skip
_DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")


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
    return datetime.datetime(start, 1, 1) + datetime.timedelta(days=whole + adjust - 1, milliseconds=millis)


def _simple_date_format(pattern: str, when: datetime.datetime) -> str:
    """``java.text.SimpleDateFormat`` in the US locale, for the letters POI's conversion writes."""
    out: list[str] = []
    i = 0
    n = len(pattern)
    while i < n:
        c = pattern[i]
        if c == "'":
            end = pattern.find("'", i + 1)
            if end < 0:
                end = n
            out.append("'" if end == i + 1 else pattern[i + 1 : end])
            i = end + 1
            continue
        if not c.isalpha():
            out.append(c)
            i += 1
            continue
        j = i
        while j < n and pattern[j] == c:
            j += 1
        count = j - i
        i = j
        if c == "y":
            out.append(f"{when.year % 100:02d}" if count == 2 else str(when.year).rjust(count, "0"))
        elif c == "M":
            if count >= 4:
                out.append(_MONTHS[when.month - 1])
            elif count == 3:
                out.append(_MONTHS[when.month - 1][:3])
            else:
                out.append(str(when.month).rjust(count, "0"))
        elif c == "d":
            out.append(str(when.day).rjust(count, "0"))
        elif c == "E":
            day = _DAYS[when.weekday()]
            out.append(day if count >= 4 else day[:3])
        elif c == "H":
            out.append(str(when.hour).rjust(count, "0"))
        elif c == "h":
            out.append(str(when.hour % 12 or 12).rjust(count, "0"))
        elif c == "m":
            out.append(str(when.minute).rjust(count, "0"))
        elif c == "s":
            out.append(str(when.second).rjust(count, "0"))
        elif c == "S":
            out.append(str(when.microsecond // 1000).rjust(count, "0"))
        elif c == "a":
            out.append("AM" if when.hour < 12 else "PM")
        else:
            raise UnsupportedFormat(f"{pattern!r}: the letter {c!r}")
    return "".join(out)


def _date(value: float, code: str, date1904: bool) -> str:
    pattern = _java_date_pattern(code)
    first_letter = "MMMMM" in pattern
    if first_letter:
        pattern = pattern.replace("MMMMM", "\ue001MMMM\ue002")
    text = _simple_date_format(pattern, _excel_date(value, date1904))
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
        number_format = _DecimalFormat(_clean_for_number(cleaned))
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
