"""Java's ``java.text`` formats, the parts docx4j's text model reaches through. CR-011 section 6.

docx4j formats a chart's numbers through POI's ``DataFormatter`` and a
field's through its ``FormattingSwitchHelper``, both of which end in Java's
``DecimalFormat`` and ``SimpleDateFormat``. These are those two, in the US
locale, for the patterns those callers write: :class:`DecimalFormat` (prefix
and suffix, quotes, integer and fraction digits, grouping, ``%`` and ``‰``, a
negative subpattern, scientific notation; Java's rule that ``#.##`` keeps a
leading zero and ``.00`` does not, and its negative zero) and
:func:`simple_date_format` (``y M d E H h K m s S a`` and quotes, English names).
What they do not take raises :class:`UnsupportedFormat`.

Not imported by ``import docx4j_py``.
"""

from __future__ import annotations

import datetime
from decimal import ROUND_HALF_UP, Decimal

__all__ = ["DAYS", "MONTHS", "DecimalFormat", "UnsupportedFormat", "simple_date_format"]


class UnsupportedFormat(ValueError):
    """A pattern outside what these formats take."""


# -- DecimalFormat ------------------------------------------------------------------


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


class DecimalFormat:
    """A ``java.text.DecimalFormat`` pattern in the US locale.

    POI rounds half up (its Excel style, the default here); a ``DecimalFormat``
    of its own rounds half even, as docx4j's field switches do.
    """

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
        "rounding",
        "scientific",
        "suffix",
    )

    def __init__(self, pattern: str, rounding: str = ROUND_HALF_UP) -> None:
        self.rounding = rounding
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
        self.grouping = max(0, grouping_count)
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
        rounded = value.quantize(quantum, rounding=self.rounding)
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



# -- SimpleDateFormat ---------------------------------------------------------------

MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December")  # fmt: skip
DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")



def simple_date_format(pattern: str, when: datetime.datetime) -> str:
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
                out.append(MONTHS[when.month - 1])
            elif count == 3:
                out.append(MONTHS[when.month - 1][:3])
            else:
                out.append(str(when.month).rjust(count, "0"))
        elif c == "d":
            out.append(str(when.day).rjust(count, "0"))
        elif c == "E":
            day = DAYS[when.weekday()]
            out.append(day if count >= 4 else day[:3])
        elif c == "H":
            out.append(str(when.hour).rjust(count, "0"))
        elif c == "h":
            out.append(str(when.hour % 12 or 12).rjust(count, "0"))
        elif c == "K":
            out.append(str(when.hour % 12).rjust(count, "0"))
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
