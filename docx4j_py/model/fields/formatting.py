"""A field's instruction and its formatting switches. CR-011 section 6, risk 5.

docx4j's ``FldSimpleModel``, ``FormattingSwitchHelper``,
``DateFormatInferencer`` and ``NumberExtractor`` at ``198d015a3``, ported for
what ``fields: update`` computes (WD REQ-506). The switches, in docx4j's
order:

* ``\\@`` (date and time): a value in one of the formats docx4j infers
  (international day-first, ISO 8601, ``dd MMM yyyy`` ...) written in Word's
  pattern converted to Java's (``MMMM``, ``dddd``, ``h``/``H``, ``am/pm``,
  ``'text'``); a value it cannot infer is left as it is;
* ``\\#`` (number), where the value is no date: Word's ``0``, ``#``, ``x``,
  ``.``, ``,``, ``%`` and literal text as a Java ``DecimalFormat`` (rounding
  half even, as Java's does), the number drawn out of a ``DOCPROPERTY``'s
  text (``€180,000.00 EUR``); a value that is no number ``!Syntax Error``;
  a pattern docx4j refuses (``x`` before the point, a character between
  digits, Word's anomalous fraction-only results) raises
  :class:`FieldFormattingError`, and the caller keeps the value;
* ``\\*`` (general): ``Caps``, ``FirstCap``, ``Upper``, ``Lower``, ``ARABIC``,
  ``Roman``, ``roman``; any other (``MERGEFORMAT``, ``CardText``) leaves the
  value.

An empty ``\\@`` or ``\\#`` argument is Word's ``Error! Switch argument not
specified.``. Not a general number-format engine: what Java's formats here do
not take raises, and the caller keeps the value.
"""

from __future__ import annotations

import datetime
import math
import re
from decimal import ROUND_HALF_EVEN, Decimal
from typing import Any

from docx4j_py.model.java_text import (
    MONTHS,
    DecimalFormat,
    UnsupportedFormat,
    simple_date_format,
)

__all__ = [
    "FieldFormattingError",
    "FieldValueError",
    "FldSimpleModel",
    "apply_formatting_switch",
    "convert_date_pattern",
    "determine_date_format",
    "extract_number",
    "find_first_switch_value",
]

SWITCH_ARGUMENT_MISSING = "Error! Switch argument not specified."
SYNTAX_ERROR = "!Syntax Error"
MERGEFORMAT = "MERGEFORMAT"


class FieldFormattingError(Exception):
    """A switch docx4j refuses to apply. docx4j ``FieldFormattingException``."""


class FieldValueError(Exception):
    """A field with no value to give. docx4j ``FieldValueException``."""


# ---------------------------------------------------------------------------
# the instruction (docx4j FldSimpleModel)
# ---------------------------------------------------------------------------


class FldSimpleModel:
    """A field instruction: its name (upper case) and its parameters, a quoted one kept with its quotes."""

    __slots__ = ("_parameters", "name", "parameter_string")

    def __init__(self, instr: str | None) -> None:
        """The instruction, as ``w:instr`` or the ``w:instrText`` runs give it."""
        self.name: str | None = None
        self.parameter_string: str | None = None
        self._parameters: list[str] | None = None
        text = instr or ""
        start = 0
        while start < len(text) and text[start] == " ":
            start += 1
        if start < len(text):
            end = start + 1
            if text[start] != "=":  # =nn: the name is the sign
                while end < len(text) and text[end] != " ":
                    end += 1
            self.name = text[start:end].upper()
            if end < len(text):
                self.parameter_string = text[end:].strip()

    @property
    def parameters(self) -> list[str]:
        """The parameters, space apart, a quoted one whole and with its quotes (docx4j ``splitParameters``)."""
        if self._parameters is None:
            self._parameters = split_parameters(self.parameter_string)
        return self._parameters


def split_parameters(text: str | None) -> list[str]:
    """docx4j ``FldSimpleModel.splitParameters``."""
    out: list[str] = []
    if not text:
        return out
    start = -1
    quoted = False
    for i, c in enumerate(text):
        if c == '"':
            if quoted:
                out.append(text[start : i + 1])
                quoted = False
                start = -1
            else:
                start = i
                quoted = True
        elif c == " ":
            if start > -1 and not quoted:
                out.append(text[start:i])
                start = -1
        elif start == -1:
            start = i
    if start > -1:
        out.append(text[start:])
    return out


def _find_switch(switch: str, start: int, parameters: list[str]) -> int:
    for i in range(start, len(parameters)):
        if parameters[i] == switch:
            return i
    return -1


def _switch_value(position: int, parameters: list[str]) -> str | None:
    value = parameters[position] if 0 <= position < len(parameters) else None
    if value is not None and len(value) > 1 and value[0] == '"' and value[-1] == '"':
        value = value[1:-1]
    if value and value[0] == "\\":
        return None  # a switch is no switch's value
    return value


def find_first_switch_value(switch: str, parameters: list[str], ignore_mergeformat: bool) -> str | None:
    """A switch's value: None where it is absent, ``""`` where it has none. docx4j ``findFirstSwitchValue``."""
    position = _find_switch(switch, 0, parameters)
    if position < 0:
        return None
    value = _switch_value(position + 1, parameters)
    if value == MERGEFORMAT and ignore_mergeformat:
        value = None
        position = _find_switch(switch, position + 1, parameters)
        if position > -1:
            value = _switch_value(position + 1, parameters)
    return "" if value is None else value


# ---------------------------------------------------------------------------
# dates: docx4j's DateFormatInferencer (its international formats) and SimpleDateFormat's parse
# ---------------------------------------------------------------------------

_DATE_FORMAT_REGEXPS = (
    (r"^\d{8}$", "yyyyMMdd"),
    (r"^\d{1,2}-\d{1,2}-\d{4}$", "dd-MM-yyyy"),
    (r"^\d{4}-\d{1,2}-\d{1,2}$", "yyyy-MM-dd"),
    (r"^\d{4}-\d{1,2}-\d{1,2}t\d{1,2}:\d{2}$", "yyyy-MM-dd'T'HH:mm"),
    (r"^\d{4}-\d{1,2}-\d{1,2}t\d{1,2}:\d{2}:\d{2}$", "yyyy-MM-dd'T'HH:mm:ss"),
    (r"^\d{1,2}/\d{1,2}/\d{4}$", "dd/MM/yyyy"),
    (r"^\d{1,2}\s[a-z]{3}\s\d{4}$", "dd MMM yyyy"),
    (r"^\d{1,2}\s[a-z]{4,}\s\d{4}$", "dd MMMM yyyy"),
    (r"^\d{1,2}-\d{1,2}-\d{4}\s\d{1,2}:\d{2}$", "dd-MM-yyyy HH:mm"),
    (r"^\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{2}$", "yyyy-MM-dd HH:mm"),
    (r"^\d{1,2}\s[a-z]{3}\s\d{4}\s\d{1,2}:\d{2}$", "dd MMM yyyy HH:mm"),
    (r"^\d{1,2}\s[a-z]{4,}\s\d{4}\s\d{1,2}:\d{2}$", "dd MMMM yyyy HH:mm"),
    (r"^\d{1,2}-\d{1,2}-\d{4}\s\d{1,2}:\d{2}:\d{2}$", "dd-MM-yyyy HH:mm:ss"),
    (r"^\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{2}:\d{2}$", "yyyy-MM-dd HH:mm:ss"),
    (r"^\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{2}:\d{2}\.\d{1,3}$", "yyyy-MM-dd HH:mm:ss.S"),
    (r"^\d{1,2}/\d{1,2}/\d{4}\s\d{1,2}:\d{2}:\d{2}$", "dd/MM/yyyy HH:mm:ss"),
    (r"^\d{1,2}\s[a-z]{3}\s\d{4}\s\d{1,2}:\d{2}:\d{2}$", "dd MMM yyyy HH:mm:ss"),
    (r"^\d{1,2}\s[a-z]{4,}\s\d{4}\s\d{1,2}:\d{2}:\d{2}$", "dd MMMM yyyy HH:mm:ss"),
)
_INFERENCE = tuple((re.compile(r, re.ASCII), f) for r, f in _DATE_FORMAT_REGEXPS)


def determine_date_format(value: str) -> str | None:
    """The ``SimpleDateFormat`` pattern a date's text is in; None for none (docx4j's international set)."""
    lowered = value.lower()
    for regex, pattern in _INFERENCE:
        if regex.match(lowered):
            return pattern
    return None


def _parse_date(value: str, pattern: str) -> datetime.datetime:
    """``SimpleDateFormat.parse``, leniently, for the patterns :func:`determine_date_format` gives.

    A number field takes every digit there is, but one that abuts another
    (``yyyyMMdd``) takes as many as its letters; a month name, either form.
    """
    fields: dict[str, int] = {}
    p = pattern.replace("'T'", "T")
    i = j = 0
    while i < len(p):
        c = p[i]
        run = i
        while run < len(p) and p[run] == c:
            run += 1
        count = run - i
        if c == "M" and count >= 3:
            found = re.match(r"[A-Za-z]+", value[j:])
            word = found.group(0).lower() if found else ""
            month = next((k for k, name in enumerate(MONTHS, 1) if word in (name.lower(), name[:3].lower())), None)
            if month is None:
                raise ValueError(f"no month in {value!r}")
            fields["M"] = month
            j += len(word)
        elif c in "yMdHmsS":
            abutting = run < len(p) and p[run] in "yMdHmsS"
            found = re.match(rf"\d{{{count}}}" if abutting else r"\d+", value[j:])
            if not found:
                raise ValueError(f"no number in {value!r}")
            fields[c] = int(found.group(0))
            j += len(found.group(0))
        elif c.isspace():
            while j < len(value) and value[j].isspace():
                j += 1
        elif j < len(value) and value[j].lower() == c.lower():
            j += 1
        else:
            raise ValueError(f"{value!r} is not in {pattern!r}")
        i = run if c in "yMdHmsS" else i + 1
    month = fields.get("M", 1) - 1
    start = datetime.datetime(fields.get("y", 1970) + month // 12, month % 12 + 1, 1)  # noqa: DTZ001 - no time zone
    return start + datetime.timedelta(
        days=fields.get("d", 1) - 1,
        hours=fields.get("H", 0),
        minutes=fields.get("m", 0),
        seconds=fields.get("s", 0),
        milliseconds=fields.get("S", 0),
    )


def convert_date_pattern(word: str | None) -> str:
    """Word's date and time pattern as Java's. docx4j ``FormattingSwitchHelper.convertDatePattern``."""
    out: list[str] = []
    start = -1
    last = "\0"
    quoted = False
    i = 0
    text = word or ""
    while i < len(text):
        c = text[i]
        if c == "'":
            if quoted:
                out.append(text[start:i])  # the closing quote left out
                k = i + 1
                while k < len(text) and text[k] == " ":
                    out.append(" ")
                    k += 1
                out.append("'")
                i = k - 1
                quoted = False
                start = -1
            else:
                if start > -1:
                    out.append(_date_item(text[start:i]))
                quoted = True
                start = i
        elif not quoted:
            if last != c:
                if start > -1:
                    out.append(_date_item(text[start:i]))
                start = i
                last = c
            if c in "aA" and _is_am_pm(text, i):
                out.append("a")
                i += 4
                last = "\0"
                start = -1
        i += 1
    if start > -1:
        out.append(_date_item(text[start:]))
    return "".join(out)


_DATE_ITEMS = {
    "M": "M", "MM": "MM", "MMM": "MMM", "MMMM": "MMMM",
    "d": "d", "D": "d", "dd": "dd", "DD": "dd", "ddd": "EEE", "DDD": "EEE", "dddd": "EEEE", "DDDD": "EEEE",
    "yy": "yy", "YY": "yy", "yyyy": "yyyy", "YYYY": "yyyy",
    "h": "K", "H": "H", "hh": "KK", "HH": "HH",
    "m": "m", "mm": "mm", "s": "s", "ss": "ss",
}  # fmt: skip


def _date_item(item: str) -> str:
    return _DATE_ITEMS.get(item, item)


def _is_am_pm(text: str, i: int) -> bool:
    return (
        i + 4 < len(text)
        and text[i + 1] in "mM"
        and text[i + 2] == "/"
        and text[i + 3] in "pP"
        and text[i + 4] in "mM"
    )


# ---------------------------------------------------------------------------
# numbers: docx4j's NumberExtractor and Java's Double.parseDouble
# ---------------------------------------------------------------------------

_NUMBER_IN_TEXT = re.compile(r"([\+-]?(\.\d+))|([\+-]?[0](\.\d+))|([\+-]?(\d+)(\.\d*)?)")
_JAVA_DOUBLE = re.compile(r"\s*[+-]?(NaN|Infinity|((\d+\.?\d*|\.\d+)([eE][+-]?\d+)?)[fFdD]?)\s*")


def extract_number(value: str, decimal_symbol: str | None = None) -> str:
    """The first number in a text, its grouping separators taken out; the text where there is none. docx4j ``NumberExtractor``."""
    symbol = decimal_symbol or "."
    prepared = value.replace(",", "")
    if symbol != ".":
        prepared = prepared.replace(symbol, ".")
    found = _NUMBER_IN_TEXT.search(prepared)
    return found.group(0) if found else value


def _parse_double(value: str) -> float:
    """Java's ``Double.parseDouble``."""
    if not _JAVA_DOUBLE.fullmatch(value):
        raise ValueError(f"{value!r} is not a number")
    text = value.strip().rstrip("fFdD")
    return float(text.replace("Infinity", "inf"))


# ---------------------------------------------------------------------------
# the switches (docx4j FormattingSwitchHelper)
# ---------------------------------------------------------------------------


def _format_number(pattern: str, number: float) -> str:
    """Word's numeric picture as a Java ``DecimalFormat``. docx4j ``formatNumber``."""
    if not pattern:
        pattern = "#.##########"
    decimal_point = x = False
    before = after = "\0"
    non_filler = False
    rounding = 0
    before_count = 0
    out: list[str] = []
    start = -1
    quoted = False
    i = 0
    while i < len(pattern):
        c = pattern[i]
        if c == "'":
            if quoted:
                out.append(pattern[start:i])
                k = i + 1
                while k < len(pattern) and pattern[k] == " ":
                    out.append(" ")
                    k += 1
                out.append("'")
                i = k - 1
                quoted = False
                start = -1
            else:
                if start > -1:
                    out.append(pattern[start:i])
                quoted = True
                start = i
                if before != "\0" or after != "\0":
                    non_filler = True
        elif not quoted:
            if c == "x":
                if not decimal_point:
                    raise FieldFormattingError("'x' digit dropper before the decimal point")
                x = True
                rounding += 1
                if after == "\0":
                    out.append("#")
                    after = "#"
                else:
                    out.append(after)
            elif c in "0#":
                if non_filler:
                    raise FieldFormattingError("Can't format arbitrary character between [0|#]*")
                if decimal_point:
                    rounding += 1
                    if after == "\0":
                        after = c
                    out.append(after)
                else:
                    if before == "\0":
                        before = c
                    out.append(before)
                    before_count += 1
            elif c in "%‰E¤":
                out.append(f"'{c}'")
                if before != "\0" or after != "\0":
                    non_filler = True
            elif c in "+-":
                pass  # dropped
            else:
                out.append(c)
                if c == ".":
                    decimal_point = True
                elif c not in ", " and (before != "\0" or after != "\0"):
                    non_filler = True
        i += 1
    if start > -1:
        out.append(_date_item(pattern[start:]))
    if before_count < 1 and (number < 0 or number >= 1):
        # Word loses the sign, or gives the fraction alone
        raise FieldFormattingError("Refusing to replicate Word's anomalous result")
    try:
        formatter = DecimalFormat("".join(out), ROUND_HALF_EVEN)
    except UnsupportedFormat as error:
        raise FieldFormattingError(f"{error} from {pattern}") from error
    if x:
        formatter.max_frac = rounding
    return formatter.format(Decimal(repr(number)))


def _first_cap(value: str) -> str:
    if not value:
        return ""
    if len(value) == 1:
        return value.upper()
    if value.startswith('"'):
        return '"' + value[1:2].upper() + value[2:].lower()
    return value[0].upper() + value[1:].lower()


def _format_general(format_: str, value: str) -> str:
    """docx4j ``formatGeneral``: the capitalisation formats, and Arabic and Roman numerals."""
    if not format_:
        return SWITCH_ARGUMENT_MISSING
    upper = format_.upper()
    if "CAPS" in upper:
        return " ".join(_first_cap(bit) for bit in value.split(" "))
    if "FIRSTCAP" in upper:
        return _first_cap(value)
    if "UPPER" in upper:
        return value.upper()
    if "LOWER" in upper:
        return value.lower()
    numerals = {"ARABIC": "decimal", "Roman": "upperRoman", "roman": "lowerRoman"}.get(format_)
    if numerals is not None:
        from docx4j_py.model.listnumbering.formats import format_value

        try:
            return format_value(numerals, int(value))
        except ValueError:
            return "1"  # docx4j's NumberFormatter for a value that is no number
    return value


def apply_formatting_switch(package: Any, model: FldSimpleModel, value: str, lang: str | None = None) -> str:
    """A field's value with its ``\\@``, ``\\#`` and ``\\*`` switches applied. docx4j ``applyFormattingSwitch``.

    Raises:
        FieldFormattingError: a picture docx4j refuses; the caller keeps the value.
    """
    parameters = model.parameters
    date = None
    date_format = find_first_switch_value("\\@", parameters, True)
    if date_format is not None:
        if not date_format:
            return SWITCH_ARGUMENT_MISSING
        pattern = determine_date_format(value)
        if pattern is not None:
            try:
                date = _parse_date(value, pattern)
            except ValueError:
                date = None  # not a date: the switch has no effect
        if model.name == "=":
            date = datetime.datetime.now()  # noqa: DTZ005 - Word's = takes today
        if date is not None:
            value = simple_date_format(convert_date_pattern(date_format), date)
    if date is None:
        number_format = find_first_switch_value("\\#", parameters, True)
        if number_format is not None:
            if not number_format:
                return SWITCH_ARGUMENT_MISSING
            try:
                number = _number(package, model, value)
            except ValueError:
                return SYNTAX_ERROR
            value = _format_number(number_format, number)
    general = find_first_switch_value("\\*", parameters, False)
    return value if general is None else _format_general(general, value)


def _number(package: Any, model: FldSimpleModel, value: str) -> float:
    """A field's value as a number; a ``DOCPROPERTY``'s or ``MERGEFIELD``'s drawn out of its text."""
    if model.name in ("DOCPROPERTY", "MERGEFIELD"):
        value = extract_number(value, _decimal_symbol(package))
    number = _parse_double(value)
    if math.isnan(number):
        return number
    return number


def _decimal_symbol(package: Any) -> str | None:
    main = getattr(package, "main_document_part", None) if package is not None else None
    part = getattr(main, "document_settings_part", None) if main is not None else None
    if part is None:
        return None
    try:
        symbol = part.read_contents().decimal_symbol
    except Exception:  # noqa: BLE001 - no settings: the default symbol
        return None
    return getattr(symbol, "val", None) if symbol is not None else None
