"""The number formats a list label is written in. docx4j ``NumberFormatter``.

CR-002 section 6.2 named eight --- ``decimal``, ``decimalZero``,
``lowerLetter``, ``upperLetter``, ``lowerRoman``, ``upperRoman``, ``bullet`` and
``none``, the ones Word's own list gallery writes --- and CR-002 section 12.11
left the rest for Phase B. **All twenty-nine docx4j ships are here** (section
15.2), one for one with its ``NumberFormat*`` class and its values: the ordinal
and cardinal-text families, ``hex``, Chicago's ``* † ‡ §``, ``numberInDash``,
the full-width, Thai and Devanagari digit families, the Russian, Arabic and Thai
alphabets, ``hebrew1``, the two Chinese counting sets and the enclosed circles.

Still **fail-soft**, which is docx4j's rule and Word's behaviour: a ``w:numFmt``
with no formatter, or a value a formatter cannot express (a Roman numeral above
3999, a circled digit above 20, a letter below 1), gives the **decimal** label
and one ``logging`` warning per format. The formatters themselves are not
fail-soft --- they raise ``ValueError``, as docx4j's throw
``NumberFormatException`` --- so a caller that wants to know still can.

:func:`format_value` is docx4j's ``getCurrentValueFormatted``; :func:`register`
adds or replaces one, which is how the counting styles docx4j does not ship
(Japanese, Korean, Vietnamese and most of the ideographic ones) would arrive.
"""

from __future__ import annotations

import logging
from collections.abc import Callable

__all__ = [
    "BULLET_LABEL",
    "FORMATS",
    "format_value",
    "formatter_for",
    "register",
    "registered_formats",
    "reset_warnings",
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


# ---------------------------------------------------------------------------
# the twenty docx4j ships beside the eight (CR-002 section 15.2)
# ---------------------------------------------------------------------------


def _digits(ten: str) -> Callable[[int], str]:
    """docx4j ``NumberFormatDigits``: a decimal label in another script's digits."""
    glyphs = tuple(ten)
    if len(glyphs) != 10:
        raise ValueError("ten digits expected")

    def format_it(value: int) -> str:
        return "".join(glyphs[ord(c) - 48] if "0" <= c <= "9" else c for c in str(value))

    return format_it


def _ordinal_suffix(value: int) -> str:
    """docx4j ``NumberFormatOrdinal.suffix``: st, nd, rd, th, and 11 to 13 th."""
    absolute = abs(value)
    if 11 <= absolute % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(absolute % 10, "th")


def _ordinal(value: int) -> str:
    """docx4j ``NumberFormatOrdinal``: 1st, 2nd, 3rd, 4th, 11th, 21st, 101st, 111th."""
    return f"{value}{_ordinal_suffix(value)}"


#: docx4j ``NumberFormatCardinalText.ONES``.
_ONES = (
    "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
)  # fmt: skip
_TENS = ("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
_SCALES = ("", "Thousand", "Million", "Billion")
_ORDINAL_ONES = (
    "Zeroth", "First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh",
    "Eighth", "Ninth", "Tenth", "Eleventh", "Twelfth", "Thirteenth", "Fourteenth",
    "Fifteenth", "Sixteenth", "Seventeenth", "Eighteenth", "Nineteenth",
)  # fmt: skip


def _below_thousand(value: int) -> str:
    out = ""
    rest = value
    if rest >= 100:
        out += f"{_ONES[rest // 100]} Hundred"
        rest %= 100
        if rest > 0:
            out += " "
    if rest >= 20:
        out += _TENS[rest // 10]
        if rest % 10 > 0:
            out += f"-{_ONES[rest % 10]}"
    elif rest > 0:
        out += _ONES[rest]
    return out


def _words(value: int) -> str:
    """docx4j ``NumberFormatCardinalText.words``: every word capitalised."""
    if value < 20:
        return _ONES[value]
    groups: list[int] = []
    rest = value
    while rest > 0:
        groups.append(rest % 1000)
        rest //= 1000
    out = ""
    for index in range(len(groups) - 1, -1, -1):
        if groups[index] == 0:
            continue
        if out:
            out += " "
        out += _below_thousand(groups[index])
        if index > 0:
            out += f" {_SCALES[index]}"
    return out


def _cardinal_text(value: int) -> str:
    """docx4j ``NumberFormatCardinalText``: One, Twenty-One, One Hundred One."""
    if value < 0:
        raise ValueError(f"no words for {value}")
    return _words(value)


def _ordinal_word(word: str) -> str:
    if word in _ONES:
        return _ORDINAL_ONES[_ONES.index(word)]
    if word.endswith("y"):
        return f"{word[:-1]}ieth"  # Twenty -> Twentieth
    return f"{word}th"  # Hundred, Thousand, Million


def _ordinal_text(value: int) -> str:
    """docx4j ``NumberFormatOrdinalText``: the cardinal words, the last made ordinal."""
    if value < 0:
        raise ValueError(f"no words for {value}")
    cardinal = _words(value)
    cut = max(cardinal.rfind(" "), cardinal.rfind("-"))
    head = "" if cut < 0 else cardinal[: cut + 1]
    return head + _ordinal_word(cardinal[cut + 1 :])


#: docx4j ``NumberFormatChicago.SYMBOLS``.
_CHICAGO = ("*", "\u2020", "\u2021", "\u00a7")


def _chicago(value: int) -> str:
    """docx4j ``NumberFormatChicago``: * then the dagger, the double dagger, the section."""
    if value < 1:
        raise ValueError(f"No symbol for {value}")
    symbol = _CHICAGO[(value - 1) % len(_CHICAGO)]
    return symbol * ((value - 1) // len(_CHICAGO) + 1)


#: docx4j ``NumberFormatDecimalEnclosedCircle``'s twenty characters.
_CIRCLED = "\u2460\u2461\u2462\u2463\u2464\u2465\u2466\u2467\u2468\u2469"
_CIRCLED += "\u246a\u246b\u246c\u246d\u246e\u246f\u2470\u2471\u2472\u2473"


def _decimal_enclosed_circle(value: int) -> str:
    """docx4j ``NumberFormatDecimalEnclosedCircle``: 1 to 20, and nothing above."""
    if value <= 0 or value > 20:
        raise ValueError("Numbers must be in range 1-20")
    return _CIRCLED[value - 1]


_HEBREW_UNITS = "\u05d0\u05d1\u05d2\u05d3\u05d4\u05d5\u05d6\u05d7\u05d8"
_HEBREW_TENS = "\u05d9\u05db\u05dc\u05de\u05e0\u05e1\u05e2\u05e4\u05e6"
_HEBREW_HUNDREDS = "\u05e7\u05e8\u05e9\u05ea"


def _hebrew1(value: int) -> str:
    """docx4j ``NumberFormatHebrew1``.

    The 15 and 16 exceptions, tav repeated above 400, and no gershayim.
    """
    if value < 1:
        raise ValueError(f"No Hebrew numeral for {value}")
    out = ""
    rest = value
    while rest >= 400:
        out += "\u05ea"
        rest -= 400
    if rest >= 100:
        out += _HEBREW_HUNDREDS[rest // 100 - 1]
        rest %= 100
    if rest == 15:
        return out + "\u05d8\u05d5"
    if rest == 16:
        return out + "\u05d8\u05d6"
    if rest >= 10:
        out += _HEBREW_TENS[rest // 10 - 1]
        rest %= 10
    if rest > 0:
        out += _HEBREW_UNITS[rest - 1]
    return out


def _chinese(
    characters: tuple[str, ...], units: tuple[str, ...], big_units: tuple[str, ...]
) -> Callable[[int], str]:
    """docx4j ``NumberFormatChineseAbstract``, lower or legal (upper) characters."""

    def four_digit(number: int) -> str:
        out = ""
        rest = number
        unit_position = 0
        last_is_zero = False
        while rest > 0:
            digit = rest % 10
            if digit != 0:
                out = f"{characters[digit]}{units[unit_position]}{out}"
                last_is_zero = False
            elif not last_is_zero:
                out = f"{characters[0]}{out}"
                last_is_zero = True
            rest //= 10
            unit_position += 1
        if out.endswith(characters[0]):  # the trailing zero character
            out = out[: -len(characters[0])]
        return out

    def format_it(value: int) -> str:
        if value == 0:
            return characters[0]
        out = ""
        rest = value
        big_unit_position = 0
        last_is_zero = False
        while rest > 0:
            part = rest % 10000
            if part != 0:
                tail = big_units[big_unit_position] if big_unit_position > 0 else ""
                out = f"{four_digit(part)}{tail}{out}"
                last_is_zero = False
            elif not last_is_zero and out:
                out = f"{characters[0]}{out}"
                last_is_zero = True
            rest //= 10000
            big_unit_position += 1
        return out

    return format_it


_CHINESE_LOWER = _chinese(
    ("\u96f6", "\u4e00", "\u4e8c", "\u4e09", "\u56db", "\u4e94", "\u516d", "\u4e03",
     "\u516b", "\u4e5d"),
    ("", "\u5341", "\u767e", "\u5343"),
    ("", "\u4e07", "\u4ebf"),
)  # fmt: skip
_CHINESE_UPPER = _chinese(
    ("\u96f6", "\u58f9", "\u8d30", "\u53c1", "\u8086", "\u4f0d", "\u9646", "\u67d2",
     "\u634c", "\u7396"),
    ("", "\u62fe", "\u4f70", "\u4edf"),
    ("", "\u4e07", "\u4ebf"),
)  # fmt: skip

FORMATS.update(
    {
        "ordinal": _ordinal,
        "cardinalText": _cardinal_text,
        "ordinalText": _ordinal_text,
        "hex": lambda value: format(value & 0xFFFFFFFF, "X"),
        "chicago": _chicago,
        "numberInDash": lambda value: f"- {value} -",
        "decimalFullWidth": _digits("\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19"),
        "decimalFullWidth2": _digits("\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19"),
        "thaiNumbers": _digits("\u0e50\u0e51\u0e52\u0e53\u0e54\u0e55\u0e56\u0e57\u0e58\u0e59"),
        "hindiNumbers": _digits("\u0966\u0967\u0968\u0969\u096a\u096b\u096c\u096d\u096e\u096f"),
        "russianLower": _alphabet(
            "\u0430\u0431\u0432\u0433\u0434\u0435\u0436\u0437\u0438\u043a\u043b\u043c"
            "\u043d\u043e\u043f\u0440\u0441\u0442\u0443\u0444\u0445\u0446\u0447\u0448"
            "\u0449\u044d\u044e\u044f"
        ),
        "russianUpper": _alphabet(
            "\u0410\u0411\u0412\u0413\u0414\u0415\u0416\u0417\u0418\u041a\u041b\u041c"
            "\u041d\u041e\u041f\u0420\u0421\u0422\u0423\u0424\u0425\u0426\u0427\u0428"
            "\u0429\u042d\u042e\u042f"
        ),
        "arabicAlpha": _alphabet(
            "\u0623\u0628\u062a\u062b\u062c\u062d\u062e\u062f\u0630\u0631\u0632\u0633"
            "\u0634\u0635\u0636\u0637\u0638\u0639\u063a\u0641\u0642\u0643\u0644\u0645"
            "\u0646\u0647\u0648\u064a"
        ),
        "thaiLetters": _alphabet(
            "\u0e01\u0e02\u0e04\u0e06\u0e07\u0e08\u0e09\u0e0a\u0e0b\u0e0c\u0e0d\u0e0e"
            "\u0e0f\u0e10\u0e11\u0e12\u0e13\u0e14\u0e15\u0e16\u0e17\u0e18\u0e19\u0e1a"
            "\u0e1b\u0e1c\u0e1d\u0e1e\u0e1f\u0e20\u0e21\u0e22\u0e23\u0e25\u0e27\u0e28"
            "\u0e29\u0e2a\u0e2b\u0e2c\u0e2d\u0e2e"
        ),
        "hebrew1": _hebrew1,
        # the two counting sets are the same in Chinese; docx4j registers both
        "chineseCounting": _CHINESE_LOWER,
        "chineseCountingThousand": _CHINESE_LOWER,
        # the legal (upper) characters
        "chineseLegalSimplified": _CHINESE_UPPER,
        # the second is the same, to suit documents written in Chinese
        "decimalEnclosedCircle": _decimal_enclosed_circle,
        "decimalEnclosedCircleChinese": _decimal_enclosed_circle,
    }
)


def formatter_for(num_fmt: str | None) -> Callable[[int], str] | None:
    """The formatter registered for a ``w:numFmt``, or None. docx4j ``NumberFormatter.get``."""
    return None if num_fmt is None else FORMATS.get(str(num_fmt))


def registered_formats() -> list[str]:
    """Every ``w:numFmt`` this package can format: a test's, and a caller's, inventory."""
    return sorted(FORMATS)


def reset_warnings() -> None:
    """Forget which formats have been warned about; a test asserting on one needs it."""
    warned.clear()
