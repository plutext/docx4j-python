"""The Unicode script of a code point. Java's ``Character.UnicodeScript.of(cp).name()``.

CR-002 section 15.3. docx4j's ``RunFontSelector.spanScript`` and
``FontFallback.coverageGroupOf`` key on the Java enum constant's name ---
``LATIN``, ``HAN``, ``GEORGIAN``, ``COMMON``, ``INHERITED``, ``UNKNOWN`` --- which
is the Unicode script's long alias upper-cased with its hyphens turned into
underscores.

**Departure, deliberate and recorded (CR-002 section 15.3).** Python's ``re`` has
no ``\\p{Script=...}`` and :mod:`unicodedata` exposes no script property, so this
is a **range table** rather than a lookup of the whole Unicode database: the
scripts a font selection distinguishes (every script the character-range table of
[MS-OI29500] 17.3.2.26 dispatches on, and every script docx4j's coverage pass
keys on), with the combining ranges answering ``INHERITED`` and everything else
``COMMON``. A code point of a script not listed therefore answers ``COMMON``
rather than its own script name, which makes it join the span it follows --- the
safe answer, and the same one a shared character gets. Nothing in the parity
goldens depends on it: the harness records no script per span.
"""

from __future__ import annotations

import bisect

__all__ = ["COMMON", "INHERITED", "UNKNOWN", "script_of"]

COMMON = "COMMON"
INHERITED = "INHERITED"
UNKNOWN = "UNKNOWN"

#: ``(first, last, script)`` for every script this table knows, in code-point order.
#: The ranges are the Unicode blocks the script owns, which is exact for these
#: scripts: none of them shares a block with another.
_RANGES: tuple[tuple[int, int, str], ...] = (
    (0x0041, 0x005A, "LATIN"),
    (0x0061, 0x007A, "LATIN"),
    (0x00AA, 0x00AA, "LATIN"),
    (0x00BA, 0x00BA, "LATIN"),
    (0x00C0, 0x00D6, "LATIN"),
    (0x00D8, 0x00F6, "LATIN"),
    (0x00F8, 0x02B8, "LATIN"),
    (0x0300, 0x036F, INHERITED),
    (0x0370, 0x03E1, "GREEK"),
    (0x03E2, 0x03EF, "COPTIC"),
    (0x03F0, 0x03FF, "GREEK"),
    (0x0400, 0x052F, "CYRILLIC"),
    (0x0531, 0x058F, "ARMENIAN"),
    (0x0591, 0x05F4, "HEBREW"),
    (0x0600, 0x06FF, "ARABIC"),
    (0x0700, 0x074F, "SYRIAC"),
    (0x0750, 0x077F, "ARABIC"),
    (0x0780, 0x07BF, "THAANA"),
    (0x07C0, 0x07FF, "NKO"),
    (0x0800, 0x083F, "SAMARITAN"),
    (0x0840, 0x085F, "MANDAIC"),
    (0x0870, 0x08FF, "ARABIC"),
    (0x0900, 0x097F, "DEVANAGARI"),
    (0x0980, 0x09FF, "BENGALI"),
    (0x0A00, 0x0A7F, "GURMUKHI"),
    (0x0A80, 0x0AFF, "GUJARATI"),
    (0x0B00, 0x0B7F, "ORIYA"),
    (0x0B80, 0x0BFF, "TAMIL"),
    (0x0C00, 0x0C7F, "TELUGU"),
    (0x0C80, 0x0CFF, "KANNADA"),
    (0x0D00, 0x0D7F, "MALAYALAM"),
    (0x0D80, 0x0DFF, "SINHALA"),
    (0x0E00, 0x0E7F, "THAI"),
    (0x0E80, 0x0EFF, "LAO"),
    (0x0F00, 0x0FFF, "TIBETAN"),
    (0x1000, 0x109F, "MYANMAR"),
    (0x10A0, 0x10FF, "GEORGIAN"),
    (0x1100, 0x11FF, "HANGUL"),
    (0x1200, 0x139F, "ETHIOPIC"),
    (0x13A0, 0x13FF, "CHEROKEE"),
    (0x1400, 0x167F, "CANADIAN_ABORIGINAL"),
    (0x1680, 0x169F, "OGHAM"),
    (0x16A0, 0x16FF, "RUNIC"),
    (0x1700, 0x171F, "TAGALOG"),
    (0x1720, 0x173F, "HANUNOO"),
    (0x1740, 0x175F, "BUHID"),
    (0x1760, 0x177F, "TAGBANWA"),
    (0x1780, 0x17FF, "KHMER"),
    (0x1800, 0x18AF, "MONGOLIAN"),
    (0x1900, 0x194F, "LIMBU"),
    (0x1950, 0x197F, "TAI_LE"),
    (0x1980, 0x19DF, "NEW_TAI_LUE"),
    (0x1A00, 0x1A1F, "BALINESE"),
    (0x1A20, 0x1AAF, "TAI_THAM"),
    (0x1B00, 0x1B7F, "BALINESE"),
    (0x1B80, 0x1BBF, "SUNDANESE"),
    (0x1BC0, 0x1BFF, "BATAK"),
    (0x1C00, 0x1C4F, "LEPCHA"),
    (0x1C50, 0x1C7F, "OL_CHIKI"),
    (0x1D00, 0x1D7F, "LATIN"),
    (0x1E00, 0x1EFF, "LATIN"),
    (0x1F00, 0x1FFF, "GREEK"),
    (0x2071, 0x2071, "LATIN"),
    (0x207F, 0x207F, "LATIN"),
    (0x20D0, 0x20FF, INHERITED),
    (0x2126, 0x2126, "GREEK"),
    (0x212A, 0x212B, "LATIN"),
    (0x2C60, 0x2C7F, "LATIN"),
    (0x2C80, 0x2CFF, "COPTIC"),
    (0x2D00, 0x2D2F, "GEORGIAN"),
    (0x2D30, 0x2D7F, "TIFINAGH"),
    (0x2D80, 0x2DDF, "ETHIOPIC"),
    (0x2DE0, 0x2DFF, "CYRILLIC"),
    (0x2E80, 0x2EFF, "HAN"),
    (0x3005, 0x3005, "HAN"),
    (0x3007, 0x3007, "HAN"),
    (0x3021, 0x3029, "HAN"),
    (0x3038, 0x303B, "HAN"),
    (0x3041, 0x309F, "HIRAGANA"),
    (0x30A1, 0x30FF, "KATAKANA"),
    (0x3105, 0x312F, "BOPOMOFO"),
    (0x3131, 0x318F, "HANGUL"),
    (0x31A0, 0x31BF, "BOPOMOFO"),
    (0x31F0, 0x31FF, "KATAKANA"),
    (0x3400, 0x4DBF, "HAN"),
    (0x4E00, 0x9FFF, "HAN"),
    (0xA000, 0xA4CF, "YI"),
    (0xA500, 0xA63F, "VAI"),
    (0xA640, 0xA69F, "CYRILLIC"),
    (0xA6A0, 0xA6FF, "BAMUM"),
    (0xA720, 0xA7FF, "LATIN"),
    (0xA800, 0xA82F, "SYLOTI_NAGRI"),
    (0xA840, 0xA87F, "PHAGS_PA"),
    (0xA880, 0xA8DF, "SAURASHTRA"),
    (0xA900, 0xA92F, "KAYAH_LI"),
    (0xA930, 0xA95F, "REJANG"),
    (0xA960, 0xA97F, "HANGUL"),
    (0xA980, 0xA9DF, "JAVANESE"),
    (0xAA00, 0xAA5F, "CHAM"),
    (0xAA80, 0xAADF, "TAI_VIET"),
    (0xABC0, 0xABFF, "MEETEI_MAYEK"),
    (0xAC00, 0xD7FF, "HANGUL"),
    (0xF900, 0xFAFF, "HAN"),
    (0xFB00, 0xFB06, "LATIN"),
    (0xFB13, 0xFB17, "ARMENIAN"),
    (0xFB1D, 0xFB4F, "HEBREW"),
    (0xFB50, 0xFDFF, "ARABIC"),
    (0xFE00, 0xFE0F, INHERITED),
    (0xFE20, 0xFE2F, INHERITED),
    (0xFE70, 0xFEFC, "ARABIC"),
    (0xFF21, 0xFF3A, "LATIN"),
    (0xFF41, 0xFF5A, "LATIN"),
    (0xFF66, 0xFF9D, "KATAKANA"),
    (0xFFA0, 0xFFDC, "HANGUL"),
    (0x10140, 0x1018F, "GREEK"),
    (0x1D400, 0x1D7CB, COMMON),
    (0x20000, 0x3FFFF, "HAN"),
)

_STARTS = tuple(first for first, _last, _script in _RANGES)


def script_of(code_point: int) -> str:
    """The Unicode script of a code point, as Java's ``UnicodeScript.of(cp).name()`` spells it.

    ``COMMON`` for a shared character and for a script this table does not carry
    (the departure the module docstring records).
    """
    index = bisect.bisect_right(_STARTS, code_point) - 1
    if index < 0:
        return COMMON
    first, last, script = _RANGES[index]
    return script if first <= code_point <= last else COMMON
