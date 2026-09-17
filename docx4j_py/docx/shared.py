"""python-docx's ``Length`` units and ``RGBColor``, to the digit.

CR-003 section 3.6, Phase J. python-docx measures everything in **English
Metric Units** --- 914,400 to the inch --- and hands them out as a
:class:`Length`, an ``int`` subclass with conversion properties. Code in the
wild does ``Inches(2)``, ``Pt(14)``, ``font.size.pt`` and
``section.page_width.inches``, so the constants, the rounding and the
properties here are python-docx 1.2.0's, copied rather than re-derived:
``Pt(24)`` must be exactly ``304800`` and ``Length(304800).pt`` exactly
``24.0``.

The engine underneath speaks points (``paragraph.left_indent``,
``InlinePicture.width``) and twips (``w:pgSz/@w``); this module is the only
place the two meet, through :func:`points_of` and :func:`twips_of`.
"""

from __future__ import annotations

from typing import Self

__all__ = [
    "Cm",
    "Emu",
    "Inches",
    "Length",
    "Mm",
    "Pt",
    "RGBColor",
    "Twips",
    "emu_for_points",
    "emu_for_twips",
    "length_or_none",
    "points_of",
    "twips_of",
]


class Length(int):
    """A count of English Metric Units, with python-docx's conversions.

    ``Length`` is an ``int``, so arithmetic on it works and gives a plain
    ``int`` back, exactly as python-docx's does: ``Inches(1) + Inches(1)`` is
    ``1828800`` and not a ``Length``. Wrap in :class:`Emu` to get one again.
    """

    _EMUS_PER_INCH = 914400
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PT = 12700
    _EMUS_PER_TWIP = 635

    def __new__(cls, emu: int) -> Self:
        """The length as a whole number of EMU."""
        return int.__new__(cls, emu)

    @property
    def cm(self) -> float:
        """The length in centimetres."""
        return self / float(self._EMUS_PER_CM)

    @property
    def emu(self) -> int:
        """The length in EMU, which is the length."""
        return self

    @property
    def inches(self) -> float:
        """The length in inches."""
        return self / float(self._EMUS_PER_INCH)

    @property
    def mm(self) -> float:
        """The length in millimetres."""
        return self / float(self._EMUS_PER_MM)

    @property
    def pt(self) -> float:
        """The length in points."""
        return self / float(self._EMUS_PER_PT)

    @property
    def twips(self) -> int:
        """The length in twips, rounded as python-docx rounds it."""
        return round(self / float(self._EMUS_PER_TWIP))


class Inches(Length):
    """A length in inches: ``width = Inches(0.5)``."""

    def __new__(cls, inches: float) -> Self:
        """The length, truncated to whole EMU as python-docx truncates it."""
        return Length.__new__(cls, int(inches * Length._EMUS_PER_INCH))


class Cm(Length):
    """A length in centimetres: ``height = Cm(12)``."""

    def __new__(cls, cm: float) -> Self:
        """The length, truncated to whole EMU."""
        return Length.__new__(cls, int(cm * Length._EMUS_PER_CM))


class Emu(Length):
    """A length already in EMU: ``width = Emu(457200)``."""

    def __new__(cls, emu: int) -> Self:
        """The length."""
        return Length.__new__(cls, int(emu))


class Mm(Length):
    """A length in millimetres: ``width = Mm(240.5)``."""

    def __new__(cls, mm: float) -> Self:
        """The length, truncated to whole EMU."""
        return Length.__new__(cls, int(mm * Length._EMUS_PER_MM))


class Pt(Length):
    """A length in points: ``font.size = Pt(24)``."""

    def __new__(cls, points: float) -> Self:
        """The length, truncated to whole EMU."""
        return Length.__new__(cls, int(points * Length._EMUS_PER_PT))


class Twips(Length):
    """A length in twips, a twentieth of a point and 635 EMU."""

    def __new__(cls, twips: float) -> Self:
        """The length, truncated to whole EMU."""
        return Length.__new__(cls, int(twips * Length._EMUS_PER_TWIP))


class RGBColor(tuple[int, int, int]):
    """An immutable ``(r, g, b)`` triple, printed as six hex digits."""

    def __new__(cls, r: int, g: int, b: int) -> Self:
        """Three integers 0 to 255."""
        message = "RGBColor() takes three integer values 0-255"
        for value in (r, g, b):
            if not isinstance(value, int):
                raise TypeError(message)
            if value < 0 or value > 255:
                raise ValueError(message)
        return super().__new__(cls, (r, g, b))

    def __repr__(self) -> str:
        """``RGBColor(0x3c, 0x2f, 0x80)``."""
        return "RGBColor({:#04x}, {:#04x}, {:#04x})".format(*self)

    def __str__(self) -> str:
        """``'3C2F80'``: the six upper-case hex digits Word writes."""
        return "{:02X}{:02X}{:02X}".format(*self)

    @classmethod
    def from_string(cls, rgb_hex_str: str) -> RGBColor:
        """From ``'3C2F80'``, with or without a leading ``#``."""
        text = rgb_hex_str.removeprefix("#")
        return cls(int(text[:2], 16), int(text[2:4], 16), int(text[4:6], 16))


# ---------------------------------------------------------------------------
# where python-docx's EMU meets the engine's points and the schema's twips
# ---------------------------------------------------------------------------


def points_of(value: int | Length | None) -> float | None:
    """A python-docx length as the points the content API takes."""
    return None if value is None else Length(int(value)).pt


def twips_of(value: int | Length | None) -> int | None:
    """A python-docx length as the twips ``w:pgSz`` and ``w:ind`` carry."""
    return None if value is None else Length(int(value)).twips


def emu_for_points(points: float | None) -> Length | None:
    """The points the content API reports, as a :class:`Length`."""
    return None if points is None else Pt(points)


def emu_for_twips(twips: float | None) -> Length | None:
    """The twips the schema carries, as a :class:`Length`."""
    return None if twips is None else Twips(twips)


def length_or_none(value: object) -> Length | None:
    """`value` as a :class:`Length`, or None when it is None."""
    if value is None:
        return None
    return Emu(int(value))  # type: ignore[arg-type]
