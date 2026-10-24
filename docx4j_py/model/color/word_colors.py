"""``WordColors``: WordprocessingML colours resolved as Word draws them. docx4j ``org.docx4j.model.color.WordColors``.

CR-009 Phase B (section 4), ported from docx4j at ``a9e0303ff`` expression for
expression, as docx4j-core-ts ported it at ``76fb45d``. ``w:color``, ``w:shd``,
the borders and ``w:u`` (docx4j CR-034 section 11, P1):

* **A theme reference wins** (``w:themeColor``, ``w:themeFill``) over the value
  written beside it (``w:val``, ``w:color``, ``w:fill``), which is a cache Word
  keeps: Word draws the theme colour even where the value disagrees (an older
  theme's, or anything else).
* **``w:themeTint`` and ``w:themeShade`` change the HSL luminance**: tint ``t``,
  ``L' = L·t + (1 − t)``; shade ``t``, ``L' = L·t``, ``t`` the byte over 255
  (MS-OI29500 2.1.72). Where both are present Word applies the tint. The
  channels are **truncated**.
* **The cache rule.** Word's own rounding of that arithmetic is not reproduced
  exactly (within one unit per channel in all 7,700 tints and shades of four
  corpora), so where the written value is within one unit of the computation
  --- a current cache --- Word's value is returned: exact for every attribute
  Word itself wrote against the document's theme. Two units off, the
  computation.
* **A value is read as Word 365 reads it**: :meth:`WordColors.normalise` and
  :meth:`WordColors.stated` are :func:`docx4j_py.model.values.colour` and
  :func:`~docx4j_py.model.values.states_colour`, the one colour reader
  (CR-018 Phase B), as docx4j's are its ``WordColorValues``' (since
  ``a9e0303ff`` for a fill, a pattern colour and a border's colour too). A
  value Word rejects is automatic.

Each returns a :class:`~docx4j_py.model.color.resolved.ResolvedColor`, or None
for **automatic** (no theme reference, and the value ``auto``, absent or
rejected): the caller's default.

docx4j's overloads ``of(Color)``, ``of(CTBorder)`` and ``of(U)`` are one
:meth:`WordColors.of` here, which tells the three apart by class; its
package-private ``luminance`` is public, for the corpus vectors' test, as
core-ts made it.
"""

from __future__ import annotations

import math
import re
from enum import Enum
from typing import Any

from docx4j_py.model.color.color_math import clamp01, hsl_to_rgb, rgb_to_hsl
from docx4j_py.model.color.resolved import ResolvedColor
from docx4j_py.model.values import AUTO, colour, states_colour

__all__ = ["WordColors"]

#: Java's ``Integer.parseInt(s, 16)`` over the ASCII digits: an optional sign, then hex digits (core-ts's
#: reading). Python's ``int(s, 16)`` would also take spaces, ``0x`` and underscores, which Java's does not.
_HEX_NUMBER = re.compile(r"[+-]?[0-9A-Fa-f]+")


def _hex_byte(value: Any) -> int | None:
    """``w:themeTint``/``w:themeShade``: a byte in hex, else None. docx4j ``hexByte``."""
    if value is None:
        return None
    text = str(value.value) if isinstance(value, Enum) else str(value)
    if not text or _HEX_NUMBER.fullmatch(text) is None:
        return None
    v = int(text, 16)
    return v if 0 <= v <= 255 else None


def _truncate(v: float) -> int:
    return math.floor(clamp01(v) * 255 + 1e-9)


#: A computation not yet made in a theme (None is a made one: the reference names no colour there).
_UNSEEN: Any = object()

#: How many computations a theme keeps: a document has a few dozen distinct (reference, tint, shade).
_MEMO_LIMIT = 4096

_classes: tuple[type, type, type] | None = None


def _colour_classes() -> tuple[type, type, type]:
    """``(Color, CTBorder, U)``, imported on first use: this module imports no model class at import."""
    global _classes
    if _classes is None:
        from docx4j_py.wml import Color, CTBorder, U

        _classes = (Color, CTBorder, U)
    return _classes


class WordColors:
    """WordprocessingML colours resolved as Word draws them. A namespace of static methods, as docx4j's class is."""

    __slots__ = ()

    @staticmethod
    def resolve(val: Any, theme_color: Any, theme_tint: Any, theme_shade: Any, theme: Any) -> ResolvedColor | None:
        """The colour Word draws for a WordprocessingML colour's attributes.

        Args:
            val: the written value, ``RRGGBB`` (a leading ``#`` allowed), or
                anything else, which is no colour (``"auto"``, None).
            theme_color: the theme reference (an ``STThemeColor`` member or its
                value), or None or ``none``.
            theme_tint: a byte in hex (``"99"``), or None.
            theme_shade: a byte in hex (``"BF"``), or None.
            theme: the :class:`~docx4j_py.model.color.theme_colors.ThemeColors`,
                or None.

        Returns:
            None for automatic (no theme reference and `val` no colour): the
            caller's default.
        """
        written = ResolvedColor.of_hex(val)
        name = None if theme_color is None else str(getattr(theme_color, "value", theme_color))
        if name is None or name == "none" or theme is None:
            return written
        # Word's value where it is a current cache: Word's own rounding, which the computation
        # below matches for Word 365 but not for every older version that wrote the file
        key = (name, theme_tint, theme_shade)
        computed = theme._computed.get(key, _UNSEEN)
        if computed is _UNSEEN:
            computed = WordColors._computed(theme, theme_color, theme_tint, theme_shade)
            if len(theme._computed) < _MEMO_LIMIT:
                theme._computed[key] = computed
        if computed is None:
            return written
        return written if written is not None and written.within(computed, 1) else computed

    @staticmethod
    def _computed(theme: Any, theme_color: Any, theme_tint: Any, theme_shade: Any) -> ResolvedColor | None:
        """The reference's colour in `theme`, its tint (else its shade) applied; None where it has none."""
        base = theme.word_theme_color(theme_color)
        if base is None:
            return None
        tint, shade = _hex_byte(theme_tint), _hex_byte(theme_shade)
        if tint is not None:
            return WordColors.luminance(base, tint / 255.0, True)
        if shade is not None:
            return WordColors.luminance(base, shade / 255.0, False)
        return base

    @staticmethod
    def normalise(val: Any) -> str | None:
        """A colour value as Word 365 reads it: six upper-case hex digits, or None. docx4j ``normalise``.

        ``w:color/@w:val``, ``w:u/@w:color``, ``w:shd/@w:fill`` and
        ``@w:color``, a border's ``w:color``: :func:`~docx4j_py.model.values.colour`
        (trimmed, six hex digits in either case with an optional ``#``, or a
        classic HTML name), with its ``auto`` as None. None is ``auto`` or a
        value Word rejects, which Word reads as automatic in a fill or a border
        colour, and as absent in a text or underline colour (:meth:`stated`).
        """
        read = colour(val)
        return None if read is None or read == AUTO else read

    @staticmethod
    def stated(val: Any) -> bool:
        """Whether a ``w:color/@w:val`` or ``w:u/@w:color`` value states a colour. docx4j ``stated``.

        :func:`~docx4j_py.model.values.states_colour`: one :meth:`normalise`
        accepts, or exactly ``auto`` (automatic, which overrides a style's
        colour). Word reads any other value as absent: a style's colour
        beneath still shows.
        """
        return states_colour(val)

    @staticmethod
    def of(obj: Any, theme: Any) -> ResolvedColor | None:
        """A run's colour (``w:color``), a border's (``CTBorder``) or an underline's (``w:u``); None for automatic.

        docx4j's three ``of`` overloads, told apart by class. None for None.
        """
        if obj is None:
            return None
        color, border, underline = _colour_classes()
        if isinstance(obj, color):
            value = obj.val
        elif isinstance(obj, (border, underline)):
            value = obj.color
        else:
            raise TypeError(f"WordColors.of takes a w:color, a border or a w:u, not {type(obj).__name__}")
        return WordColors.resolve(
            WordColors.normalise(value), obj.theme_color, obj.theme_tint, obj.theme_shade, theme
        )

    @staticmethod
    def fill(shd: Any, theme: Any) -> ResolvedColor | None:
        """A shading's fill (``w:fill``, ``w:themeFill``); None for automatic, which is also what Word draws for a value it rejects."""
        if shd is None:
            return None
        return WordColors.resolve(
            WordColors.normalise(shd.fill), shd.theme_fill, shd.theme_fill_tint, shd.theme_fill_shade, theme
        )

    @staticmethod
    def pattern(shd: Any, theme: Any) -> ResolvedColor | None:
        """A shading's pattern colour (``w:color``, ``w:themeColor``); None for automatic."""
        if shd is None:
            return None
        return WordColors.resolve(
            WordColors.normalise(shd.color), shd.theme_color, shd.theme_tint, shd.theme_shade, theme
        )

    @staticmethod
    def luminance(c: ResolvedColor, t: float, tint: bool) -> ResolvedColor:
        """Word's ``themeTint`` (`tint` true) or ``themeShade`` on a colour: float HSL, the channels truncated.

        That is what Word 365 draws (docx4j CR-034 section 11: the three P1
        cases with no usable written value, all exact); the values older
        versions of Word wrote are within one unit of it (the 117 distinct ones
        of four corpora, 76 exact: docx4j's javadoc says 119 and 77, which its
        own vectors and test do not bear out).
        """
        h, s, lum = rgb_to_hsl(c.red / 255.0, c.green / 255.0, c.blue / 255.0)
        target = lum * t + (1 - t) if tint else lum * t
        r, g, b = hsl_to_rgb(h, s, clamp01(target))
        return ResolvedColor.rgb(_truncate(r), _truncate(g), _truncate(b))
