"""The colour arithmetic Word's theme tints and shades need. docx4j ``org.docx4j.model.color.ColorMath``, its HSL half.

CR-009 Phase B. Ported **expression for expression, in the same order** from
docx4j at ``a9e0303ff`` (as docx4j-core-ts did at ``76fb45d``): Python's floats
are IEEE doubles, so the same expressions in the same order give Java's bits,
and a result that lands near a unit's edge truncates as docx4j's does. Don't
simplify these algebraically. Java's ``%`` on a double is ``math.fmod`` (the
dividend's sign), followed here by the same explicit correction as the Java's,
rather than Python's floored ``%``.

Only what :class:`~docx4j_py.model.color.word_colors.WordColors` uses is here:
``rgbToHsl``, ``hslToRgb`` (the saturation unclamped) and ``clamp01``. The rest
of ``ColorMath`` --- sRGB's transfer function, ``channel`` and
``channelHalfDown``, Foley and van Dam's ``hslToRgbFoley`` for DrawingML's
luminance transforms, and Windows' integer HLS for SpreadsheetML tints --- is
CR-009 Phase C's, with ``DrawingMLColors`` and ``SpreadsheetColors``.
"""

from __future__ import annotations

import math

__all__ = ["clamp01", "hsl_to_rgb", "rgb_to_hsl"]


def clamp01(v: float) -> float:
    """``v`` clamped to 0 to 1."""
    return 0 if v < 0 else min(v, 1)


def rgb_to_hsl(r: float, g: float, b: float) -> tuple[float, float, float]:
    """``(h, s, l)``, each 0 to 1, from sRGB channels 0 to 1; a grey's hue is 0. docx4j ``rgbToHsl``."""
    mx = max(r, max(g, b))
    mn = min(r, min(g, b))
    lum = (mx + mn) / 2
    if mx == mn:
        return (0, 0, lum)
    d = mx - mn
    s = d / (mx + mn) if lum <= 0.5 else d / (2 - mx - mn)
    if mx == r:
        h = math.fmod((g - b) / d, 6)
        if h < 0:
            h += 6
    elif mx == g:
        h = (b - r) / d + 2
    else:
        h = (r - g) / d + 4
    return (h / 6, s, lum)


def hsl_to_rgb(h: float, s: float, lum: float) -> tuple[float, float, float]:
    """sRGB channels (0 to 1, clamped) from HSL. docx4j ``hslToRgb``.

    The saturation is not clamped, as Office does not clamp it (a saturation
    above 1 pushes a channel past its range, which is then clamped).
    """
    return (_hsl_channel(0, h, s, lum), _hsl_channel(8, h, s, lum), _hsl_channel(4, h, s, lum))


def _hsl_channel(n: int, h: float, s: float, lum: float) -> float:
    k = math.fmod(n + h * 12, 12)
    if k < 0:
        k += 12
    a = s * min(lum, 1 - lum)
    return clamp01(lum - a * max(-1, min(min(k - 3, 9 - k), 1)))
