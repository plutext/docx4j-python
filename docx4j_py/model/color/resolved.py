"""``ResolvedColor``: a colour as Office draws it. docx4j ``org.docx4j.model.color.ResolvedColor``.

CR-009 Phase B (section 4), the port of docx4j's CR-034 at ``a9e0303ff``, as
docx4j-core-ts ported it (``ResolvedColor.mts`` at ``76fb45d``). Red, green and
blue in sRGB (0 to 255) and an alpha in **OOXML's unit**, 0 to 100,000 with
100,000 opaque (docx4j CR-034 D5), not a colour library's. Immutable.

docx4j's getters are attributes here (``getRed()`` is ``red``), ``toHex()`` is
:meth:`hex`, ``isOpaque()`` is :attr:`is_opaque`, and its ``null`` is ``None``.
"""

from __future__ import annotations

import dataclasses
import re
from typing import Any, ClassVar

__all__ = ["OPAQUE", "ResolvedColor"]

#: OOXML's opaque alpha, 100% in thousandths of a percent.
OPAQUE = 100000

#: Six hex digits, ASCII only (docx4j-core-ts's reading of Java's ``Integer.parseInt(h, 16)`` over a
#: six-character string; Java's also takes a sign and non-ASCII digits, which no colour Office writes has).
_SIX_HEX = re.compile(r"[0-9A-Fa-f]{6}")


def _clamp(value: int, maximum: int) -> int:
    return 0 if value < 0 else min(value, maximum)


@dataclasses.dataclass(frozen=True, slots=True)
class ResolvedColor:
    """A colour as Office draws it: ``red``, ``green``, ``blue`` (0 to 255) and ``alpha`` (0 to 100,000).

    Each channel is clamped (and truncated to an ``int``, as Java's are) as
    docx4j's private constructor clamps it; equality compares the four values.
    """

    red: int
    green: int
    blue: int
    alpha: int = OPAQUE

    BLACK: ClassVar[ResolvedColor]
    WHITE: ClassVar[ResolvedColor]
    OPAQUE: ClassVar[int] = OPAQUE

    def __post_init__(self) -> None:
        """Clamp the channels to 0 to 255 and the alpha to 0 to 100,000."""
        object.__setattr__(self, "red", _clamp(int(self.red), 255))
        object.__setattr__(self, "green", _clamp(int(self.green), 255))
        object.__setattr__(self, "blue", _clamp(int(self.blue), 255))
        object.__setattr__(self, "alpha", _clamp(int(self.alpha), OPAQUE))

    @staticmethod
    def rgb(red: int, green: int, blue: int) -> ResolvedColor:
        """An opaque colour; each channel is clamped to 0 to 255."""
        return ResolvedColor(red, green, blue, OPAQUE)

    @staticmethod
    def of_hex(value: Any) -> ResolvedColor | None:
        """An opaque colour from six hex digits (``RRGGBB``, either case; a leading ``#`` allowed).

        None where it is not six hex digits (``"auto"``, ``None``, anything
        not a ``str``). docx4j ``ofHex``.
        """
        if not isinstance(value, str):
            return None
        digits = value.removeprefix("#")
        if _SIX_HEX.fullmatch(digits) is None:
            return None
        v = int(digits, 16)
        return ResolvedColor.rgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF)

    def with_alpha(self, alpha: int) -> ResolvedColor:
        """This colour with another alpha (0 to 100,000, clamped)."""
        return ResolvedColor(self.red, self.green, self.blue, alpha)

    @property
    def is_opaque(self) -> bool:
        """Whether the alpha is :data:`OPAQUE`."""
        return self.alpha == OPAQUE

    def hex(self) -> str:
        """``RRGGBB``, upper case, without the alpha. docx4j ``toHex``."""
        return f"{self.red:02X}{self.green:02X}{self.blue:02X}"

    def within(self, other: ResolvedColor | None, units: int) -> bool:
        """Whether each channel is within `units` of the other colour's (alpha not compared)."""
        return (
            other is not None
            and abs(self.red - other.red) <= units
            and abs(self.green - other.green) <= units
            and abs(self.blue - other.blue) <= units
        )

    def __str__(self) -> str:
        """``RRGGBB``, or ``RRGGBB/alpha`` where not opaque. docx4j ``toString``."""
        return self.hex() if self.is_opaque else f"{self.hex()}/{self.alpha}"


ResolvedColor.BLACK = ResolvedColor(0, 0, 0, OPAQUE)
ResolvedColor.WHITE = ResolvedColor(255, 255, 255, OPAQUE)
