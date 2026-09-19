"""``FontRegistry``: the fonts a consumer has. docx4j ``org.docx4j.fonts.PhysicalFonts``.

CR-002 section 15.3. docx4j discovers the machine's fonts once, in the ``Mapper``
static initialiser, by walking the font directories and the ``fonts/`` entries of
its own jars. None of that is portable --- a server has no idea what the author
had, a browser cannot enumerate fonts, and reading a name table needs a font
library this package does not depend on --- so a ``FontRegistry`` is a
**parameter** here (CR-002 section 6.3: the interface is fixed so that consumers
can plug their own in), and ``PhysicalFonts.discover`` over installed fonts is a
later CR.

:data:`DEFAULT_FONT_REGISTRY` is the environment the parity goldens were made in
(``docx4j.fonts.discoverPhysicalFonts.enabled=false`` with the symbol, croscore,
crosextra and theme2023 jars on the classpath), so a mapping computed here
reproduces the one in a golden on any machine.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Protocol

from docx4j_py.model.fonts.physical_font import PhysicalFont, strip_suffixes

__all__ = [
    "DEFAULT_FONT_REGISTRY",
    "EMPTY_FONT_REGISTRY",
    "FontRegistry",
    "SimpleFontRegistry",
    "default_font_registry",
]


class FontRegistry(Protocol):
    """The fonts available, by name: all the mapping passes ask of a font environment.

    Lookups are case-insensitive and strip a twin suffix, as docx4j's
    ``PhysicalFonts.get`` does.
    """

    def get(self, name: str | None) -> PhysicalFont | None:
        """The font of this name, or None where the environment has none."""
        ...  # pragma: no cover

    def all(self) -> tuple[PhysicalFont, ...]:
        """Every font, for the passes that search by a short key."""
        ...  # pragma: no cover


class SimpleFontRegistry:
    """A :class:`FontRegistry` over a fixed list of faces."""

    __slots__ = ("_by_key", "_fonts")

    def __init__(self, fonts: Iterable[PhysicalFont] = ()) -> None:
        """Build the registry over `fonts`, in their order."""
        self._by_key: dict[str, PhysicalFont] = {}
        self._fonts: list[PhysicalFont] = []
        for font in fonts:
            self.add(font)

    def add(self, font: PhysicalFont) -> SimpleFontRegistry:
        """Add a face; a second face of the same key replaces it but keeps its place."""
        key = strip_suffixes(font.name).lower()
        if key not in self._by_key:
            self._fonts.append(font)
        self._by_key[key] = font
        return self

    def get(self, name: str | None) -> PhysicalFont | None:
        """The font of this name, or None; None for a slot nothing names."""
        if name is None:
            return None
        return self._by_key.get(strip_suffixes(name).lower())

    def all(self) -> tuple[PhysicalFont, ...]:
        """Every font, in the order it was added."""
        return tuple(self._fonts)

    def __len__(self) -> int:
        """How many faces the registry holds."""
        return len(self._fonts)

    def __repr__(self) -> str:
        """``<SimpleFontRegistry 43 faces>``."""
        return f"<SimpleFontRegistry {len(self._fonts)} faces>"


#: A registry with nothing in it: every document font is ``UNMAPPED``.
EMPTY_FONT_REGISTRY: FontRegistry = SimpleFontRegistry()

_default: SimpleFontRegistry | None = None


def default_font_registry() -> FontRegistry:
    """The faces docx4j's four font jars provide, and nothing else.

    Liberation's croscore metric clones, crosextra's Carlito and Caladea, the
    theme2023 stand-ins for Word 365's Aptos, and the symbol jar's DejaVu Serif
    and Noto Sans Symbols: 43 faces. It is what
    :class:`~docx4j_py.model.fonts.mapper.IdentityPlusMapper` uses where the
    caller supplies none, and the reason a mapping computed here can be compared
    with a parity golden. A consumer with real fonts passes its own.

    Built on first use, so that the JSON is read only by a caller who asks.
    """
    global _default
    if _default is None:
        from docx4j_py.model.fonts.data import jar_fonts

        _default = SimpleFontRegistry(
            PhysicalFont(name, family) for name, family in jar_fonts()
        )
    return _default


class _DefaultRegistry:
    """:func:`default_font_registry` as a value, so importing costs no JSON."""

    __slots__ = ()

    def get(self, name: str | None) -> PhysicalFont | None:
        """The font of this name in the default registry, or None."""
        return default_font_registry().get(name)

    def all(self) -> tuple[PhysicalFont, ...]:
        """Every face of the default registry."""
        return default_font_registry().all()

    def __len__(self) -> int:
        """How many faces docx4j's font jars carry."""
        return len(default_font_registry())  # type: ignore[arg-type]

    def __repr__(self) -> str:
        """``<DEFAULT_FONT_REGISTRY ...>``."""
        return f"<DEFAULT_FONT_REGISTRY {len(self)} faces>"


#: The default :class:`FontRegistry`: docx4j's four font jars' faces.
DEFAULT_FONT_REGISTRY: FontRegistry = _DefaultRegistry()
