"""``ThemeColors``: a theme's colour scheme and the colour map that leads to it. docx4j ``org.docx4j.model.color.ThemeColors``.

CR-009 Phase B, the WordprocessingML half of docx4j's at ``a9e0303ff`` (as
docx4j-core-ts ported the whole at ``76fb45d``). All a WordprocessingML theme
reference needs: the twelve scheme slots (``dk1``, ``lt1``, ``dk2``, ``lt2``,
``accent1`` to ``accent6``, ``hlink``, ``folHlink``) and the map that takes
``bg1``, ``tx1``, ``bg2``, ``tx2`` (and, in principle, the accents and
hyperlinks) to a slot. Immutable; each ``with_...`` returns a new one.

**A docx** (:meth:`ThemeColors.of`): the main document part's theme part, or,
where the package has none, the bundled theme ``pkg.default_theme`` names (the
process-wide :func:`~docx4j_py.model.fonts.default_theme_setting` where it names
none; CR-002 section 15.3, docx4j's ``docx4j.fonts.defaultTheme``). The 2007
and 2013 settings both carry Office 2013's colours, ``accent1`` ``4472C4``
(docx4j CR-034 D2). Then ``w:settings/w:clrSchemeMapping``, which applies to
WordprocessingML theme colours and DrawingML scheme colours alike (docx4j
CR-034 section 11, P2), as :meth:`ThemeColors.with_word_mapping`. Both parts are
read with :meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart.read_contents`,
so a document whose theme and settings nobody edited still saves byte for byte.

**The seam for Phase C.** docx4j reads each slot through
``DrawingMLColors.resolve(slot, null)``. That class is CR-009 Phase C's, so a
slot here is read by :func:`_slot_colour`: an ``a:srgbClr`` or an ``a:sysClr``
(Windows' default colour for the name, else its ``lastClr``) with no
transform, which is every slot of every theme in this repository's corpora
(:mod:`~docx4j_py.model.color.tables`). Any other slot --- the other four kinds,
or a transform --- is left unresolved and logged, so a reference to it resolves
to the value written beside it, as before this phase; Phase C replaces
:func:`_slot_colour` with ``DrawingMLColors.resolve``. docx4j's ``of(Part)`` for
slides, ``of(SpreadsheetMLPackage)``, ``with_mapping`` and ``with_override``
(``p:clrMap``, ``p:clrMapOvr``), ``with_scheme``, ``for_chart``,
``with_placeholder``, ``sheet_theme_color`` and the format scheme are Phase C's
too.
"""

from __future__ import annotations

import logging
import threading
from enum import Enum
from types import MappingProxyType
from typing import Any

from docx4j_py.model.color.resolved import ResolvedColor
from docx4j_py.model.color.tables import SYSTEM

__all__ = ["MAPPED", "SLOTS", "WORD_NAMES", "ThemeColors"]

logger = logging.getLogger(__name__)

#: The scheme's slots, in ``a:clrScheme``'s order.
SLOTS: tuple[str, ...] = (
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
)

#: What a colour map takes to a slot.
MAPPED: tuple[str, ...] = (
    "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
)

#: WordprocessingML's names for the scheme's slots (``w:themeColor``'s ``dark1`` ... and ``w:clrSchemeMapping``'s).
WORD_NAMES = MappingProxyType(
    {
        "dark1": "dk1",
        "light1": "lt1",
        "dark2": "dk2",
        "light2": "lt2",
        **{f"accent{i}": f"accent{i}" for i in range(1, 7)},
        "hyperlink": "hlink",
        "followedHyperlink": "folHlink",
    }
)

#: ``a:clrScheme``'s field for each slot (the model's names).
_SCHEME_FIELDS: tuple[tuple[str, str], ...] = tuple(
    (slot, "fol_hlink" if slot == "folHlink" else slot) for slot in SLOTS
)

#: ``w:clrSchemeMapping``'s field for each mapped name: ``t1`` is ``tx1`` there, ``hyperlink`` ``hlink``.
_WORD_MAPPING_FIELDS: tuple[tuple[str, str], ...] = (
    ("bg1", "bg1"),
    ("tx1", "t1"),
    ("bg2", "bg2"),
    ("tx2", "t2"),
    *((f"accent{i}", f"accent{i}") for i in range(1, 7)),
    ("hlink", "hyperlink"),
    ("folHlink", "followed_hyperlink"),
)


def _default_map() -> dict[str, str]:
    """bg1 to lt1, tx1 to dk1, bg2 to lt2, tx2 to dk2, the rest to themselves. docx4j ``defaultMap``."""
    m = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
    for name in MAPPED[4:]:
        m[name] = name
    return m


def _text(value: Any) -> str | None:
    """An enumeration member's value, a kept ``str`` itself, None for None."""
    if value is None:
        return None
    return str(value.value) if isinstance(value, Enum) else str(value)


_ENUMERATIONS: dict[str, type[Enum]] = {}


def _member(value: Any, enumeration: str) -> str | None:
    """A ``wml`` enumeration's value as Word reads it (rule 2, :func:`~docx4j_py.model.values.in_enumeration`).

    A member, or a ``str`` that is exactly a member's value (Word's
    enumerations are case-exact), is that value; anything else is None, as
    if absent, which is also docx4j's reading (JAXB unmarshals a value outside
    the enumeration as null). The enumeration is imported on first use: this
    module imports no model class at import.
    """
    from docx4j_py.model.values import in_enumeration

    cls = _ENUMERATIONS.get(enumeration)
    if cls is None:
        from docx4j_py import wml

        cls = _ENUMERATIONS[enumeration] = getattr(wml, enumeration)
    member = in_enumeration(value, cls)
    return None if member is None else str(member.value)


def _slot_colour(slot: Any) -> ResolvedColor | None:
    """A scheme slot's colour: ``a:srgbClr`` or ``a:sysClr`` with no transform; else None (Phase C's).

    The seam where Phase C's ``DrawingMLColors.resolve(slot, None)`` goes. An
    ``a:sysClr`` is Windows' default colour for its name (docx4j CR-034 section
    11: Office draws the machine's colour, not ``lastClr``), else its
    ``lastClr`` where it is three bytes.
    """
    kind = getattr(slot, "content", None)
    if kind is None:
        return None
    if getattr(kind, "content", None):
        logger.debug("A theme slot with colour transforms is not resolved before CR-009 Phase C: %r", kind)
        return None
    name = type(kind).__name__
    if name == "CTSRgbColor":
        return ResolvedColor.of_hex(_text(kind.val))
    if name == "CTSystemColor":
        found = SYSTEM.get(_text(kind.val) or "")
        if found is not None:
            return ResolvedColor.of_hex(found)
        last = kind.last_clr
        if isinstance(last, (bytes, bytearray)) and len(last) == 3:
            return ResolvedColor.rgb(last[0], last[1], last[2])
        return None
    logger.debug("A theme slot of kind %s is not resolved before CR-009 Phase C", name)
    return None


def _scheme(clr_scheme: Any) -> dict[str, ResolvedColor]:
    """The slots of an ``a:clrScheme`` that resolve. docx4j ``scheme``."""
    out: dict[str, ResolvedColor] = {}
    if clr_scheme is None:
        return out
    for slot, field in _SCHEME_FIELDS:
        value = getattr(clr_scheme, field, None)
        colour = None if value is None else _slot_colour(value)
        if colour is not None:
            out[slot] = colour
    return out


_bundled: dict[str, Any] = {}
_bundled_lock = threading.Lock()


def _bundled_theme(default_theme: Any) -> Any:
    """The ``a:theme`` of the bundled theme a setting names, parsed once per process."""
    from docx4j_py.model.fonts.default_theme import theme_of_setting

    chosen = theme_of_setting(default_theme)
    theme = _bundled.get(chosen.value)
    if theme is None:
        from docx4j_py.dml.main import Theme
        from docx4j_py.runtime import parser

        parsed = parser().from_string(chosen.xml, Theme)
        with _bundled_lock:
            theme = _bundled.setdefault(chosen.value, parsed)
    return theme


def _read_quietly(part: Any) -> Any:
    """A part's typed tree by ``read_contents()``, None where there is no part or it can't be read."""
    if part is None:
        return None
    read = getattr(part, "read_contents", None)
    if read is None:
        return None
    try:
        return read()
    except Exception as error:  # noqa: BLE001 - docx4j logs "Couldn't read ..." and carries on
        logger.warning("Couldn't read %s: %s", getattr(part, "part_name", part), error)
        return None


class ThemeColors:
    """A theme's colour scheme and the colour map that leads to it: all a colour reference needs."""

    __slots__ = ("_computed", "_map", "_slots")

    def __init__(self, slots: dict[str, ResolvedColor], mapping: dict[str, str]) -> None:
        """Hold `slots` and `mapping` (copied). Use :meth:`of` or :meth:`of_theme`."""
        self._slots = MappingProxyType(dict(slots))
        self._map = MappingProxyType(dict(mapping))
        #: :meth:`WordColors.resolve <docx4j_py.model.color.WordColors.resolve>`'s computations in this
        #: theme, by (reference, tint, shade): the theme is immutable, so each is computed once
        #: (an optimisation docx4j does not have; the answers are the same).
        self._computed: dict[tuple[Any, Any, Any], ResolvedColor | None] = {}

    # -- making one ---------------------------------------------------------

    @staticmethod
    def of_theme(theme: Any = None, default_theme: Any = None) -> ThemeColors:
        """A theme with the default colour map. docx4j ``of(Theme)``.

        Args:
            theme: an ``a:theme`` (:class:`docx4j_py.dml.main.Theme`), or None
                for the theme Office supplies where a package has none.
            default_theme: which bundled theme that is (``"2023"``, ``"2013"``,
                ``"2007"`` or a :class:`~docx4j_py.model.fonts.DefaultTheme`);
                None takes the process-wide setting.
        """
        if theme is None:
            theme = _bundled_theme(default_theme)
        elements = getattr(theme, "theme_elements", None)
        return ThemeColors(_scheme(None if elements is None else elements.clr_scheme), _default_map())

    @staticmethod
    def of(package: Any) -> ThemeColors:
        """A docx's theme colours. docx4j ``of(WordprocessingMLPackage)``.

        The main document part's theme (or the default theme
        ``package.default_theme`` names) and ``w:clrSchemeMapping`` from the
        settings part, both read without being unmarshalled.
        """
        theme = _read_quietly(getattr(package, "theme_part", None))
        if theme is not None and not hasattr(theme, "theme_elements"):
            theme = None  # a part at the theme relationship that is not an a:theme
        colours = ThemeColors.of_theme(theme, getattr(package, "default_theme", None))
        settings = _read_quietly(getattr(package, "document_settings_part", None))
        mapping = getattr(settings, "clr_scheme_mapping", None)
        return colours if mapping is None else colours.with_word_mapping(mapping)

    # -- deriving one -------------------------------------------------------

    def with_word_mapping(self, mapping: Any) -> ThemeColors:
        """This theme with WordprocessingML's ``w:clrSchemeMapping`` (whose names differ: ``dark1``, ``t1``, ...).

        docx4j ``withWordMapping``. A name the mapping leaves out, or states
        outside ``ST_ColorSchemeIndex`` (rule 2:
        :func:`~docx4j_py.model.values.in_enumeration`'s exact match), keeps
        its slot. None leaves the map.
        """
        if mapping is None:
            return self
        m = dict(self._map)
        for key, field in _WORD_MAPPING_FIELDS:
            slot = WORD_NAMES.get(_member(getattr(mapping, field, None), "STColorSchemeIndex") or "")
            if slot is not None:
                m[key] = slot
        return ThemeColors(dict(self._slots), m)

    # -- reading it -----------------------------------------------------------

    def slot(self, slot: str) -> ResolvedColor | None:
        """A scheme slot's colour (``dk1``, ``lt1``, ..., ``folHlink``), not through the map."""
        return self._slots.get(slot)

    def scheme_color(self, val: str | None) -> ResolvedColor | None:
        """A DrawingML ``schemeClr`` value's colour: ``bg1``, ``tx1``, ``bg2``, ``tx2`` (and the rest) through the map.

        ``dk1`` ... ``lt2`` directly. docx4j ``schemeColor``, less ``phClr``
        (a style reference's placeholder, Phase C's), which is None here.
        """
        if val is None:
            return None
        slot = self._map.get(val)
        return self._slots.get(val if slot is None else slot)

    def word_theme_color(self, theme_color: Any) -> ResolvedColor | None:
        """A WordprocessingML ``w:themeColor`` value's colour. docx4j ``wordThemeColor``.

        ``text1``, ``background1``, ``text2`` and ``background2`` through the
        map; ``dark1`` ... ``followedHyperlink`` directly; None for ``none``
        and for anything outside ``ST_ThemeColor``. A member or its exact value
        (rule 2: Word's enumerations are case-exact).
        """
        name = _member(theme_color, "STThemeColor")
        if name is None:
            return None
        mapped = _WORD_THEME_MAPPED.get(name)
        if mapped is not None:
            return self.scheme_color(mapped)
        slot = WORD_NAMES.get(name)
        return None if slot is None else self._slots.get(slot)

    @property
    def mapping(self) -> MappingProxyType[str, str]:
        """The colour map: ``bg1``, ``tx1``, ... to a slot. Read-only."""
        return self._map

    def __repr__(self) -> str:
        """``<ThemeColors accent1=156082 tx1=dk1 ...>``."""
        slots = " ".join(f"{name}={colour.hex()}" for name, colour in self._slots.items())
        swapped = " ".join(f"{k}={v}" for k, v in self._map.items() if _default_map().get(k) != v)
        return f"<ThemeColors {slots}{' ' + swapped if swapped else ''}>"


#: ``w:themeColor``'s four names that go through the map.
_WORD_THEME_MAPPED = MappingProxyType({"text1": "tx1", "background1": "bg1", "text2": "tx2", "background2": "bg2"})
