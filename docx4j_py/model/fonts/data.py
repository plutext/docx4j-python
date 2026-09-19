"""docx4j's font tables, as the committed JSON of ``docx4j_py/resources/fonts/``.

CR-002 section 15.3. ``scripts/generate_fonts.py`` writes these from docx4j's own
resources --- ``font-substitutes.xml``, ``FontSubstitutions.xml``,
``MicrosoftFonts.xml`` and ``word-line-metrics.properties`` --- in docx4j's own
order, so that this port and docx4j read the same tables. Each file carries a
``_source`` object naming the docx4j commit it was made from.

Everything here is read once per process, on first use: the tables are 130 KB of
JSON and a caller that never asks about a font should not pay for them.
"""

from __future__ import annotations

import dataclasses
import functools
import json
from typing import Any

__all__ = [
    "EAST_ASIAN_FAMILIES",
    "LINE_METRIC_FAMILIES",
    "MS_FONTS_WITH_BOLD",
    "MS_FONT_NAMES",
    "SUBSTITUTE_ROWS",
    "VCL_ENTRIES",
    "Substitute",
    "SubstituteRow",
    "VclEntry",
    "jar_fonts",
    "theme_xml",
]


def _resource(name: str) -> Any:
    from importlib import resources

    text = (resources.files("docx4j_py.resources.fonts") / name).read_text(encoding="utf-8")
    return json.loads(text)


@functools.cache
def _substitutions() -> Any:
    return _resource("substitutions.json")


@functools.cache
def _families() -> Any:
    return _resource("families.json")


@functools.cache
def _registry() -> Any:
    return _resource("registry.json")


@dataclasses.dataclass(frozen=True, slots=True)
class Substitute:
    """One open substitute for a document font, with what docx4j measured of it."""

    font: str
    """The physical font's name, as a font registry keys it (``Carlito Regular``)."""
    quality: str
    """``metric`` (Word's advances), ``measured`` (the closest face measured against
    Word's own PDF), or ``class`` (a face of the same class only)."""
    error: str | None = None
    """What was measured, in the words of the comment it came from; None for a
    class substitute."""


@dataclasses.dataclass(frozen=True, slots=True)
class SubstituteRow:
    """A document font and its substitutes, best first."""

    document_font: str
    substitutes: tuple[Substitute, ...]


@dataclasses.dataclass(frozen=True, slots=True)
class VclEntry:
    """One ``<replace>`` of ``FontSubstitutions.xml`` (OpenOffice's VCL.xcu, transformed)."""

    subst_fonts: str
    """``;``-separated short keys, closest first."""
    font_type: str
    """``,``-separated words: Default, SansSerif, Serif, Fixed, Typewriter, CJK, Symbol, ..."""
    font_width: str
    """Normal, Condensed, Expanded ..."""


class _SubstituteRows:
    """``font-substitutes.xml``'s rows, in the file's order, built on first use."""

    __slots__ = ("_rows",)

    def __init__(self) -> None:
        self._rows: tuple[SubstituteRow, ...] | None = None

    def _build(self) -> tuple[SubstituteRow, ...]:
        if self._rows is None:
            self._rows = tuple(
                SubstituteRow(
                    row["documentFont"],
                    tuple(
                        Substitute(s["font"], s["quality"], s.get("error"))
                        for s in row["substitutes"]
                    ),
                )
                for row in _substitutions()["substituteRows"]
            )
        return self._rows

    def __iter__(self) -> Any:
        return iter(self._build())

    def __len__(self) -> int:
        return len(self._build())

    def __getitem__(self, index: int) -> SubstituteRow:
        return self._build()[index]


#: ``font-substitutes.xml``'s ``<substitutes>`` rows, in the file's order, which is
#: the order ``Mapper.add_metrically_compatible_substitutes`` takes them in.
SUBSTITUTE_ROWS = _SubstituteRows()


class _VclEntries:
    """``FontSubstitutions.xml``'s entries, keyed as that file keys them.

    The key is the font's name lower-cased with everything but letters and digits
    removed (:func:`~docx4j_py.model.fonts.fallback.short_key`).
    """

    __slots__ = ("_entries",)

    def __init__(self) -> None:
        self._entries: dict[str, VclEntry] | None = None

    def _build(self) -> dict[str, VclEntry]:
        if self._entries is None:
            self._entries = {
                key: VclEntry(value["substFonts"], value["fontType"], value["fontWidth"])
                for key, value in _substitutions()["vclEntries"].items()
            }
        return self._entries

    def get(self, key: str) -> VclEntry | None:
        """The entry of that key, or None."""
        return self._build().get(key)

    def __len__(self) -> int:
        return len(self._build())

    def __contains__(self, key: str) -> bool:
        return key in self._build()


VCL_ENTRIES = _VclEntries()


@functools.cache
def MS_FONT_NAMES() -> frozenset[str]:
    """The families ``MicrosoftFonts.xml`` lists, as it writes them.

    ``MicrosoftFontsRegistry`` keys that map by the name exactly, so the lookup is
    case-sensitive here too.
    """
    return frozenset(_families()["msFontNames"])


@functools.cache
def MS_FONTS_WITH_BOLD() -> frozenset[str]:
    """Those of :func:`MS_FONT_NAMES` whose entry has a ``<bold>`` element.

    A family in the registry without one has no bold face of its own (Franklin
    Gothic Book), which is what ``Mapper.has_bold_face`` answers on.
    """
    return frozenset(_families()["msFontsWithBold"])


@functools.cache
def LINE_METRIC_FAMILIES() -> frozenset[str]:
    """The families ``word-line-metrics.properties`` holds, lower-cased."""
    return frozenset(_families()["lineMetricFamilies"])


@functools.cache
def EAST_ASIAN_FAMILIES() -> frozenset[str]:
    """Those of :func:`LINE_METRIC_FAMILIES` the table's optional field 7 flags East Asian."""
    return frozenset(_families()["eastAsianFamilies"])


@functools.cache
def jar_fonts() -> tuple[tuple[str, str], ...]:
    """``(name, family)`` of every face docx4j's four font jars carry."""
    return tuple((entry["name"], entry["family"]) for entry in _registry()["fonts"])


@functools.cache
def theme_xml(value: str) -> str:
    """``theme-2007.xml``, ``theme-2013.xml`` or ``theme-2023.xml``, as docx4j ships it."""
    from importlib import resources

    return (resources.files("docx4j_py.resources.fonts") / f"theme-{value}.xml").read_text(
        encoding="utf-8"
    )


def source_of(table: str) -> dict[str, Any]:
    """The ``_source`` of one of the JSON tables: the docx4j commit and the resources."""
    return {
        "substitutions": _substitutions,
        "families": _families,
        "registry": _registry,
    }[table]()["_source"]
