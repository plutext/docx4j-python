"""Which Office theme a themeless package is read as having. docx4j ``docx4j.fonts.defaultTheme``.

CR-002 section 15.3. Word supplies its own default theme to a document with **no
theme part**, and which faces that names has changed twice: measured on a Word
365 rendering of a package with no theme part (docx4j CR-016 probe
``theme-fonts-no-theme-part``), the body face is **Aptos** and the heading face
**Aptos Display**, where Office 2013 to 2022 answered Calibri and Calibri Light
and Office 2007 to 2010 answered Calibri and Cambria. So this says which version
of Word a themeless document is read as having come from --- not a preference ---
and the default is the current one.

**The shape here.** docx4j has one JVM-wide property (``docx4j.fonts.defaultTheme``,
its CR-001 batch 49 item 6); this package has no properties file and a server may
hold several packages at once, so the setting lives on the package:
``pkg.default_theme = "2013"``. :func:`default_theme_setting` is the fallback for
a selector built without a package, and is what a caller changes to move every
package at once.
"""

from __future__ import annotations

import dataclasses

from docx4j_py.model.fonts.data import theme_xml

__all__ = [
    "DEFAULT_THEMES",
    "THEME_2007",
    "THEME_2013",
    "THEME_2023",
    "DefaultTheme",
    "default_theme_of",
    "default_theme_setting",
    "set_default_theme_setting",
    "theme_of_setting",
]


@dataclasses.dataclass(frozen=True, slots=True)
class DefaultTheme:
    """A default theme: the Latin faces its font scheme names, and docx4j's theme part."""

    value: str
    """``2023``, ``2013`` or ``2007``."""
    major_latin: str
    """The face its font scheme names for the major (heading) Latin slot."""
    minor_latin: str
    """The face its font scheme names for the minor (body) Latin slot."""

    @property
    def xml(self) -> str:
        """That theme's own ``theme1.xml``, as docx4j ships it."""
        return theme_xml(self.value)


#: Word 365's Office theme: Aptos Display / Aptos. The default.
THEME_2023 = DefaultTheme("2023", "Aptos Display", "Aptos")
#: Office 2013 to 2022: Calibri Light / Calibri.
THEME_2013 = DefaultTheme("2013", "Calibri Light", "Calibri")
#: Office 2007 to 2010: Cambria / Calibri; docx4j's answer up to 17.1.0.
THEME_2007 = DefaultTheme("2007", "Cambria", "Calibri")

DEFAULT_THEMES: tuple[DefaultTheme, ...] = (THEME_2023, THEME_2013, THEME_2007)

_setting: DefaultTheme = THEME_2023


def default_theme_of(value: str | DefaultTheme | None) -> DefaultTheme | None:
    """The theme a value names; None where it names none."""
    if value is None:
        return None
    if isinstance(value, DefaultTheme):
        return value
    wanted = str(value).strip().lower()
    return next((t for t in DEFAULT_THEMES if t.value.lower() == wanted), None)


def default_theme_setting() -> DefaultTheme:
    """The process-wide default: ``2023`` unless :func:`set_default_theme_setting` moved it."""
    return _setting


def set_default_theme_setting(value: str | DefaultTheme) -> DefaultTheme:
    """Move the process-wide default; an unknown value raises rather than being ignored."""
    global _setting
    theme = default_theme_of(value)
    if theme is None:
        raise ValueError(f"docx4j.fonts.defaultTheme: not one of 2023, 2013, 2007: {value!r}")
    _setting = theme
    return theme


def theme_of_setting(value: str | DefaultTheme | None) -> DefaultTheme:
    """A package's setting, falling back to the process-wide one."""
    return default_theme_of(value) or default_theme_setting()
