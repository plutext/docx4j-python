"""The theme part's answer to a font reference. docx4j ``ThemePart.getFont(STTheme, CTLanguage)``.

CR-002 section 15.3. A ``w:rFonts`` slot may name a *theme* font
(``w:asciiTheme="minorHAnsi"``) instead of a face. The theme part's font scheme
holds a major and a minor collection, each with a Latin, an East Asian and a
complex-script face and a list of per-script faces; ``w:themeFontLang`` in the
settings part says which language the document is in, and the script that
language selects (:mod:`~docx4j_py.model.fonts.language_tags`) picks from the
list.

Split out from the part class because the selector, font discovery and the
content API all resolve references, and none of them should have to hold a
``ThemePart``: what they need is the ``a:fontScheme``.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.fonts.language_tags import get_script_for_language_tag

__all__ = ["font_scheme_of", "theme_font_of"]

_MAJOR = frozenset({"majorAscii", "majorBidi", "majorEastAsia", "majorHAnsi"})
_LATIN_REFS = frozenset({"majorAscii", "minorAscii", "majorHAnsi", "minorHAnsi"})
_BIDI_REFS = frozenset({"majorBidi", "minorBidi"})
_EAST_ASIA_REFS = frozenset({"majorEastAsia", "minorEastAsia"})


def _value(enum_or_text: Any) -> str | None:
    if enum_or_text is None:
        return None
    return str(getattr(enum_or_text, "value", enum_or_text))


def font_scheme_of(theme: Any) -> Any:
    """A theme's ``a:fontScheme``: the theme itself, its font scheme, or None."""
    if theme is None:
        return None
    if hasattr(theme, "major_font"):
        return theme
    elements = getattr(theme, "theme_elements", None)
    return None if elements is None else elements.font_scheme


def _is_major(reference: str) -> bool:
    return reference in _MAJOR


def _lang_for(theme_font_lang: Any, reference: str) -> str | None:
    """docx4j ``ThemePart.getLang``: which of ``w:themeFontLang``'s three this reads."""
    if reference in _LATIN_REFS:
        return theme_font_lang.val
    if reference in _BIDI_REFS:
        return theme_font_lang.bidi
    return theme_font_lang.east_asia


def _text_font_from_theme(scheme: Any, reference: str) -> Any:
    """The collection's own ``a:latin`` / ``a:ea`` / ``a:cs`` for this reference."""
    collection = scheme.major_font if _is_major(reference) else scheme.minor_font
    if collection is None:
        return None
    if reference in _EAST_ASIA_REFS:
        return collection.ea
    if reference in _BIDI_REFS:
        return collection.cs
    # majorAscii / minorAscii / majorHAnsi / minorHAnsi, and anything unknown
    return collection.latin


def _font_from_theme(scheme: Any, reference: str) -> str | None:
    """The typeface of a collection's own face; None where it is absent or empty.

    ``<a:ea typeface=""/>``, which the Office theme writes, is empty.
    """
    text_font = _text_font_from_theme(scheme, reference)
    typeface = None if text_font is None else text_font.typeface
    if typeface is None or not typeface.strip():
        return None
    return typeface


def theme_font_of(theme: Any, reference: Any, theme_font_lang: Any = None) -> str | None:
    """The face a theme reference names in this theme, for this ``w:themeFontLang``.

    With no ``w:themeFontLang`` (or none that applies to this reference), the
    collection's own ``a:latin`` / ``a:ea`` / ``a:cs``. With one, the language's
    script (``Jpan``, ``Hans``, ``Geor`` ...) picks from the collection's
    ``a:font`` list, falling back to the collection's own face where the language
    selects no script or the list has no entry for it.

    Args:
        theme: an ``a:theme``, an ``a:fontScheme``, or None.
        reference: the ``w:asciiTheme`` / ``w:hAnsiTheme`` / ``w:eastAsiaTheme``
            / ``w:cstheme`` value, or None.
        theme_font_lang: ``w:themeFontLang`` from the settings part, or None.

    Returns:
        The typeface, or None where the theme answers nothing.
    """
    name = _value(reference)
    if name is None:
        return None
    scheme = font_scheme_of(theme)
    if scheme is None:
        return None
    if theme_font_lang is None:
        return _font_from_theme(scheme, name)

    lang = _lang_for(theme_font_lang, name)
    if lang is None:
        return _font_from_theme(scheme, name)
    script = get_script_for_language_tag(lang)
    if script is None:
        return _font_from_theme(scheme, name)

    collection = scheme.major_font if _is_major(name) else scheme.minor_font
    entry = None
    if collection is not None:
        entry = next((f for f in collection.font or () if f.script == script), None)
    # as docx4j: only a script the list does not carry falls back; an entry with an
    # empty typeface answers the empty string, and the caller's explicit attribute
    # does not stand
    return _font_from_theme(scheme, name) if entry is None else entry.typeface
