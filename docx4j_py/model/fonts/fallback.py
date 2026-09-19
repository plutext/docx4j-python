"""``FontFallback``: a font's class, and a face of that class. docx4j ``org.docx4j.fonts.FontFallback``.

CR-002 section 15.3, its **class and candidate half**. What is ported: the font
*class* of a name (from ``FontSubstitutions.xml`` where it knows the font, else
from the name), the families deliberately left to the document default, and
:func:`select_by_class` --- a face of the document font's class, which is what
``Mapper.add_class_based_substitutes`` takes.

The coverage half (``selectCovering``, ``covers``, ``needsCoverage``,
``warnNoCoverage``) reads a font file's cmap and is a later CR (CR-002 section
6.3); :func:`is_symbol`, :func:`is_emoji`, :func:`is_east_asian_form` and
:func:`coverage_group_of` are here because
:func:`~docx4j_py.model.fonts.run_font_selector.span_script` cuts spans by them.
"""

from __future__ import annotations

import re
from typing import Any

from docx4j_py.model.fonts.data import VCL_ENTRIES, VclEntry
from docx4j_py.model.fonts.scripts import script_of

__all__ = [
    "EMOJI_GROUP",
    "SYMBOL_GROUP",
    "class_from_name",
    "class_of",
    "coverage_group_of",
    "is_condensed",
    "is_east_asian_form",
    "is_emoji",
    "is_symbol",
    "left_to_the_document_default",
    "select_by_class",
    "short_key",
    "substitution_class",
    "vcl_entry_for",
]

SERIF = "SERIF"
SANS = "SANS"
MONO = "MONO"
UNKNOWN = "UNKNOWN"

_LETTER_OR_DIGIT = re.compile(r"[^\w]|_", re.UNICODE)


def short_key(font_name: str) -> str:
    """The key convention ``FontSubstitutions.xml`` uses.

    The name lower-cased with everything but letters and digits removed, which is
    ``Character.isLetterOrDigit`` over the whole of Unicode, as docx4j's ``key()``
    uses.
    """
    return "".join(c.lower() for c in font_name if c.isalnum())


def vcl_entry_for(document_font_name: str | None) -> VclEntry | None:
    """The ``FontSubstitutions.xml`` entry for this name, or None.

    The table knows families, not the foundry and weight suffixes a document adds
    to them (``Calisto MT``, ``Segoe UI Light``), so trailing words are dropped
    one at a time until one matches.
    """
    if document_font_name is None:
        return None
    direct = VCL_ENTRIES.get(short_key(document_font_name))
    if direct is not None:
        return direct
    words = document_font_name.strip().split()
    for drop in range(1, len(words)):
        entry = VCL_ENTRIES.get(short_key("".join(words[: len(words) - drop])))
        if entry is not None:
            return entry
    return None


def is_condensed(document_font_name: str | None) -> bool:
    """Whether ``FontSubstitutions.xml`` calls this font condensed.

    Arial Narrow and the like: no ordinary stand-in has its widths, and docx4j's
    corpus showed the document default is the lesser evil, so these are
    deliberately left unmapped.
    """
    entry = vcl_entry_for(document_font_name)
    if entry is not None and "condensed" in entry.font_width.lower():
        return True
    name = (document_font_name or "").lower()
    return "narrow" in name or "condensed" in name


#: ``.+-(bold|italic|…)(mt|ps)?``: a PostScript name, which no system has a family
#: by, so Word does not resolve it either and falls back to the document default.
_POSTSCRIPT_NAME = re.compile(
    r"^.+-(bold|italic|bolditalic|regular|roman|light|medium|semibold|demibold|black|oblique)"
    r"(mt|ps)?$"
)


def _leave_unmapped(document_font_name: str | None) -> bool:
    """Families measured to be better off falling back to the document default.

    Because the stand-in's widths are further from theirs than the default's are.
    Kept short and evidence-based, as docx4j's is.
    """
    if document_font_name is None:
        return False
    name = document_font_name.lower()
    # Lato is a narrow humanist sans; Arimo (Arial's widths) set a Lato document a page longer
    if name.startswith("lato"):
        return True
    return _POSTSCRIPT_NAME.match(name) is not None


def left_to_the_document_default(document_font_name: str | None) -> bool:
    """Whether docx4j deliberately leaves this family to the document default.

    A condensed face, or one of the measured exceptions.
    """
    return is_condensed(document_font_name) or _leave_unmapped(document_font_name)


def _class_from_table(document_font_name: str | None) -> str | None:
    """The class ``FontSubstitutions.xml`` gives this font, or None where it does not say."""
    entry = vcl_entry_for(document_font_name)
    if entry is None or not entry.font_type:
        return None
    kind = entry.font_type.lower()
    # CJK, Symbol, Special, Decorative and Script faces have no stand-in of "their class"
    if any(
        word in kind
        for word in ("cjk", "symbol", "special", "decorative", "script", "ctl")
    ):
        return UNKNOWN
    if "fixed" in kind or "typewriter" in kind:
        return MONO
    if "sansserif" in kind:
        return SANS
    if "serif" in kind:
        return SERIF
    return None


_MONO_WORDS = (
    "cousine", "monospace", "courier", "consol", "typewriter", "terminal",
    "fixedsys", "inconsolata", "menlo",
)  # fmt: skip
# NB no bare "gothic": in a Microsoft font name it is as often Japanese (MS Gothic,
# Yu Gothic) as it is a sans.  "urw gothic" is the whole family name of the Avant
# Garde clone.
_SANS_WORDS = (
    "arimo", "carlito", "cantarell", "helvetica", "arial", "urw gothic", "selawik",
    "verdana", "tahoma", "segoe", "gadugi", "trebuchet", "calibri", "grotesk",
    "grotesque", "futura", "frutiger", "myriad", "univers", "avenir", "lato", "roboto",
    "franklin", "open sans",
)  # fmt: skip
# "p052" and "c059" are the URW Palatino and Century Schoolbook clones, "gelasio"
# Georgia's.
_SERIF_WORDS = (
    "tinos", "caladea", "charter", "times", "georgia", "gelasio", "garamond",
    "palatino", "p052", "c059", "bookman", "book antiqua", "cambria", "constantia",
    "century", "baskerville", "caslon", "utopia", "minion", "sylfaen", "roman",
)  # fmt: skip

_WORD_SPLIT = re.compile(r"[^a-z0-9]+")


def class_from_name(document_font_name: str | None, generic: bool = True) -> str:
    """The class a name says, sans before serif.

    So that ``HelveticaNeue LT 55 Roman`` is a sans.

    Args:
        document_font_name: the name.
        generic: whether a name which only ends in ``Sans`` or ``Serif`` counts.
            It does when the question is what class an installed font is
            (Liberation Sans, DejaVu Serif), but not when it is whether to stand
            in for a document font: measured, that guess was worth less than the
            document default.
    """
    if document_font_name is None:
        return UNKNOWN
    name = document_font_name.lower()
    for mono in _MONO_WORDS:
        if mono in name:
            return MONO
    # "mono" only as a word: Monotype Corsiva is a script face, not a monospace one
    if any(word == "mono" for word in _WORD_SPLIT.split(name)):
        return MONO
    for sans in _SANS_WORDS:
        if sans in name:
            return SANS
    for serif in _SERIF_WORDS:
        if serif in name:
            return SERIF
    if generic:
        if "sans" in name:
            return SANS
        if "serif" in name:
            return SERIF
    return UNKNOWN


def class_of(document_font_name: str | None) -> str:
    """The font's class: the table's word where it knows the font, else the name's."""
    from_table = _class_from_table(document_font_name)
    return from_table if from_table is not None else class_from_name(document_font_name)


def substitution_class(document_font_name: str | None) -> str:
    """The class to stand in for: the table's word, else what the name says for certain."""
    from_table = _class_from_table(document_font_name)
    return (
        from_table
        if from_table is not None
        else class_from_name(document_font_name, generic=False)
    )


_SANS_DEFAULTS = (
    "Arimo Regular", "Arimo", "Liberation Sans", "Nimbus Sans", "DejaVu Sans",
    "Noto Sans", "FreeSans",
)  # fmt: skip
_SERIF_DEFAULTS = (
    "Tinos Regular", "Tinos", "Liberation Serif", "Nimbus Roman", "DejaVu Serif",
    "Noto Serif", "FreeSerif",
)  # fmt: skip
_MONO_DEFAULTS = (
    "Cousine Regular", "Cousine", "Liberation Mono", "Nimbus Mono PS",
    "DejaVu Sans Mono", "Noto Mono", "FreeMono",
)  # fmt: skip

_DEFAULTS = {SANS: _SANS_DEFAULTS, SERIF: _SERIF_DEFAULTS, MONO: _MONO_DEFAULTS}


def _vcl_substitute_keys(document_font_name: str) -> tuple[str, ...]:
    """The families ``FontSubstitutions.xml`` offers for this font, in its order."""
    entry = vcl_entry_for(document_font_name)
    if entry is None or not entry.subst_fonts:
        return ()
    return tuple(
        key for key in (short_key(name) for name in entry.subst_fonts.split(";")) if key
    )


def select_by_class(registry: Any, document_font_name: str | None) -> Any:
    """A physical font of this document font's class, without regard to what it renders.

    Used where the characters are Latin (or Greek, or Cyrillic), which any of
    these covers. ``FontSubstitutions.xml``'s own list first --- it is ordered by
    closeness --- then the class defaults. None where the font's class is
    unknown, or it is condensed or one of the measured exceptions, or the
    registry has nothing of that class.
    """
    if document_font_name is None:
        return None
    if is_condensed(document_font_name) or _leave_unmapped(document_font_name):
        return None
    font_class = substitution_class(document_font_name)
    if font_class == UNKNOWN:
        return None

    by_key = {short_key(font.name): font for font in registry.all()}
    for candidate in _vcl_substitute_keys(document_font_name):
        font = by_key.get(candidate)
        if font is not None and class_of(font.name) == font_class:
            return font
    for candidate in _DEFAULTS.get(font_class, ()):
        font = registry.get(candidate)
        if font is not None:
            return font
    return None


# ---------------------------------------------------------------- coverage groups
#
# Used by RunFontSelector.span_script, which cuts a span where the script changes.


def is_symbol(code_point: int) -> bool:
    """U+2190 to U+2BFF, the symbol blocks.

    Arrows, Mathematical Operators, Miscellaneous Technical, Box Drawing, Block
    Elements, Geometric Shapes, Miscellaneous Symbols, Dingbats, Braille and the
    supplemental arrow and maths blocks. Every one of them is ``COMMON``, so they
    are their own group rather than shared: the space and the digits beside an
    arrow are not dragged into the symbol font with it.
    """
    return 0x2190 <= code_point <= 0x2BFF


def is_emoji(code_point: int) -> bool:
    """U+1F000 to U+1FAFF, the emoji blocks.

    ``COMMON``, and like the symbols a group of their own.
    """
    return 0x1F000 <= code_point <= 0x1FAFF


def is_east_asian_form(code_point: int) -> bool:
    """The East Asian punctuation and form blocks.

    Which are ``COMMON`` --- a fullwidth comma U+FF0C, an ideographic comma
    U+3001 --- and which a Latin face does not have.
    """
    return (
        0x2E80 <= code_point <= 0x33FF
        or 0xFE10 <= code_point <= 0xFE4F
        or 0xFF00 <= code_point <= 0xFFEF
    )


SYMBOL_GROUP = "SYMBOL"
EMOJI_GROUP = "EMOJI"


def coverage_group_of(code_point: int) -> str:
    """The key a code point groups under.

    Its Unicode script, except that the symbol and emoji blocks are their own
    groups.
    """
    if is_symbol(code_point):
        return SYMBOL_GROUP
    if is_emoji(code_point):
        return EMOJI_GROUP
    return script_of(code_point)
