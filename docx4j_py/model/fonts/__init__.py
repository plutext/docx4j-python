"""Fonts: which face a character is set in, and what the consumer draws it with.

CR-002 section 6.3, implemented in Phase B (section 15.3). Two halves, as in
docx4j's ``org.docx4j.fonts``:

* :class:`~docx4j_py.model.fonts.run_font_selector.RunFontSelector` decides which
  **document** font formats each character of a run: the effective ``w:rFonts``
  with its theme references resolved, ``w:cs`` and ``w:rtl`` by value, the symbol
  fonts, and the [MS-OI29500] 17.3.2.26 character-range table. Its answer is a
  list of :class:`~docx4j_py.model.fonts.run_font_selector.FontSpan`, which is
  what a renderer, a text extractor or a measurement routine needs.
* :class:`~docx4j_py.model.fonts.mapper.Mapper` maps a document font name to a
  :class:`~docx4j_py.model.fonts.physical_font.PhysicalFont`. This phase ships
  :class:`~docx4j_py.model.fonts.mapper.IdentityPlusMapper` over a
  :class:`~docx4j_py.model.fonts.registry.FontRegistry` the caller supplies;
  docx4j's ``BestMatchingMapper``, the glyph-coverage pass and font metrics need
  font files and are a later CR, but the interface is fixed here so that a
  consumer which can read them plugs its own in.

The theme a package with no theme part is read as having is
:func:`~docx4j_py.model.fonts.default_theme.default_theme_setting` /
``pkg.default_theme``, and ``create_package()`` gives a new package that theme's
part.
"""

from docx4j_py.model.fonts.data import (
    SUBSTITUTE_ROWS,
    VCL_ENTRIES,
    Substitute,
    VclEntry,
)
from docx4j_py.model.fonts.default_theme import (
    DEFAULT_THEMES,
    THEME_2007,
    THEME_2013,
    THEME_2023,
    DefaultTheme,
    default_theme_of,
    default_theme_setting,
    set_default_theme_setting,
    theme_of_setting,
)
from docx4j_py.model.fonts.fallback import (
    EMOJI_GROUP,
    SYMBOL_GROUP,
    class_from_name,
    class_of,
    coverage_group_of,
    is_condensed,
    is_east_asian_form,
    left_to_the_document_default,
    select_by_class,
    short_key,
    substitution_class,
    vcl_entry_for,
)
from docx4j_py.model.fonts.fonts_in_use import (
    FontsAndStyles,
    fonts_in_use,
    walk_stories,
)
from docx4j_py.model.fonts.language_tags import get_script_for_language_tag, to_english
from docx4j_py.model.fonts.mapper import (
    FONT_FALLBACK,
    UNKNOWN_ERROR,
    FontDecision,
    IdentityPlusMapper,
    Mapper,
    font_table_of,
    has_bold_face,
    is_east_asian_family,
    is_known_family,
    word_default_for,
)
from docx4j_py.model.fonts.physical_font import (
    NOBOLD_SUFFIX,
    PhysicalFont,
    strip_suffixes,
)
from docx4j_py.model.fonts.registry import (
    DEFAULT_FONT_REGISTRY,
    EMPTY_FONT_REGISTRY,
    FontRegistry,
    SimpleFontRegistry,
    default_font_registry,
)
from docx4j_py.model.fonts.run_font_selector import (
    BUILT_IN_DEFAULT_FONT,
    FontSpan,
    RunFontSelector,
    is_emoji,
    is_on,
    preamble_rule,
    run_locale,
    span_script,
    symbol_font_name,
)
from docx4j_py.model.fonts.scripts import script_of
from docx4j_py.model.fonts.theme_fonts import font_scheme_of, theme_font_of

__all__ = [
    "BUILT_IN_DEFAULT_FONT",
    "DEFAULT_FONT_REGISTRY",
    "DEFAULT_THEMES",
    "EMOJI_GROUP",
    "EMPTY_FONT_REGISTRY",
    "FONT_FALLBACK",
    "NOBOLD_SUFFIX",
    "SUBSTITUTE_ROWS",
    "SYMBOL_GROUP",
    "THEME_2007",
    "THEME_2013",
    "THEME_2023",
    "UNKNOWN_ERROR",
    "VCL_ENTRIES",
    "DefaultTheme",
    "FontDecision",
    "FontRegistry",
    "FontSpan",
    "FontsAndStyles",
    "IdentityPlusMapper",
    "Mapper",
    "PhysicalFont",
    "RunFontSelector",
    "SimpleFontRegistry",
    "Substitute",
    "VclEntry",
    "class_from_name",
    "class_of",
    "coverage_group_of",
    "default_font_registry",
    "default_theme_of",
    "default_theme_setting",
    "font_scheme_of",
    "font_table_of",
    "fonts_in_use",
    "get_script_for_language_tag",
    "has_bold_face",
    "is_condensed",
    "is_east_asian_family",
    "is_east_asian_form",
    "is_emoji",
    "is_known_family",
    "is_on",
    "left_to_the_document_default",
    "preamble_rule",
    "run_locale",
    "script_of",
    "select_by_class",
    "set_default_theme_setting",
    "short_key",
    "span_script",
    "strip_suffixes",
    "substitution_class",
    "symbol_font_name",
    "theme_font_of",
    "theme_of_setting",
    "to_english",
    "vcl_entry_for",
    "walk_stories",
    "word_default_for",
]
