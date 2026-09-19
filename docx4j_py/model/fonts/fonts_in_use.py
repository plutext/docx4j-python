"""``fonts_in_use`` and ``styles_in_use``: **a walk for names**, not a selection.

docx4j ``MainDocumentPart.fontsInUse()`` / ``getStylesInUse()`` as its CR-016
phase 4 made them. CR-002 section 15.3.

Until docx4j 17.1.1 discovery ran a ``RunFontSelector`` in a DISCOVERY mode over
every run, deciding among the run's fonts by glyph checks it could not yet answer
--- the mapper is populated from this list, so nothing was mapped --- and took
``w:ascii`` alone from the numbering levels. What discovery needs is every name
the document could ask for, so that the mapper can map it: the four slots of
every ``w:rFonts`` on the runs, the paragraph marks and the content controls,
over the body, headers, footers, footnotes, endnotes and comments; the styles in
use with what they are based on (a table style's ``w:tblStylePr`` run properties
included); the numbering levels; the document defaults; and the default font.
Each theme reference is resolved through the theme part for the document's
``w:themeFontLang``, and a CJK font name is collected by its English name.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterable
from typing import Any

from docx4j_py.model.fonts.language_tags import to_english

__all__ = ["FontsAndStyles", "fonts_in_use", "walk_stories"]


@dataclasses.dataclass(slots=True)
class FontsAndStyles:
    """What one walk found."""

    fonts: set[str]
    """The document font names, theme references resolved, CJK names in English."""
    styles: set[str]
    """The style ids used directly, not those merely based on."""


class _Collector:
    __slots__ = ("_into", "_selector")

    def __init__(self, selector: Any, into: set[str]) -> None:
        self._selector = selector
        self._into = into

    def name(self, font_name: str | None) -> None:
        if font_name is None:
            return
        name = font_name.strip()
        if not name:
            return
        # where there is an English name, no point adding the original CJK name
        self._into.add(to_english(name) or name)

    def r_fonts(self, r_fonts: Any) -> None:
        if r_fonts is None or self._selector is None:
            return
        for name in self._selector.document_fonts_of(r_fonts):
            self.name(name)


def _sdt_r_pr(sdt_pr: Any) -> Any:
    """The ``w:rPr`` a content control's ``w:sdtPr`` carries, if any."""
    from docx4j_py.wml import RPr

    for entry in sdt_pr.content or ():
        if isinstance(entry, RPr):
            return entry
    return None


def _style_by_id(styles: Any, style_id: str) -> Any:
    if styles is None:
        return None
    return next((s for s in styles.style or () if s.style_id == style_id), None)


def _linked_style_of(styles: Any, style_id: str) -> str | None:
    """The paragraph style a character style is linked to (``w:link``), by style id."""
    style = _style_by_id(styles, style_id)
    link = getattr(getattr(style, "link", None), "val", None)
    if link is None:
        return None
    linked = _style_by_id(styles, link)
    return None if linked is None else linked.style_id


def walk_stories(
    stories: Iterable[Any],
    resolver: Any,
    selector: Any,
    collect_fonts: bool = True,
) -> FontsAndStyles:
    """One traversal of every story, collecting font names and the style ids used.

    docx4j's ``FontAndStyleFinder``. The default paragraph, character and table
    styles count as used where a paragraph, run or table names none of its own.

    Args:
        stories: the body and the other story parts' content.
        resolver: the package's property resolver, for the styles part and the
            three default style ids; None where there is none.
        selector: resolves theme references; None to collect styles only.
        collect_fonts: False to skip the font half.
    """
    from docx4j_py.traversal import iter_nodes
    from docx4j_py.wml import P, R, RSym, SdtPr, Tbl

    fonts: set[str] = set()
    styles: set[str] = set()
    add = _Collector(selector, fonts)
    default_paragraph_style_used = False
    default_character_style_used = False
    default_table_style_used = False
    all_styles = None if resolver is None else getattr(resolver, "styles_tree", None)

    for story in stories:
        if story is None:
            continue
        for value in iter_nodes(story):
            if isinstance(value, P):
                p_pr = value.p_pr
                if collect_fonts and p_pr is not None and p_pr.r_pr is not None:
                    add.r_fonts(p_pr.r_pr.r_fonts)  # the paragraph mark's
                custom_p_style = False
                if p_pr is not None:
                    if getattr(p_pr.p_style, "val", None) is not None:
                        custom_p_style = True
                        styles.add(p_pr.p_style.val)
                    mark_style = getattr(
                        getattr(getattr(p_pr, "r_pr", None), "r_style", None), "val", None
                    )
                    if mark_style is not None:
                        styles.add(mark_style)
                default_paragraph_style_used = default_paragraph_style_used or not custom_p_style
            elif isinstance(value, R):
                r_pr = value.r_pr
                if collect_fonts and r_pr is not None:
                    add.r_fonts(r_pr.r_fonts)
                if r_pr is not None:
                    r_style = getattr(r_pr.r_style, "val", None)
                    if r_style is None:
                        default_character_style_used = True
                    else:
                        styles.add(r_style)
            elif isinstance(value, Tbl):
                tbl_style = getattr(
                    getattr(getattr(value, "tbl_pr", None), "tbl_style", None), "val", None
                )
                if tbl_style is not None:
                    styles.add(tbl_style)
                default_table_style_used = default_table_style_used or tbl_style is None
            elif isinstance(value, RSym):
                if collect_fonts:
                    add.name(value.font)
            elif isinstance(value, SdtPr):
                r_pr = _sdt_r_pr(value)
                if r_pr is None:
                    continue
                if collect_fonts:
                    add.r_fonts(r_pr.r_fonts)
                r_style = getattr(r_pr.r_style, "val", None)
                if r_style is not None:
                    styles.add(r_style)
                    linked = _linked_style_of(all_styles, r_style)
                    if linked is not None:
                        styles.add(linked)

    if resolver is not None:
        if default_paragraph_style_used and resolver.default_paragraph_style_id is not None:
            styles.add(resolver.default_paragraph_style_id)
            style = _style_by_id(all_styles, resolver.default_paragraph_style_id)
            mark_style = getattr(
                getattr(getattr(style, "r_pr", None), "r_style", None), "val", None
            )
            if mark_style is not None:
                styles.add(mark_style)
        character_default = resolver.default_character_style_id
        if default_character_style_used and character_default is not None:
            styles.add(character_default)
        table_default = resolver.default_table_style_id
        if default_table_style_used and table_default is not None:
            styles.add(table_default)
    return FontsAndStyles(fonts, styles)


def fonts_in_use(
    stories: Iterable[Any], resolver: Any, numbering: Any, selector: Any
) -> set[str]:
    """The font names a document uses.

    The walk above, then the styles in use with their ``w:basedOn`` chains, the
    numbering levels, the document defaults, and the default font.

    Args:
        stories: the body and the other story parts' content.
        resolver: the package's property resolver (the styles part).
        numbering: the numbering part's content, or None.
        selector: the package's selector; it resolves the theme references and
            answers the default font.
    """
    found = walk_stories(stories, resolver, selector, collect_fonts=True)
    fonts = found.fonts
    add = _Collector(selector, fonts)
    all_styles = None if resolver is None else getattr(resolver, "styles_tree", None)

    doc_defaults_r_fonts = None
    if all_styles is not None:
        doc_defaults = all_styles.doc_defaults
        r_pr_default = None if doc_defaults is None else doc_defaults.r_pr_default
        r_pr = None if r_pr_default is None else r_pr_default.r_pr
        doc_defaults_r_fonts = None if r_pr is None else r_pr.r_fonts
        add.r_fonts(doc_defaults_r_fonts)
        # the styles in use, each with the chain it is based on
        for style_id in found.styles:
            style = _style_by_id(all_styles, style_id)
            guard = 0
            while style is not None and guard < 64:
                guard += 1
                if style.r_pr is not None:
                    add.r_fonts(style.r_pr.r_fonts)
                for tbl_style_pr in style.tbl_style_pr or ():
                    if tbl_style_pr.r_pr is not None:
                        add.r_fonts(tbl_style_pr.r_pr.r_fonts)
                based_on = getattr(style.based_on, "val", None)
                style = None if based_on is None else _style_by_id(all_styles, based_on)

    # Fonts can also be used in the numbering part. For now, treat any font
    # mentioned in that part as in use (docx4j's own comment: ideally only the
    # levels actually used).
    for abstract_num in getattr(numbering, "abstract_num", None) or ():
        for lvl in abstract_num.lvl or ():
            if lvl.r_pr is not None:
                add.r_fonts(lvl.r_pr.r_fonts)

    fonts.add(selector.default_font_of(doc_defaults_r_fonts))
    return fonts
