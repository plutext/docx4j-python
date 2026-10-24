"""``PropertyResolver``: the properties that actually apply. docx4j ``org.docx4j.model.PropertyResolver``.

CR-002 section 6.1, as docx4j's CR-015 revised it. The order is ECMA-376 17.7.2's,
and Word's::

    effective_ppr(direct)      = docDefaults.pPr (+) chain_ppr(style_of(direct)) (+) direct
    effective_rpr(direct, pPr) = docDefaults.rPr (+) chain_rpr(style_of(pPr))
                                 (+) chain_rpr(direct.rStyle) (+) direct
    paragraph_mark_rpr(pPr)    = docDefaults.rPr (+) chain_rpr(style_of(pPr)) (+) pPr.rPr
    table_style(tblPr)         = Word's built-in Normal Table (where the chain reaches the
                                 default table style, or the table names none)
                                 (+) the chain (+) tblPr

where ``style_of(pPr)`` is the paragraph's ``w:pStyle`` if it names a style that
exists, else the ``w:default="1"`` paragraph style (Word writes no ``w:pStyle``
for it, and treats a missing style as it), and ``chain_ppr`` / ``chain_rpr`` are
a style's ``w:basedOn`` chain merged root-first *without* the document defaults,
cached per style id. The merging itself is
:mod:`docx4j_py.model.style_util`, driven by the property catalogue. Numbering
level indents are folded in per layer; table styles are not applied to
paragraphs here (the resolver is handed a ``w:pPr`` and does not know the table).

**Colours** (CR-009 Phase B, docx4j CR-034). A resolved colour's ``w:themeColor``
(or a shading's ``w:themeFill``) is what Word draws, the value beside it a cache
which can be stale: :meth:`PropertyResolver.resolve_color`,
:meth:`~PropertyResolver.resolve_fill` and
:meth:`~PropertyResolver.resolve_pattern_color` give the colour drawn, in the
document's :attr:`~PropertyResolver.theme_colors` (its theme part, or the bundled
theme ``pkg.default_theme`` names where it has none, and the settings'
``w:clrSchemeMapping``). The effective properties themselves keep the attributes
as written, as docx4j's do. A colour inherits as one (CR-009 Phase A).

**Live objects.** What the style forms and ``get_effective_ppr(p_pr)`` return is
cached and shared: copy it before changing it. The cached objects share no leaf
with the styles part, and resolution writes nothing into the part.

**Nothing is unmarshalled.** The styles, numbering, theme and settings parts are read with
:meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart.read_contents`, which
does not mark a part for re-marshalling, so a document whose styles nobody
edited still saves byte for byte. Once something else unmarshals the part ---
the comment styles, a caller's ``contents`` --- :meth:`PropertyResolver.refresh`
picks up that live tree instead of the private copy.
"""

from __future__ import annotations

import logging
import re
from typing import Any

from docx4j_py.model.color import ThemeColors, WordColors
from docx4j_py.model.style_util import (
    apply_ppr_base,
    apply_rpr,
    apply_style,
    apply_style_level,
    apply_tbl_pr,
    has_direct_formatting_ppr,
    has_direct_formatting_rpr,
    is_cyclic,
)
from docx4j_py.model.values import effective, is_on, on_off

__all__ = [
    "WORD_DEFAULT_CELL_MARGIN_TWIPS",
    "PropertyResolver",
    "heading_level_by_name",
    "resolver_of",
]

logger = logging.getLogger(__name__)

#: The cache key for "no style": a styles part with no default paragraph style.
_NO_STYLE = ""

#: 108 twips (0.08 in): the left and right cell margin of Word's built-in Normal Table.
WORD_DEFAULT_CELL_MARGIN_TWIPS = 108

_HEADING_NAME = re.compile(r"^heading ([1-9])$", re.IGNORECASE)
_HEADING_STYLE = "Heading"

#: ``StyleTree``'s virtual style for the document defaults, which this resolver
#: applies itself: a ``w:basedOn`` naming it is not a missing style.
_DOC_DEFAULTS = "DocDefaults"


def heading_level_by_name(style: Any) -> int:
    """The outline level (1 to 9) of a built-in heading style, or -1.

    From its ``w:name`` (``heading 1`` ... ``heading 9``). Word identifies
    built-in styles by name, so this holds in every locale;
    :meth:`PropertyResolver.get_lvl_from_heading_style` keys on the English
    style id instead, which is what docx4j's callers pass.
    """
    name = getattr(getattr(style, "name", None), "val", None)
    if name is None:
        return -1
    match = _HEADING_NAME.match(name.strip())
    return int(match.group(1)) if match else -1


def _style_kind(style: Any) -> str | None:
    value = getattr(style, "type_value", None)
    return None if value is None else str(getattr(value, "value", value))


class PropertyResolver:
    """The effective paragraph, run and table properties of a package.

    Built over a :class:`~docx4j_py.openpackaging.packages.wordprocessingml_package.WordprocessingMLPackage`
    (``PropertyResolver.of(package)``, which
    :meth:`~docx4j_py.openpackaging.packages.wordprocessingml_package.WordprocessingMLPackage.get_property_resolver`
    caches one of per package), or directly over a ``w:styles`` tree for a test.
    """

    __slots__ = (
        "_any_vanish",
        "_any_vanish_count",
        "_chain_ppr",
        "_chain_rpr",
        "_character_level_rpr",
        "_default_character_style_id",
        "_default_table_style_id",
        "_document_default_ppr",
        "_document_default_rpr",
        "_effective_ppr_by_style",
        "_effective_rpr_by_style",
        "_live_styles",
        "_missing_logged",
        "_numbering_indents",
        "_package",
        "_scanned_style_count",
        "_styles",
        "_styles_are_private",
        "_theme_colors",
        "_vanish",
        "default_paragraph_style_id",
    )

    def __init__(self, package: Any = None, styles: Any = None) -> None:
        """Build a resolver over `package`, or over a ``w:styles`` tree directly."""
        self._package = package
        self._styles = styles
        self._styles_are_private = False
        self._chain_ppr: dict[str, Any] = {}
        self._chain_rpr: dict[str, Any] = {}
        #: A character style's chain as a level over a run or a mark, where it differs from
        #: ``_chain_rpr``: without its highlight (:meth:`_character_level`).
        self._character_level_rpr: dict[str, Any] = {}
        self._effective_ppr_by_style: dict[str, Any] = {}
        self._effective_rpr_by_style: dict[str, Any] = {}
        self._live_styles: dict[str, Any] = {}
        self._missing_logged: set[str] = set()
        self._numbering_indents: Any = None
        self._scanned_style_count = -1
        #: ``is_vanished_by_styles``'s answers, by paragraph style then character style.
        self._vanish: dict[tuple[str | None, str | None], bool] = {}
        #: Whether any style or the document defaults state ``w:vanish``; None: not known yet.
        self._any_vanish: bool | None = None
        self._any_vanish_count = -1
        self.default_paragraph_style_id: str | None = None
        self._default_character_style_id: str | None = None
        self._default_table_style_id: str | None = None
        #: :attr:`theme_colors`, made on first use and dropped by :meth:`refresh`.
        self._theme_colors: Any = None
        self._read_source()
        self._init()

    @staticmethod
    def of(package: Any) -> PropertyResolver:
        """The resolver for a package. docx4j's constructor."""
        return PropertyResolver(package)

    @property
    def package(self) -> Any:
        """The package this resolver is over, or None where it is over a tree."""
        return self._package

    # -- the source --------------------------------------------------------

    def _read_source(self) -> None:
        """Read the styles part, without marking it unmarshalled."""
        if self._package is None:
            return
        part = getattr(self._package, "style_definitions_part", None)
        if part is None:
            self._styles = None
            self._styles_are_private = False
            return
        if part.is_unmarshalled:
            self._styles = part.contents
            self._styles_are_private = False
        else:
            self._styles = part.read_contents()
            self._styles_are_private = True

    def _init(self) -> None:
        from docx4j_py.child import deep_copy
        from docx4j_py.wml import HpsMeasure, PPr, RPr

        styles = self._styles
        self._live_styles.clear()
        self._scanned_style_count = -1
        self.default_paragraph_style_id = None
        self._default_character_style_id = None
        self._default_table_style_id = None
        if styles is not None:
            self._scan_styles()
            for style in styles.style or ():
                if not on_off(effective(style, "default")) or style.style_id is None:
                    continue
                kind = _style_kind(style)
                if kind == "paragraph" and self.default_paragraph_style_id is None:
                    self.default_paragraph_style_id = style.style_id
                elif kind == "character" and self._default_character_style_id is None:
                    self._default_character_style_id = style.style_id
                elif kind == "table" and self._default_table_style_id is None:
                    self._default_table_style_id = style.style_id
        if self.default_paragraph_style_id is None:
            logger.warning("No default paragraph style!!")

        # Private copies: the resolver's defaults are its own, so nothing below
        # writes into the styles part. Until docx4j's CR-015 phase 3 the w:sz 20
        # default went into the part, and a docx saved after an export carried it.
        doc_defaults = None if styles is None else styles.doc_defaults
        p_pr_default = None if doc_defaults is None else doc_defaults.p_pr_default
        r_pr_default = None if doc_defaults is None else doc_defaults.r_pr_default
        source_ppr = None if p_pr_default is None else p_pr_default.p_pr
        source_rpr = None if r_pr_default is None else r_pr_default.r_pr
        self._document_default_ppr = PPr() if source_ppr is None else deep_copy(source_ppr)
        self._document_default_rpr = RPr() if source_rpr is None else deep_copy(source_rpr)
        if self._document_default_rpr.sz is None:
            # Word's default made explicit: 10 pt where nothing states a size
            # (measured, docx4j CR-015's probe styles-no-size-anywhere: identical
            # to an explicit 10 pt run)
            self._document_default_rpr.sz = HpsMeasure(val=20)

        self._numbering_indents = None

    def _scan_styles(self) -> None:
        """(Re)read the styles part into the live map. A style with no id is skipped."""
        entries = [] if self._styles is None else list(self._styles.style or ())
        for style in entries:
            if style.style_id is None:
                continue
            self._live_styles[style.style_id] = style
        self._scanned_style_count = len(entries)

    def refresh(self) -> PropertyResolver:
        """Discard the cached state and re-read the styles part.

        A style merely *added* is picked up on its own; call this when a style
        has been modified or removed, or the styles part's contents replaced.
        The content API's ``ensure_style`` calls it where it adds a style. The
        theme colours are read again too (a changed theme part or
        ``w:clrSchemeMapping``), as docx4j's ``refresh`` re-reads them.
        """
        self._theme_colors = None
        self._chain_ppr.clear()
        self._chain_rpr.clear()
        self._character_level_rpr.clear()
        self._effective_ppr_by_style.clear()
        self._effective_rpr_by_style.clear()
        self._missing_logged.clear()
        self._vanish.clear()
        self._any_vanish = None
        self._read_source()
        self._init()
        return self

    # -- styles ------------------------------------------------------------

    def _live_style(self, style_id: str | None) -> Any:
        """The style with this id, rescanning on a miss when the part has grown.

        A style *added* to the part after this resolver was built is still
        found; one *modified* or *removed* needs :meth:`refresh`.
        """
        if style_id is None:
            return None
        style = self._live_styles.get(style_id)
        if style is None:
            count = 0 if self._styles is None else len(self._styles.style or ())
            if count != self._scanned_style_count:
                self._scan_styles()
                style = self._live_styles.get(style_id)
        return style

    def get_style(self, style_id: str | None) -> Any:
        """The style with this id, or None. docx4j ``getStyle``."""
        return self._live_style(style_id)

    def _log_missing(self, style_id: str) -> None:
        if style_id in self._missing_logged:
            return
        self._missing_logged.add(style_id)
        logger.warning("Style definition not found: %s (logged once)", style_id)

    def ancestry(self, style_id: str | None) -> list[Any]:
        """The style and the styles it is based on, **root first**.

        The base of the chain at index 0 and the style itself last; empty for an
        unknown id. A cycle, or a chain deeper than 32, ends the walk where it
        is detected (and raises ``CyclicStylesException`` when
        :func:`~docx4j_py.model.style_util.throw_on_cyclic_styles` is on).
        """
        leaf_first: list[Any] = []
        seen: list[str] = []
        current = style_id
        while current is not None:
            if is_cyclic(current, seen):
                break
            seen.append(current)
            style = self._live_style(current)
            if style is None:
                if current != _DOC_DEFAULTS:
                    self._log_missing(current)
                break
            leaf_first.append(style)
            current = getattr(style.based_on, "val", None)
        leaf_first.reverse()
        return leaf_first

    # -- the numbering seam ------------------------------------------------

    @property
    def numbering_indents(self) -> Any:
        """What the paragraph merge asks for a numbering level's ``w:ind``.

        The package's numbering emulator (docx4j's
        ``NumberingDefinitionsPart.getInd``), or None where there is no package
        or no numbering part. Looked up lazily and not cached beyond the
        package's own cache, so that a numbering part added or replaced after
        this resolver was built is still seen.
        """
        if self._package is None:
            return None
        from docx4j_py.model.listnumbering import emulator_of

        emulator = emulator_of(self._package)
        return None if emulator is None or emulator.part is None else emulator

    # -- colours (CR-009 Phase B) -----------------------------------------

    @property
    def theme_colors(self) -> Any:
        """The document's :class:`~docx4j_py.model.color.ThemeColors`. docx4j ``getThemeColors``.

        Its theme part (or the bundled theme ``pkg.default_theme`` names where
        it has none) and the settings' ``w:clrSchemeMapping``, both read without
        being unmarshalled. Made once, and again after :meth:`refresh`. A
        resolver over a ``w:styles`` tree alone has the process-wide default
        theme and the default colour map.
        """
        colours = self._theme_colors
        if colours is None:
            colours = ThemeColors.of_theme(None) if self._package is None else ThemeColors.of(self._package)
            self._theme_colors = colours
        return colours

    def get_theme_colors(self) -> Any:
        """:attr:`theme_colors`, by docx4j's name."""
        return self.theme_colors

    def resolve_color(self, obj: Any) -> Any:
        """The colour Word draws for a ``w:color``, a border or a ``w:u``; None for automatic.

        docx4j ``resolveColor(Color)``, ``resolveColor(CTBorder)`` and
        ``resolveColor(U)``: the theme reference where there is one (the value
        beside it is a cache), else the value, as a
        :class:`~docx4j_py.model.color.ResolvedColor`. None is the caller's
        default: black text, as a rule; an underline's is the text's colour.
        """
        return WordColors.of(obj, self.theme_colors)

    def resolve_fill(self, shd: Any) -> Any:
        """A shading's fill (``w:fill``, ``w:themeFill``); None for automatic. docx4j ``resolveFill``."""
        return WordColors.fill(shd, self.theme_colors)

    def resolve_pattern_color(self, shd: Any) -> Any:
        """A shading's pattern colour (``w:color``, ``w:themeColor``); None for automatic. docx4j ``resolvePatternColor``."""
        return WordColors.pattern(shd, self.theme_colors)

    # -- defaults ----------------------------------------------------------

    def get_document_default_ppr(self) -> Any:
        """``w:docDefaults/w:pPrDefault/w:pPr``, as a private copy."""
        return self._document_default_ppr

    def get_document_default_rpr(self) -> Any:
        """``w:docDefaults/w:rPrDefault/w:rPr``, with Word's 10 pt made explicit."""
        return self._document_default_rpr

    def get_default_paragraph_style_id(self) -> str | None:
        """The ``w:default="1"`` paragraph style's id, or None."""
        return self.default_paragraph_style_id

    @property
    def default_character_style_id(self) -> str | None:
        """The ``w:default="1"`` character style's id, or None."""
        return self._default_character_style_id

    @property
    def default_table_style_id(self) -> str | None:
        """The ``w:default="1"`` table style's id, or None."""
        return self._default_table_style_id

    @property
    def styles_tree(self) -> Any:
        """The ``w:styles`` this resolver read, or None where the package has no part.

        Read without unmarshalling the part unless something else already had
        (:meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart.read_contents`);
        the font walk and the numbering definitions read it through here rather
        than opening the part again.
        """
        return self._styles

    def get_resolved_default_paragraph_style(self) -> Any:
        """The default paragraph style's effective ``w:pPr``."""
        return self.get_effective_ppr(self.default_paragraph_style_id)

    # -- the chains --------------------------------------------------------

    def get_chain_ppr(self, style_id: str | None) -> Any:
        """A paragraph style's ``w:basedOn`` chain merged root-first, without the defaults.

        What the style contributes on its own. A live, cached object: copy it
        before changing it.
        """
        from docx4j_py.wml import PPr

        if style_id is None:
            return PPr()
        cached = self._chain_ppr.get(style_id)
        if cached is not None:
            return cached
        chain = PPr()
        for style in self.ancestry(style_id):
            self._apply_ppr_layer(self._heading_layer(style), chain)
        self._chain_ppr[style_id] = chain
        return chain

    def get_chain_rpr(self, style_id: str | None) -> Any:
        """A style's ``w:basedOn`` chain merged root-first, without the defaults.

        For a character style, what it contributes over the paragraph's run
        properties. A live, cached object: copy it before changing it.
        """
        from docx4j_py.wml import RPr

        if style_id is None:
            return RPr()
        cached = self._chain_rpr.get(style_id)
        if cached is not None:
            return cached
        chain = RPr()
        for style in self.ancestry(style_id):
            apply_rpr(style.r_pr, chain, RPr)
        self._chain_rpr[style_id] = chain
        return chain

    def _character_level(self, style_id: str) -> Any:
        """A character style's chain as a level over a run or a paragraph mark: :meth:`get_chain_rpr` without its highlight.

        Word draws no highlight from a character style (CR-018 section 3.1c item
        3, round six): alone, with bold, or through ``w:basedOn``, with no repair
        involved, and its save writes ``w:highlight w:val="none"`` on such a run.
        A paragraph style's highlight and a run's own are drawn, and so is a
        character style's ``w:shd``. docx4j's ``PropertyResolver.characterLevel``
        (CR-060 phase 4, ``1d1b824a9``). A character style's own resolution
        (``get_effective_rpr(style_id)``) keeps it, as docx4j's does. Cached;
        live: copy it before changing it.
        """
        from docx4j_py.child import deep_copy

        chain = self.get_chain_rpr(style_id)
        if getattr(chain, "highlight", None) is None:
            return chain
        level = self._character_level_rpr.get(style_id)
        if level is None:
            level = deep_copy(chain)
            level.highlight = None
            self._character_level_rpr[style_id] = level
        return level

    def _heading_layer(self, style: Any) -> Any:
        """A style's own ``w:pPr`` as a layer of its chain.

        The built-in heading styles have a fixed outline level, which Word takes
        from the style's built-in **name** (``heading 1`` ... ``heading 9``, the
        same in every locale; the style id is localised: ``Überschrift1``,
        ``Titre1``), whatever ``w:outlineLvl`` the style declares. Where the two
        differ the layer is a copy with the level from the name, and the styles
        part is not touched.
        """
        from docx4j_py.child import deep_copy

        layer = style.p_pr
        level = heading_level_by_name(style)
        if (
            level > 0
            and layer is not None
            and layer.outline_lvl is not None
            and layer.outline_lvl.val is not None
            and layer.outline_lvl.val != level - 1
        ):
            copy = deep_copy(layer)
            copy.outline_lvl.val = level - 1
            return copy
        return layer

    def _apply_ppr_layer(self, source: Any, destination: Any) -> None:
        from docx4j_py.wml import ParaRPr

        if source is None:
            return
        apply_ppr_base(source, destination, self.numbering_indents)
        destination.r_pr = apply_rpr(
            getattr(source, "r_pr", None), destination.r_pr, ParaRPr
        )

    def _paragraph_style_of(self, p_pr: Any) -> str | None:
        """The paragraph's style: its ``w:pStyle`` where that exists, else the default."""
        if p_pr is None or p_pr.p_style is None:
            return self.default_paragraph_style_id
        return self._existing_paragraph_style(p_pr.p_style.val)

    def _existing_paragraph_style(self, style_id: str | None) -> str | None:
        """`style_id` if it names a style that exists, else the default's id (logged once)."""
        if style_id is None:
            return self.default_paragraph_style_id
        if self._live_style(style_id) is None:
            self._log_missing(style_id)
            return self.default_paragraph_style_id
        return style_id

    # -- paragraphs --------------------------------------------------------

    def get_effective_ppr(self, p_pr_or_style_id: Any) -> Any:
        """The paragraph properties which actually apply.

        Given a paragraph's own ``w:pPr``: the document defaults, the paragraph
        style's chain, then the direct formatting. Given a **style id** (a
        ``str``, or None): what that style resolves to, document defaults
        included; a style id naming no style resolves as the default paragraph
        style does.

        Run properties are not resolved here; the paragraph mark's ``w:rPr`` is
        carried as ``w:pPr/w:rPr`` but is not a run's formatting
        (:meth:`get_effective_paragraph_mark_rpr`).

        What is returned is a live, cached object when the paragraph states no
        direct formatting: copy it before changing it.
        """
        from docx4j_py.child import deep_copy

        if p_pr_or_style_id is None or isinstance(p_pr_or_style_id, str):
            return self._effective_ppr_of_style(p_pr_or_style_id)
        resolved = self._effective_ppr_of_style(self._paragraph_style_of(p_pr_or_style_id))
        if not has_direct_formatting_ppr(p_pr_or_style_id):
            return resolved
        effective = deep_copy(resolved)
        self._apply_ppr_layer(p_pr_or_style_id, effective)
        return effective

    def _effective_ppr_of_style(self, style_id: str | None) -> Any:
        from docx4j_py.child import deep_copy

        existing = self._existing_paragraph_style(style_id)
        key = _NO_STYLE if existing is None else existing
        cached = self._effective_ppr_by_style.get(key)
        if cached is not None:
            return cached
        resolved = deep_copy(self._document_default_ppr)
        self._apply_ppr_layer(self.get_chain_ppr(existing), resolved)
        self._effective_ppr_by_style[key] = resolved
        return resolved

    # -- runs --------------------------------------------------------------

    def get_effective_rpr(self, r_pr_or_style_id: Any, p_pr: Any = None) -> Any:
        """The run properties which apply.

        Given a run's own ``w:rPr`` and the ``w:pPr`` of the paragraph it is in:
        document defaults, the paragraph style's run properties, the run's
        character style, then its direct formatting. A new object each call.

        Given a **style id** (a ``str``): what that style resolves to on its own
        --- document defaults, then its ``w:basedOn`` chain. For a paragraph
        style, the run properties of its paragraphs; for a character style, what
        it contributes with nothing under it. None where no such style exists.

        The paragraph mark's ``w:rPr`` is never applied to a run; for the mark
        itself see :meth:`get_effective_paragraph_mark_rpr`.
        """
        from docx4j_py.child import deep_copy
        from docx4j_py.wml import RPr

        if isinstance(r_pr_or_style_id, str):
            return self._effective_rpr_of_style(r_pr_or_style_id)
        effective = deep_copy(self._document_default_rpr)
        apply_rpr(self.get_chain_rpr(self._paragraph_style_of(p_pr)), effective, RPr)
        self._apply_character_style_and_direct(r_pr_or_style_id, effective)
        return effective

    def is_vanished_by_styles(self, paragraph_style_id: str | None, run_style_id: str | None) -> bool:
        """Whether a run with no ``w:vanish`` of its own is hidden by its styles.

        The effective ``w:vanish`` of its paragraph style (None: the default
        paragraph style) and its character style, the table style left out, as
        the text model's hidden rule has it (WD REQ-507;
        ``model/text/extractor.py``'s ``_run``). What ``Paragraph.text`` leaves
        out, as Office JS does (CR-013 Phase A; docx4j-core-ts
        ``PropertyResolver.isVanishedByStyles`` at ``9307bdb``, check 48b case
        20). Cached per pair; a pair naming a style the part lacks is not
        cached, so that adding the style is seen. A document none of whose
        styles states ``w:vanish``, nor its defaults, answers False at once.
        """
        if self._any_vanish is None or self._any_vanish_count != self._style_count():
            self._any_vanish = self._states_vanish()
            self._any_vanish_count = self._style_count()
        if not self._any_vanish:
            return False
        key = (paragraph_style_id, run_style_id)
        found = self._vanish.get(key)
        if found is not None:
            return found
        from docx4j_py.wml import PPr, RPr, RStyle

        r_pr = RPr(r_style=RStyle(val=run_style_id)) if run_style_id is not None else None
        p_pr = None
        if paragraph_style_id is not None:
            from docx4j_py.wml import CtPprBasePStyle

            p_pr = PPr(p_style=CtPprBasePStyle(val=paragraph_style_id))
        try:
            vanish = getattr(self.get_effective_rpr(r_pr, p_pr), "vanish", None)
        except Exception:  # noqa: BLE001 - not resolved: not hidden
            return False
        hidden = is_on(vanish)
        if (paragraph_style_id is None or self._live_style(paragraph_style_id) is not None) and (
            run_style_id is None or self._live_style(run_style_id) is not None
        ):
            self._vanish[key] = hidden
        return hidden

    def _style_count(self) -> int:
        return 0 if self._styles is None else len(self._styles.style or ())

    def _states_vanish(self) -> bool:
        """Whether the document defaults or any style's run properties state ``w:vanish``."""
        if getattr(self._document_default_rpr, "vanish", None) is not None:
            return True
        for style in () if self._styles is None else (self._styles.style or ()):
            if getattr(getattr(style, "r_pr", None), "vanish", None) is not None:
                return True
        return False

    def _effective_rpr_of_style(self, style_id: str) -> Any:
        from docx4j_py.child import deep_copy
        from docx4j_py.wml import RPr

        cached = self._effective_rpr_by_style.get(style_id)
        if cached is not None:
            return cached
        if self._live_style(style_id) is None:
            self._log_missing(style_id)
            return None
        resolved = deep_copy(self._document_default_rpr)
        apply_rpr(self.get_chain_rpr(style_id), resolved, RPr)
        self._effective_rpr_by_style[style_id] = resolved
        return resolved

    def _apply_character_style_and_direct(self, r_pr: Any, effective: Any) -> None:
        """The character style's chain, then the direct formatting.

        The character style is a **level** of the style hierarchy over the
        paragraph's, so its toggle properties are XORed with what is beneath
        rather than overriding it (ECMA-376-1 17.7.3). The direct formatting
        after it is not a level: an explicit value there is used as it stands.
        """
        from docx4j_py.wml import RPr

        run_style_id = getattr(getattr(r_pr, "r_style", None), "val", None)
        if run_style_id is not None:
            if self._live_style(run_style_id) is None:
                self._log_missing(run_style_id)
            else:
                apply_style_level(
                    self._character_level(run_style_id), effective, self._document_default_rpr
                )
        if has_direct_formatting_rpr(r_pr):
            apply_rpr(r_pr, effective, RPr)

    def get_effective_paragraph_mark_rpr(self, p_pr: Any) -> Any:
        """The run properties of the paragraph mark.

        Document defaults, the paragraph style's run properties, the mark's
        character style, then the ``w:pPr``'s own ``w:rPr``. What sizes an empty
        paragraph, and what a list label starts from.
        """
        from docx4j_py.child import deep_copy
        from docx4j_py.wml import RPr

        effective = deep_copy(self._document_default_rpr)
        apply_rpr(self.get_chain_rpr(self._paragraph_style_of(p_pr)), effective, RPr)
        mark = getattr(p_pr, "r_pr", None)
        if mark is not None:
            self._apply_mark_rpr(mark, effective)
        return effective

    def _apply_mark_rpr(self, mark: Any, effective: Any) -> None:
        """The mark's character style (``w:pPr/w:rPr/w:rStyle``) as a level, then its direct formatting.

        The order :meth:`_apply_character_style_and_direct` gives a run. docx4j
        ``applyMarkRPr``, since 17.3.2 (``fb2a764f7``): until then the mark's
        ``w:rStyle`` was copied as an attribute and never resolved, so an empty
        paragraph whose mark names a 9pt character style was sized at its
        paragraph style's 12pt.
        """
        from docx4j_py.wml import RPr

        run_style_id = getattr(getattr(mark, "r_style", None), "val", None)
        if run_style_id is not None:
            if self._live_style(run_style_id) is None:
                self._log_missing(run_style_id)
            else:
                apply_style_level(
                    self._character_level(run_style_id), effective, self._document_default_rpr
                )
        apply_rpr(mark, effective, RPr)

    def has_direct_rpr_formatting(self, r_pr: Any) -> bool:
        """Whether the run states any formatting of its own. docx4j ``hasDirectRPrFormatting``."""
        return has_direct_formatting_rpr(r_pr)

    # -- tables ------------------------------------------------------------

    def get_effective_table_style(self, tbl_pr: Any) -> Any:
        """The table style which applies, merged root-first, then the table's own ``w:tblPr``.

        Word's built-in *Normal Table* (``w:tblInd`` 0; cell margins 108 twips
        left and right, 0 top and bottom) underlies a table naming no style and
        a table whose chain reaches the document's default table style --- and
        it is the **built-in** that applies, not the document's definition of
        that style (measured, docx4j CR-015's probe ``styles-table-default``).
        So the default style's own layer is skipped in favour of the built-in,
        and a chain that does not reach it starts from nothing.

        A new ``w:style`` each call, with a non-None ``w:tblPr``.
        """
        from docx4j_py.wml import CTTblPrBase

        chain = self._table_chain(tbl_pr)
        result = (
            self._built_in_table_normal()
            if self._reaches_default(chain)
            else _empty_table_style()
        )
        for layer in chain:
            # the built-in stands in for the document's definition of the default
            if (
                self._default_table_style_id is not None
                and layer.style_id == self._default_table_style_id
            ):
                continue
            apply_style(layer, result)
        if tbl_pr is not None:
            result.tbl_pr = apply_tbl_pr(tbl_pr, result.tbl_pr)
        if result.tbl_pr is None:
            result.tbl_pr = CTTblPrBase()
        return result

    def reaches_default_table_style(self, tbl_pr: Any) -> bool:
        """Whether a table's style chain reaches the document's default table style.

        The flag :meth:`get_effective_table_style` decides Word's built-in Normal
        Table by; a table naming no style counts as reaching it.
        """
        return self._reaches_default(self._table_chain(tbl_pr))

    def _table_chain(self, tbl_pr: Any) -> list[Any]:
        style_id = getattr(getattr(tbl_pr, "tbl_style", None), "val", None)
        return [] if style_id is None else self.ancestry(style_id)

    def _reaches_default(self, chain: list[Any]) -> bool:
        return not chain or (
            self._default_table_style_id is not None
            and any(s.style_id == self._default_table_style_id for s in chain)
        )

    def _built_in_table_normal(self) -> Any:
        """Word's built-in Normal Table, whatever the document's own definition says."""
        from docx4j_py.wml import CTTblCellMar, StyleName

        style = _empty_table_style()
        style.style_id = self._default_table_style_id or "TableNormal"
        style.name = StyleName(val="Normal Table")
        style.tbl_pr.tbl_ind = _twips(0)
        style.tbl_pr.tbl_cell_mar = CTTblCellMar(
            top=_twips(0),
            left=_twips(WORD_DEFAULT_CELL_MARGIN_TWIPS),
            bottom=_twips(0),
            right=_twips(WORD_DEFAULT_CELL_MARGIN_TWIPS),
        )
        return style

    # -- headings and known styles ----------------------------------------

    def get_lvl_from_heading_style(self, style_id: str) -> int:
        """The level of a heading style from its **id** (``Heading2`` -> 2), or -1.

        docx4j keeps this because callers pass a style id;
        :func:`heading_level_by_name` is the rule Word applies, and is what
        resolution uses.
        """
        if not style_id.startswith(_HEADING_STYLE):
            return -1
        rest = style_id[len(_HEADING_STYLE) :].strip()
        return int(rest) if rest.isdigit() else -1

    def activate_style(self, style: Any) -> bool:
        """Make a style available for use, adding it to the styles part.

        And, recursively, whatever it is ``w:basedOn`` and whatever it is linked
        to. Given a **style id**, the style must be one docx4j's
        ``KnownStyles.xml`` has (the content API's
        :func:`~docx4j_py.model.content.styles.ensure_style` is what puts it
        there), or one already live.

        Raises:
            Docx4JException: the styles part is not unmarshalled, so activating
                would write into the resolver's private copy and nothing would
                be saved.
        """
        from docx4j_py.openpackaging.exceptions import Docx4JException

        if isinstance(style, str):
            if self._live_style(style) is not None:
                return True
            logger.warning("Unknown style: %s", style)
            return False
        if self._styles_are_private:
            raise Docx4JException(
                "activate_style writes into the styles part: read "
                "style_definitions_part.contents and refresh() the resolver first"
            )
        return self._activate(style, replace=True)

    def _activate(self, style: Any, replace: bool) -> bool:
        if style.style_id is None or self._styles is None:
            return False
        entries = self._styles.style
        existing = self._live_style(style.style_id)
        if existing is not None:
            if not replace:
                return False
            entries.remove(existing)
        entries.append(style)
        self._live_styles[style.style_id] = style
        self._scanned_style_count = len(entries)
        # resolution of this style, and of anything based on it, has to be redone
        self._chain_ppr.clear()
        self._chain_rpr.clear()
        self._character_level_rpr.clear()
        self._effective_ppr_by_style.clear()
        self._effective_rpr_by_style.clear()
        self._vanish.clear()
        self._any_vanish = None

        ok = True
        based_on = getattr(style.based_on, "val", None)
        if based_on is not None:
            ok = self.activate_style(based_on)
        link = getattr(style.link, "val", None)
        if link is not None:
            ok = self.activate_style(link) and ok
        return ok

    def __repr__(self) -> str:
        """The style count and the default paragraph style."""
        return (
            f"PropertyResolver({len(self._live_styles)} styles, "
            f"default={self.default_paragraph_style_id!r})"
        )


def _twips(value: int) -> Any:
    from docx4j_py.wml import CtTblWidthType, TblWidth

    return TblWidth(w=value, type_value=CtTblWidthType.DXA)


def _empty_table_style() -> Any:
    from docx4j_py.wml import CTTblPrBase, Style, StyleType

    return Style(type_value=StyleType.TABLE, tbl_pr=CTTblPrBase())


def resolver_of(package: Any) -> PropertyResolver | None:
    """The package's resolver, made on first use and kept. None without a package."""
    if package is None:
        return None
    get = getattr(package, "get_property_resolver", None)
    return None if get is None else get()
