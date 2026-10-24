"""Colours as Word draws them: docx4j ``org.docx4j.model.color`` (its CR-034), the WordprocessingML half.

CR-009 Phase B. Any WordprocessingML colour reference --- ``w:color``,
``w:shd``'s fill and pattern colour, a border's colour, ``w:u``'s colour ---
resolved to the RGB value Word draws for it. A reference resolves against a
:class:`ThemeColors`: a theme's colour scheme and the colour map that leads to
it, made for a docx by :meth:`ThemeColors.of` (the main document part's theme,
or the bundled one ``pkg.default_theme`` names where it has none, and the
settings' ``w:clrSchemeMapping``), and held once per
:class:`~docx4j_py.model.property_resolver.PropertyResolver` as its
:attr:`~docx4j_py.model.property_resolver.PropertyResolver.theme_colors`.

* :class:`WordColors`: Word draws the theme reference and keeps the value beside
  it as a cache, which can be stale; ``w:themeTint`` and ``w:themeShade`` change
  the HSL luminance, the tint winning where both are given; Word's written value
  is returned where it is within one unit of the computation.
* :class:`ResolvedColor`: red, green, blue and an alpha in OOXML's unit.

Each returns a :class:`ResolvedColor`, or None for automatic, which is the
caller's default. Every rule was read from Word's own renderings (docx4j
CR-034 section 11); its probes, Word's readings and the corpus tint and shade
triples are this package's vectors too (``tests/fixtures/theme-colors/``).

**Phase C** (DrawingML and SpreadsheetML: ``DrawingMLColors``,
``SpreadsheetColors``, ``StyleFill``, the slide, workbook and chart
``ThemeColors``) is not here; :func:`theme_colors._slot_colour` is the seam
where ``DrawingMLColors`` will read a theme's slots.
"""

from __future__ import annotations

from docx4j_py.model.color.resolved import OPAQUE, ResolvedColor
from docx4j_py.model.color.theme_colors import ThemeColors
from docx4j_py.model.color.word_colors import WordColors

__all__ = ["OPAQUE", "ResolvedColor", "ThemeColors", "WordColors"]
