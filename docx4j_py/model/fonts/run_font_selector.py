"""``RunFontSelector``: which of a run's fonts formats each character.

docx4j ``org.docx4j.fonts.RunFontSelector``. CR-002 section 15.3.

**The selection core only.** docx4j's class is 2,572 lines, most of them XSL-FO
output: element creation, the kerned and no-ligature twins, character scaling,
small caps, line metrics, and the glyph-coverage pass that replaces a face which
lacks a script. What CR-002 section 6.3 promised, and what a browser, a text
extractor or a measurer needs, is the **decision per character**, so what is
ported is ``documentFontsOf``, ``defaultFontOf``, the theme lookup by
``themeFontLang``, ``resolvedSlots``, ``complexScriptFont`` (``w:cs`` and
``w:rtl`` **by value**), ``preambleRule``, the code-point-to-slot table
``fontFor`` branch for branch, ``spanScript``, ``isEmoji``, ``symbolFontName``
and the ``w:hint`` rules.

**One resolution** (docx4j CR-016 phase 1): the selector consumes the run's
*effective* properties. It resolves them itself through the
:class:`~docx4j_py.model.property_resolver.PropertyResolver` unless the caller
says it has done so already (``r_pr_is_effective``), which is what a walk over a
whole document does.

What the goldens settled, and is reproduced here: ``<w:cs w:val="0"/>`` turns an
inherited complex-script flag **off** (until docx4j 17.1.1 the element's presence
was tested, so a false value took the cs font too); ``w:rtl`` alone on Latin text
*does* take the cs font; the theme language is matched by exact subtag, so
Estonian is not Ethiopic; U+0020 is in the ASCII range and takes the ascii font,
so the space between two East Asian words is the Latin font's; U+0590 to U+07BF
in a run with no ``w:cs`` takes the ascii font; a symbol font is recognised
whatever case the document wrote it in; an East Asian reference with no East
Asian font falls to hAnsi in every branch.
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.model.fonts.default_theme import DefaultTheme, default_theme_setting
from docx4j_py.model.fonts.fallback import coverage_group_of
from docx4j_py.model.fonts.theme_fonts import theme_font_of

__all__ = [
    "BUILT_IN_DEFAULT_FONT",
    "FontSpan",
    "RunFontSelector",
    "is_emoji",
    "is_on",
    "preamble_rule",
    "run_locale",
    "span_script",
    "symbol_font_name",
]

#: Word's built-in default where the document defaults name no font at all
#: (docx4j CR-016 probe ``fonts-missing-slots`` (a)).
BUILT_IN_DEFAULT_FONT = "Times New Roman"

#: The fonts :func:`symbol_font_name` recognises; docx4j's ``SymbolMapper`` knows these five.
_SYMBOL_FONTS = ("Symbol", "Webdings", "Wingdings", "Wingdings 2", "Wingdings 3")


@dataclasses.dataclass(frozen=True, slots=True)
class FontSpan:
    """One stretch of a run's text set in one document font."""

    text: str
    """The text of the span."""
    document_font: str | None
    """The document font: the name a ``w:rFonts`` slot gives, theme reference
    resolved. None where the run names no font at all and there is no default."""
    bold: bool = False
    italic: bool = False
    cs: bool = False
    """``w:cs`` of the effective run properties, by value."""
    rtl: bool = False
    """``w:rtl`` of the effective run properties, by value."""
    script: str | None = None
    """The span's script, as :func:`span_script` groups it (``LATIN``, ``CJK``,
    ``SYMBOL``, ``EMOJI`` ...); None where every character of the span is shared
    (a space, a digit, punctuation)."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, as the parity goldens record a span."""
        return {
            "text": self.text,
            "documentFont": self.document_font,
            "bold": self.bold,
            "italic": self.italic,
            "cs": self.cs,
            "rtl": self.rtl,
        }


def is_on(value: Any) -> bool:
    """``w:cs`` and ``w:rtl`` are ``ST_OnOff``: present without ``w:val`` means true."""
    return value is not None and value.val is not False


def symbol_font_name(document_font_name: str | None) -> str | None:
    """The canonical name of one of the symbol fonts, whatever case the document used.

    None for any other font. Font names are case-insensitive, and Word draws
    ``symbol`` with Symbol (docx4j CR-016 probe ``fonts-symbol-and-emoji`` (a),
    (b)).
    """
    if document_font_name is None:
        return None
    name = document_font_name.strip().lower()
    return next((known for known in _SYMBOL_FONTS if known.lower() == name), None)


def is_emoji(code_point: int) -> bool:
    """The emoji blocks, U+1F000 to U+1FAFF; Word sets them in Segoe UI Emoji."""
    return 0x1F000 <= code_point <= 0x1FAFF


def span_script(code_point: int) -> str | None:
    """The group a span is cut by.

    The coverage group of the code point (its Unicode script, or ``SYMBOL`` /
    ``EMOJI``), the CJK scripts as one, and None for a shared character
    (``COMMON``, ``INHERITED``), which joins whatever it follows.

    docx4j's rendering walk cuts a span where the *script* changes between two
    non-shared characters as well as where the font does, so that a Greek stretch
    inside a Latin font is substituted as one top-level span; measured, nesting it
    changed FOP's line stacking and cost a corpus document a page.
    :meth:`RunFontSelector.spans` folds by font alone --- which is what
    ``document_font_for`` per code point answers, and what the parity goldens
    record --- and reports the script, so a renderer can cut further.
    """
    group = coverage_group_of(code_point)
    if group in ("COMMON", "INHERITED"):
        return None
    if group in ("HAN", "HIRAGANA", "KATAKANA", "HANGUL", "BOPOMOFO"):
        return "CJK"
    return group


def preamble_rule(east_asia: str | None, ascii_font: str | None, h_ansi: str | None) -> bool:
    """The table's preamble.

    An eastAsia of Times New Roman with ascii and hAnsi equal sets the whole run
    in the ascii font.
    """
    return (
        east_asia == "Times New Roman"
        and ascii_font is not None
        and ascii_font == h_ansi
    )


def run_locale(r_pr: Any) -> str | None:
    """The run's language, for case conversion: ``w:lang/@w:val``, else None."""
    tag = getattr(getattr(r_pr, "lang", None), "val", None)
    return tag if tag else None


def _contains(lang_east_asia: str | None, lang: str) -> bool:
    # eg <w:lang w:eastAsia="zh-CN" .. />
    return lang_east_asia is not None and lang in lang_east_asia


@dataclasses.dataclass(slots=True)
class _Slots:
    """The four slots and the run's hint, as one run's dispatch needs them."""

    single: str | None = None
    """Set where the whole run goes in one font: a symbol font, the cs font, the
    preamble rule."""
    east_asia: str | None = None
    ascii_font: str | None = None
    h_ansi: str | None = None
    cs: str | None = None
    hint: Any = None
    lang_east_asia: str | None = None


class RunFontSelector:
    """The document font of each character of a run.

    Built over a source: the package's property resolver, its theme part's
    ``a:theme``, ``w:themeFontLang`` from the settings part, which Office theme
    answers a reference where there is no theme part, and the emoji font.
    """

    __slots__ = (
        "_default_font",
        "default_theme_value",
        "emoji_font",
        "resolver",
        "theme",
        "theme_font_lang",
    )

    def __init__(
        self,
        resolver: Any = None,
        theme: Any = None,
        theme_font_lang: Any = None,
        default_theme: Any = None,
        emoji_font: str | None = None,
    ) -> None:
        """Build the selector.

        Args:
            resolver: the package's resolver; without one, the properties passed
                in are taken as they are.
            theme: the theme part's ``a:theme`` (or its font scheme); None for a
                package with no theme part.
            theme_font_lang: ``w:themeFontLang`` from the settings part.
            default_theme: which Office theme answers a reference where there is
                no theme part; the process-wide setting by default.
            emoji_font: the face for an astral emoji, docx4j's
                ``docx4j.fonts.RunFontSelector.EmojiFont``. Unset by default, as
                docx4j's property is, so an emoji takes the hAnsi font.
        """
        self.resolver = resolver
        self.theme = theme
        self.theme_font_lang = theme_font_lang
        self.default_theme_value = default_theme
        self.emoji_font = emoji_font
        self._default_font: str | None = None

    @property
    def default_theme(self) -> DefaultTheme:
        """Which Office theme answers a reference where the package has no theme part."""
        from docx4j_py.model.fonts.default_theme import default_theme_of

        return default_theme_of(self.default_theme_value) or default_theme_setting()

    def theme_font(self, reference: Any) -> str | None:
        """The face a theme reference names.

        The theme part's answer for the document's ``w:themeFontLang``; or, where
        the document has **no theme part**, the Office theme's own Latin faces ---
        Word supplies that theme to such a document, and the explicit ``w:ascii``
        beside the reference is unused (docx4j CR-016 probe
        ``fonts-missing-slots`` (b)).

        Only the two Latin slots are answered in the themeless case: the Office
        theme's East Asian and complex-script entries are empty, so those
        references resolve to nothing and the explicit attribute, if any, stands.
        """
        name = reference if reference is None else str(getattr(reference, "value", reference))
        if name is None:
            return None
        if self.theme is not None:
            return theme_font_of(self.theme, name, self.theme_font_lang)
        if name in ("minorAscii", "minorHAnsi"):
            return self.default_theme.minor_latin
        if name in ("majorAscii", "majorHAnsi"):
            return self.default_theme.major_latin
        return None

    def document_fonts_of(self, r_fonts: Any) -> list[str]:
        """The document fonts an ``w:rFonts`` names once its theme references are resolved.

        For each of the four slots, the theme's face, else the explicit
        attribute. Blank names are left out. This is what font discovery collects
        (``MainDocumentPart.fonts_in_use()``).
        """
        names: list[str] = []
        if r_fonts is None:
            return names
        refs = (
            r_fonts.ascii_theme,
            r_fonts.h_ansi_theme,
            r_fonts.east_asia_theme,
            r_fonts.cstheme,
        )
        explicit = (r_fonts.ascii, r_fonts.h_ansi, r_fonts.east_asia, r_fonts.cs)
        for index in range(4):
            name = self.theme_font(refs[index])
            if name is None:
                name = explicit[index]
            if name is not None and name.strip():
                names.append(name.strip())
        return names

    def default_font_of(self, doc_defaults_r_fonts: Any) -> str:
        """The document's default font for Latin text.

        The document defaults' ``w:ascii`` (a theme reference resolved first),
        else ``w:hAnsi``; Times New Roman where the defaults name no font at all.
        A theme reference the theme cannot answer gives Calibri, as docx4j's does.
        """
        font: str | None = None
        if doc_defaults_r_fonts is not None:
            r_fonts = doc_defaults_r_fonts
            if r_fonts.ascii_theme is not None:
                font = self.theme_font(r_fonts.ascii_theme)
            if font is None:
                font = r_fonts.ascii
            if font is None and r_fonts.ascii_theme is not None:
                # a theme reference the theme part cannot answer (an empty typeface)
                font = "Calibri"
            if font is None and r_fonts.h_ansi_theme is not None:
                font = self.theme_font(r_fonts.h_ansi_theme)
            if font is None:
                font = r_fonts.h_ansi
        return font if font is not None else BUILT_IN_DEFAULT_FONT

    @property
    def default_font(self) -> str:
        """:meth:`default_font_of` of the document defaults the resolver holds, once."""
        if self._default_font is None:
            r_fonts = None
            if self.resolver is not None:
                default_r_pr = self.resolver.get_document_default_rpr()
                r_fonts = None if default_r_pr is None else default_r_pr.r_fonts
            self._default_font = self.default_font_of(r_fonts)
        return self._default_font

    def refresh(self) -> None:
        """Forget the cached default font: the styles part changed under us."""
        self._default_font = None

    def ascii_font_name(self, r_pr: Any) -> str:
        """The document font an ``w:rPr`` asks for its ASCII text.

        ``w:ascii``, or the face ``w:asciiTheme`` names; failing those
        ``w:hAnsi``, and failing that the document default. Never empty.

        This is what a *name* is --- what Office JS's ``font.name`` reports and
        what docx4j uses for the line metrics of generated text and paragraph
        marks. It is not the dispatch: a run of Arabic in a font named here may
        still be drawn in the cs font (:meth:`document_font_for`).
        """
        r_fonts = getattr(r_pr, "r_fonts", None)
        if r_fonts is not None:
            if r_fonts.ascii_theme is not None:
                font = self.theme_font(r_fonts.ascii_theme)
                if font:
                    return font
            if r_fonts.ascii is not None:
                return r_fonts.ascii
            if r_fonts.h_ansi is not None:
                return r_fonts.h_ansi
        return self.default_font

    def _effective_r_pr(self, p_pr: Any, r_pr: Any, r_pr_is_effective: bool) -> Any:
        """The run's effective properties, never None."""
        from docx4j_py.wml import RPr

        if not r_pr_is_effective and self.resolver is not None:
            return self.resolver.get_effective_rpr(r_pr, p_pr) or RPr()
        return r_pr if r_pr is not None else RPr()

    def _r_fonts_of(self, r_pr: Any) -> Any:
        """The run's ``w:rFonts``, or the document default for the two Latin slots.

        Where nothing anywhere names a font: the East Asian and complex-script
        slots are left unnamed as for any run that never named them, and then the
        range dispatch runs as for any other run. (Not the eastAsia slot: a Times
        New Roman there would fire the preamble rule and set the whole run in one
        span.)
        """
        from docx4j_py.wml import RFonts

        r_fonts = getattr(r_pr, "r_fonts", None)
        if r_fonts is not None:
            return r_fonts
        return RFonts(ascii=self.default_font, h_ansi=self.default_font)

    def complex_script_font(self, r_pr: Any, r_fonts: Any) -> str | None:
        """The complex-script font the run is set in as a whole.

        Where ``w:cs`` or ``w:rtl`` is **on** and the cs slot names a font (the
        theme reference resolved first, else the attribute); None where the run
        is dispatched by range instead --- neither is on, or the reference
        resolves to nothing and there is no ``w:cs``.
        """
        if not is_on(r_pr.cs) and not is_on(r_pr.rtl):
            return None
        if r_fonts.cstheme is not None:
            answer = self.theme_font(r_fonts.cstheme)
            return answer if answer is not None else r_fonts.cs
        return r_fonts.cs  # None: no CS value, dispatched by range

    def resolved_slots(self, r_fonts: Any) -> tuple[str | None, str | None, str | None, str | None]:
        """The four slots resolved, in the order eastAsia, ascii, hAnsi, cs.

        Each the theme reference where there is one, else the explicit attribute
        (which also stands where the reference resolves to nothing).
        """
        east_asia = self.theme_font(r_fonts.east_asia_theme)
        if east_asia is None:
            east_asia = r_fonts.east_asia
        ascii_font = self.theme_font(r_fonts.ascii_theme)
        if ascii_font is None:
            ascii_font = r_fonts.ascii
        h_ansi = self.theme_font(r_fonts.h_ansi_theme)
        if h_ansi is None:
            h_ansi = r_fonts.h_ansi
        cs = self.theme_font(r_fonts.cstheme)
        if cs is None:
            cs = r_fonts.cs
        # eg LibreOffice writes w:cs="" in w:docDefaults
        if cs is not None and not cs.strip():
            cs = None
        return east_asia, ascii_font, h_ansi, cs

    def _slots_of(self, p_pr: Any, r_pr: Any) -> _Slots:
        """The run's slots, ready for :meth:`font_for`."""
        r_fonts = self._r_fonts_of(r_pr)

        symbol = symbol_font_name(r_fonts.h_ansi)
        if symbol is not None:
            return _Slots(single=symbol)

        cs_font = self.complex_script_font(r_pr, r_fonts)
        if cs_font is not None:
            return _Slots(single=cs_font)

        east_asia, ascii_font, h_ansi, cs = self.resolved_slots(r_fonts)
        if preamble_rule(east_asia, ascii_font, h_ansi):
            return _Slots(single=ascii_font)

        return _Slots(
            east_asia=east_asia,
            ascii_font=ascii_font if ascii_font is not None else self.default_font,
            h_ansi=h_ansi if h_ansi is not None else self.default_font,
            cs=cs,
            hint=r_fonts.hint,
            lang_east_asia=getattr(getattr(r_pr, "lang", None), "east_asia", None),
        )

    def _font_of(self, slots: _Slots, code_point: int) -> str | None:
        if slots.single is not None:
            return slots.single
        return self.font_for(
            code_point,
            slots.hint,
            slots.lang_east_asia,
            slots.east_asia,
            slots.ascii_font,
            slots.h_ansi,
            slots.cs,
        )

    def document_font_for(
        self, p_pr: Any, r_pr: Any, code_point: int, r_pr_is_effective: bool = False
    ) -> str:
        """The document font for one code point in a run with these properties.

        The symbol font by its canonical name where the run's hAnsi is one, the
        complex-script font where ``w:cs`` or ``w:rtl`` is on and names one, the
        ascii font under the preamble rule, else the font the character-range
        table gives the code point. Never None: the document default at least.
        """
        effective = self._effective_r_pr(p_pr, r_pr, r_pr_is_effective)
        answer = self._font_of(self._slots_of(p_pr, effective), code_point)
        return answer if answer is not None else self.default_font

    def spans(
        self, p_pr: Any, r_pr: Any, text: str | None, r_pr_is_effective: bool = False
    ) -> list[FontSpan]:
        """The run's text folded into maximal stretches of one document font.

        Every code point is asked for its font (:meth:`document_font_for`) and
        joins the span being built while the answer is the same, whatever its
        range --- docx4j's walk since 17.1.1, which replaced cutting by *range* (a
        Latin run was cut at every exception character, a space went to the East
        Asian font of the word before it, and the ASCII letters after an accented
        one stayed in the hAnsi font). ``bold``, ``italic``, ``cs`` and ``rtl``
        are the effective run properties' and are the same for every span of a
        run.
        """
        out: list[FontSpan] = []
        if not text:
            return out

        effective = self._effective_r_pr(p_pr, r_pr, r_pr_is_effective)
        bold = is_on(effective.b)
        italic = is_on(effective.i)
        cs = is_on(effective.cs)
        rtl = is_on(effective.rtl)
        slots = self._slots_of(p_pr, effective)

        memo: dict[int, str | None] = {}
        current: list[str] = []
        current_font: str | None = None
        started = False
        current_script: str | None = None
        for character in text:
            code_point = ord(character)
            if code_point in memo:
                font = memo[code_point]
            else:
                font = self._font_of(slots, code_point)
                memo[code_point] = font
            if started and current_font != font:
                out.append(
                    FontSpan("".join(current), current_font, bold, italic, cs, rtl, current_script)
                )
                current = []
                current_script = None
            current_font = font
            started = True
            script = span_script(code_point)
            if script is not None and current_script is None:
                current_script = script
            current.append(character)
        if started:
            out.append(
                FontSpan("".join(current), current_font, bold, italic, cs, rtl, current_script)
            )
        return out

    def font_for(
        self,
        code_point: int,
        hint: Any,
        lang_east_asia: str | None,
        east_asia: str | None,
        ascii_font: str | None,
        h_ansi: str | None,
        cs: str | None,
    ) -> str | None:
        """The document font which formats one code point, given the run's four fonts.

        Each already resolved through the theme, plus its ``w:hint``: the table in
        [MS-OI29500] 17.3.2.26, with these departures from it, each measured by
        docx4j:

        * the Indic, Thai, Lao, Myanmar and Khmer ranges use cs, where the table
          says hAnsi (docx4j issues 622 and 666);
        * an East Asian reference is "eastAsia (or eastAsiaTheme) if defined": a
          run naming none gets hAnsi in those places, never nothing;
        * characters outside the Basic Multilingual Plane, which the table does
          not cover: the emoji font where one is configured, else hAnsi (Word uses
          Segoe UI Emoji, which docx4j's coverage pass finds).

        What the table says and Word confirmed: U+0020 is ASCII and takes the
        ascii font, so the space between two East Asian words is the Latin font's;
        Hebrew, Arabic and the rest of U+0590 to U+07BF in a run with no ``w:cs``
        take the ascii font. The table's other condition for U+0100 to U+02AF and
        U+1E00 to U+1EFF under ``hint=eastAsia``, "or the character set of the
        eastAsia font is Chinese5 or GB2312", needs the font's OS/2 code-page bits
        and is not implemented here either.

        **U+2190 to U+2BFF**: docx4j asks whether the hAnsi font has the glyph and
        names Segoe UI Symbol where it does not. That is a glyph check over a font
        file, which this phase does not port (CR-002 section 6.3), so the hAnsi
        font is answered --- which is also docx4j's answer wherever Segoe UI
        Symbol is not among the fonts it has, the environment the parity goldens
        record.

        Returns:
            A document font name; None only where the run names no hAnsi font.
        """
        # "eastAsia (or eastAsiaTheme if defined)": where the run names none, hAnsi
        ea = east_asia if east_asia is not None else h_ansi
        hint_value = None if hint is None else str(getattr(hint, "value", hint))
        hint_ea = hint_value == "eastAsia"
        c = code_point

        if c > 0xFFFF:
            if is_emoji(c) and self.emoji_font is not None:
                return self.emoji_font
            return h_ansi
        # Basic Latin: "@ascii is used to format all characters in the ASCII range (0-127)"
        if c <= 0x7F:
            return ascii_font
        # Latin-1 Supplement: hAnsi, with the table's exceptions under hint=eastAsia
        if 0xA0 <= c <= 0xFF:
            if hint_ea and east_asia is not None:
                # "If hint is eastAsia, the following characters use eastAsia: A1,
                #  A4, A7-A8, AA, AD, AF, B0-B4, B6-BA, BC-BF, D7, F7"
                if (
                    c in (0xA1, 0xA4, 0xAA, 0xAD, 0xAF, 0xD7, 0xF7)
                    or 0xA7 <= c <= 0xA8
                    or 0xB0 <= c <= 0xB4
                    or 0xB6 <= c <= 0xBA
                    or 0xBC <= c <= 0xBF
                ):
                    return east_asia
                # "If hint is eastAsia and the language of the run is either Chinese
                #  Traditional or Chinese Simplified, the following characters use
                #  eastAsia: E0-E1, E8-EA, EC-ED, F2-F3, F9-FA, FC"
                if _contains(lang_east_asia, "zh") and (
                    0xE0 <= c <= 0xE1
                    or 0xE8 <= c <= 0xEA
                    or 0xEC <= c <= 0xED
                    or 0xF2 <= c <= 0xF3
                    or 0xF9 <= c <= 0xFA
                    or c == 0xFC
                ):
                    return east_asia
            return h_ansi
        # Latin Extended-A, Latin Extended-B, IPA Extensions
        if 0x100 <= c <= 0x2AF:
            return ea if (hint_ea and _contains(lang_east_asia, "zh")) else h_ansi
        # Spacing Modifier Letters .. Cyrillic Supplement
        if 0x2B0 <= c <= 0x4FF:
            return ea if hint_ea else h_ansi
        # Hebrew, Arabic, Syriac, Thaana, NKo: ascii (measured, fonts-hebrew-no-cs)
        if 0x590 <= c <= 0x7BF:
            return ascii_font
        # the Indic scripts, Thai, Lao, Myanmar, Khmer and the Khmer symbols (622, 666)
        if (
            0x900 <= c <= 0xDFF
            or 0xE00 <= c <= 0xEFF
            or 0x1000 <= c <= 0x109F
            or 0x1780 <= c <= 0x17FF
            or 0x19E0 <= c <= 0x19FF
        ):
            return cs if cs is not None else h_ansi
        # Hangul Jamo
        if 0x1100 <= c <= 0x11FF:
            return ea
        # Latin Extended Additional
        if 0x1E00 <= c <= 0x1EFF:
            return ea if (hint_ea and _contains(lang_east_asia, "zh")) else h_ansi
        # General Punctuation .. Number Forms: ordinary text
        if 0x2000 <= c <= 0x218F:
            return ea if hint_ea else h_ansi
        # the symbol blocks: Arrows .. Braille (docx4j glyph-checks these)
        if 0x2190 <= c <= 0x2BFF:
            return ea if hint_ea else h_ansi
        # Glagolitic .. CJK Radicals Supplement
        if 0x2C00 <= c <= 0x2EFF:
            return ea if hint_ea else h_ansi
        # the CJK ranges (and the surrogate range, which a paired surrogate never reaches)
        if 0x2F00 <= c <= 0xDFFF:
            return ea
        # Private Use Area; F000-F0FF is where the symbol fonts live
        if 0xE000 <= c <= 0xF8FF:
            return ea if hint_ea else h_ansi
        # CJK Compatibility Ideographs
        if 0xF900 <= c <= 0xFAFF:
            return ea
        # Alphabetic Presentation Forms: "If the hint is eastAsia then eastAsia is
        # used for FB00-FB1C.  For FB1D-FB4F, ascii is used."
        if 0xFB00 <= c <= 0xFB4F:
            if c >= 0xFB1D:
                return ascii_font
            return ea if hint_ea else h_ansi
        # Arabic Presentation Forms-A
        if 0xFB50 <= c <= 0xFDFF:
            return ascii_font
        # CJK Compatibility Forms, Small Form Variants
        if 0xFE30 <= c <= 0xFE6F:
            return ea
        # Arabic Presentation Forms-B
        if 0xFE70 <= c <= 0xFEFE:
            return ascii_font
        # Halfwidth and Fullwidth Forms
        if 0xFF00 <= c <= 0xFFEF:
            return ea
        # "for all ranges not listed in the above, the hAnsi font shall be used"
        # (Georgian, Armenian, Ethiopic, Tibetan, Mongolian, Greek Extended, ...)
        return h_ansi

    def __repr__(self) -> str:
        """``<RunFontSelector default=Aptos theme=yes>``."""
        return (
            f"<RunFontSelector default={self.default_font!r} "
            f"theme={'yes' if self.theme is not None else 'no'}>"
        )
