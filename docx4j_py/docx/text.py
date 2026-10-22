"""``Paragraph``, ``Run``, ``Font``, ``ParagraphFormat``, ``ColorFormat``, ``Hyperlink``.

CR-003 section 3.6, Phase J: python-docx's text proxies, with python-docx's
semantics, over CR-001's typed tree and CR-003's content API.

**python-docx's text, not the text model's** (CR-011 Q5; WD REQ-202's exemption
for a reader of an emulated API). ``Paragraph.text`` is its ``w:r`` and
``w:hyperlink`` children's text in order, as python-docx 1.2.0 reads it: a
``w:ins`` and a ``w:del`` both skipped, no content control, a break nothing but
``\n``. The text model is :func:`docx4j_py.model.text.extract`; each of its runs
records its source ``w:r``, which is the facade's ``Run`` it lies in, or none.

Two of python-docx's habits shape this module.

**Tri-state.** ``run.bold`` is ``True``, ``False`` or ``None``, and ``None``
means "inherited" while ``False`` means "off, whatever the style says". The
content API's :class:`~docx4j_py.model.content.Font` is Office JS's, which is
two-state --- ``bold = False`` *removes* ``w:b`` --- so every tri-state boolean
here reads and writes the typed ``w:rPr`` itself, writing ``<w:b w:val="0"/>``
for ``False`` and removing the element for ``None``. Section 20 records the
departure; the complex-script twins the content API writes beside ``w:b`` and
``w:i`` (CR-003 section 3.12) are written here too, so a run formatted through
either surface reads back the same through both.

**Points, not EMU.** The content API measures in points; python-docx in EMU
through :class:`~docx4j_py.docx.shared.Length`. Every number crossing this
module is converted, and the conversion is python-docx's own rounding.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import IO, TYPE_CHECKING, Any

from docx4j_py.child import ChildList
from docx4j_py.docx._base import Story, recording
from docx4j_py.docx.enums import (
    ALIGNMENT_TO_OFFICE_JS,
    MSO_COLOR_TYPE,
    MSO_THEME_COLOR,
    WD_BREAK,
    WD_COLOR_INDEX,
    WD_LINE_SPACING,
    WD_PARAGRAPH_ALIGNMENT,
    WD_STYLE_TYPE,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
    WD_UNDERLINE,
)
from docx4j_py.docx.shared import Emu, Length, Pt, RGBColor, Twips, points_of, twips_of
from docx4j_py.model.content.paragraph import Paragraph as ContentParagraph
from docx4j_py.wml import CTTabStop, P, PPr, R, RPr, el
from docx4j_py.wml import r as r_builder

if TYPE_CHECKING:  # pragma: no cover - annotations only
    from docx4j_py.docx.shape import InlineShape
    from docx4j_py.docx.styles import CharacterStyle, ParagraphStyle

__all__ = [
    "ColorFormat",
    "Font",
    "Hyperlink",
    "Paragraph",
    "ParagraphFormat",
    "RenderedPageBreak",
    "Run",
]

#: python-docx's ``Font`` boolean, the ``w:rPr`` field it is, and the
#: complex-script twin CR-003 section 3.12 keeps beside it.
BOOL_PROPERTIES: dict[str, tuple[str, str | None]] = {
    "all_caps": ("caps", None),
    "bold": ("b", "b_cs"),
    "complex_script": ("cs", None),
    "cs_bold": ("b_cs", None),
    "cs_italic": ("i_cs", None),
    "double_strike": ("dstrike", None),
    "emboss": ("emboss", None),
    "hidden": ("vanish", None),
    "imprint": ("imprint", None),
    "italic": ("i", "i_cs"),
    "math": ("o_math", None),
    "no_proof": ("no_proof", None),
    "outline": ("outline", None),
    "rtl": ("rtl", None),
    "shadow": ("shadow", None),
    "small_caps": ("small_caps", None),
    "snap_to_grid": ("snap_to_grid", None),
    "spec_vanish": ("spec_vanish", None),
    "strike": ("strike", None),
    "web_hidden": ("web_hidden", None),
}

#: ``w:rPr`` field -> the ``el`` factory that builds it, for the boolean toggles.
_BOOL_ELEMENTS: dict[str, str] = {
    "b": "b",
    "b_cs": "bCs",
    "caps": "caps",
    "cs": "cs",
    "dstrike": "dstrike",
    "emboss": "emboss",
    "i": "i",
    "i_cs": "iCs",
    "imprint": "imprint",
    "no_proof": "noProof",
    "o_math": "oMath",
    "outline": "outline",
    "rtl": "rtl",
    "shadow": "shadow",
    "small_caps": "smallCaps",
    "snap_to_grid": "snapToGrid",
    "spec_vanish": "specVanish",
    "strike": "strike",
    "vanish": "vanish",
    "web_hidden": "webHidden",
}

_HIGHLIGHT_BY_XML = {member.xml_value: member for member in WD_COLOR_INDEX}

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: What a run child contributes to ``Run.text``, as python-docx reads it: a
#: ``w:br`` with a ``page`` or ``column`` type contributes **nothing**, which is
#: where python-docx and the content API's ``text_of`` part company (the content
#: API gives every ``w:br`` a newline). The facade follows python-docx, because
#: ``Run.text`` is python-docx's member (section 20, departure 6).
_RUN_TEXT_ITEMS = {
    f"{_W}tab": "\t",
    f"{_W}cr": "\n",
    f"{_W}noBreakHyphen": "\u2011",
    f"{_W}softHyphen": "\u00ad",
}


def run_text(run: Any) -> str:
    """The text of one ``w:r``, as python-docx's ``Run.text`` reports it."""
    from docx4j_py.model.content.text_model import item_text
    from docx4j_py.traversal import element_name

    out: list[str] = []
    for item in getattr(run, "content", None) or ():
        name = element_name(item)
        if name == f"{_W}br":
            kind = getattr(item, "type_value", None)
            kind = getattr(kind, "value", kind)
            out.append("" if kind in ("page", "column") else "\n")
        elif name in _RUN_TEXT_ITEMS:
            out.append(_RUN_TEXT_ITEMS[name])
        elif name == f"{_W}t" or name == f"{_W}sym":
            out.append(item_text(name, item) or "")
    return "".join(out)


def check_python_docx_text(text: str | None) -> None:
    """Refuse a text python-docx 1.2.0 refuses: one holding a character XML 1.0 excludes.

    Measured on 2026-10-23 (CR-013 Phase A, section 4.1): python-docx's
    ``add_paragraph``, ``add_run``, ``Run.text = ...``, ``Paragraph.text =
    ...``, ``cell.text = ...`` and ``add_comment`` raise ``ValueError`` (lxml's
    "All strings must be XML compatible") at the write for any control
    character but a tab, a newline and a carriage return --- Office JS's
    U+000B, U+000C, U+000E, U+001E, U+001F and U+0002 among them. The facade
    is python-docx's names and semantics (CR-003 Phase J), so it refuses them
    there too, before anything changes, where the content API's writers map
    them (CR-013 Q2).

    Raises:
        ValueError: the text holds such a character.
    """
    if not text:
        return
    for index, char in enumerate(text):
        code = ord(char)
        if (code < 0x20 and char not in "\t\n\r") or 0xD800 <= code <= 0xDFFF or code in (0xFFFE, 0xFFFF):
            raise ValueError(
                "All strings must be XML compatible: Unicode or ASCII, no NULL bytes or control "
                f"characters (U+{code:04X} at offset {index}, as python-docx 1.2.0 refuses it)"
            )


def holder_text(holder: Any) -> str:
    """The text of a ``w:p`` or ``w:hyperlink``: its runs and its links.

    python-docx's ``CT_P.text``, which is the runs and the hyperlinks and
    nothing else --- not a ``w:ins``, not a content control.
    """
    from docx4j_py.traversal import element_name

    out: list[str] = []
    for item in getattr(holder, "content", None) or ():
        if isinstance(item, R):
            out.append(run_text(item))
        elif element_name(item) == f"{_W}hyperlink":
            out.append(holder_text(item))
    return "".join(out)


def _r_pr(holder: Any, *, create: bool) -> RPr | None:
    """The holder's ``w:rPr``, made if `create` and absent."""
    r_pr = getattr(holder, "r_pr", None)
    if r_pr is None and create:
        r_pr = RPr()
        holder.r_pr = r_pr
        r_pr.parent = holder
    return r_pr


def _toggle(field: str, value: bool) -> Any:
    """``<w:b/>`` for True, ``<w:b w:val="0"/>`` for False."""
    factory = getattr(el, _BOOL_ELEMENTS[field])
    return factory() if value else factory(val=False)


# ---------------------------------------------------------------------------
# Font and ColorFormat
# ---------------------------------------------------------------------------


class ColorFormat:
    """``font.color``: ``w:rPr/w:color`` as python-docx reads and writes it."""

    __slots__ = ("_holder", "_story")

    def __init__(self, holder: Any, story: Story | None) -> None:
        """Wrap the ``w:r`` or ``w:style`` whose ``w:rPr`` carries the colour."""
        self._holder = holder
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one holder are equal."""
        return isinstance(other, ColorFormat) and other._holder is self._holder

    def __hash__(self) -> int:
        """Hashes by the holder's identity."""
        return hash(id(self._holder))

    @property
    def element(self) -> Any:
        """The typed element whose ``w:rPr/w:color`` this is over."""
        return self._holder

    @property
    def part(self) -> Any:
        """The part this colour is in, or None for a colour on a style."""
        return None if self._story is None else self._story.part

    @property
    def _color(self) -> Any:
        r_pr = _r_pr(self._holder, create=False)
        return getattr(r_pr, "color", None) if r_pr is not None else None

    @staticmethod
    def _val(color: Any) -> str | None:
        value = getattr(color, "val", None)
        if isinstance(value, bytes):
            return value.hex().upper()
        value = getattr(value, "value", value)
        return None if value is None else str(value)

    @property
    def rgb(self) -> RGBColor | None:
        """The colour as an :class:`~docx4j_py.docx.shared.RGBColor`, or None."""
        value = self._val(self._color)
        if value is None or value.lower() == "auto":
            return None
        try:
            return RGBColor.from_string(value)
        except (ValueError, IndexError):
            return None

    @rgb.setter
    def rgb(self, value: RGBColor | None) -> None:
        if value is None and self._color is None:
            return
        with _write(self._story, "docx.font.color"):
            r_pr = _r_pr(self._holder, create=True)
            assert r_pr is not None
            if value is None:
                r_pr.color = None
                return
            r_pr.color = el.color(val=str(value))
            r_pr.color.parent = r_pr

    @property
    def theme_color(self) -> Any:
        """``w:color/@themeColor`` as an ``MSO_THEME_COLOR`` member, or None."""
        theme = getattr(self._color, "theme_color", None)
        theme = getattr(theme, "value", theme)
        if theme is None:
            return None
        try:
            return MSO_THEME_COLOR.from_xml(str(theme))
        except ValueError:
            return None

    @theme_color.setter
    def theme_color(self, value: Any) -> None:
        with _write(self._story, "docx.font.color"):
            if value is None:
                r_pr = _r_pr(self._holder, create=False)
                if r_pr is not None:
                    r_pr.color = None
                return
            r_pr = _r_pr(self._holder, create=True)
            assert r_pr is not None
            color = r_pr.color
            if color is None:
                color = el.color(val="auto")
                r_pr.color = color
                color.parent = r_pr
            color.theme_color = MSO_THEME_COLOR.to_xml(value)

    @property
    def type(self) -> Any:
        """``MSO_COLOR_TYPE.RGB``, ``.THEME``, ``.AUTO`` or None."""
        color = self._color
        if color is None:
            return None
        if getattr(color, "theme_color", None) is not None:
            return MSO_COLOR_TYPE.THEME
        value = self._val(color)
        if value is not None and value.lower() == "auto":
            return MSO_COLOR_TYPE.AUTO
        return MSO_COLOR_TYPE.RGB


class _write:
    """``with _write(story, op):`` --- a recording block, or nothing at all.

    Every facade mutation records a
    :class:`~docx4j_py.model.content.reports.ChangeReport`, which needs the
    content API body; a ``Font`` over a *style* has no body, and then this is a
    no-op context manager.
    """

    __slots__ = ("inner",)

    def __init__(self, story: Story | None, operation: str) -> None:
        """Open a report when there is a body to open one on."""
        self.inner = None if story is None else recording(story.body, operation)

    def __enter__(self) -> Any:
        """The recorder, or None."""
        return None if self.inner is None else self.inner.__enter__()

    def __exit__(self, *exc_info: object) -> None:
        """Close the report, if there was one."""
        if self.inner is not None:
            self.inner.__exit__(*exc_info)


class Font:
    """``run.font`` and ``style.font``: character formatting, tri-state.

    Over the ``w:rPr`` of a ``w:r`` or a ``w:style``, as python-docx's is. Every
    property is ``None`` when the run says nothing about it, which is what
    python-docx means by "inherited from the style hierarchy"; CR-002 Phase B's
    ``PropertyResolver`` is what would resolve it, and until then this reports
    direct formatting, exactly as the content API's ``Font`` does (CR-003
    section 6).
    """

    __slots__ = ("_holder", "_story")

    def __init__(self, holder: Any, story: Story | None = None) -> None:
        """Wrap a ``w:r`` or ``w:style``."""
        self._holder = holder
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one holder are equal."""
        return isinstance(other, Font) and other._holder is self._holder

    def __hash__(self) -> int:
        """Hashes by the holder's identity."""
        return hash(id(self._holder))

    def __repr__(self) -> str:
        """``<docx4j_py.docx Font 'Calibri' 11.0pt>``."""
        size = self.size
        return (
            f"<docx4j_py.docx Font {self.name!r} "
            f"{'inherited' if size is None else f'{size.pt}pt'}>"
        )

    @property
    def element(self) -> Any:
        """The typed element whose ``w:rPr`` this font is over."""
        return self._holder

    @property
    def part(self) -> Any:
        """The part this font is in, or None for a font on a style.

        python-docx raises here for a style's font, because its style objects
        carry no parent; None is the kinder answer and the reads and writes
        work either way.
        """
        return None if self._story is None else self._story.part

    # -- the tri-state booleans -------------------------------------------

    def _get_bool(self, name: str) -> bool | None:
        field, _twin = BOOL_PROPERTIES[name]
        r_pr = _r_pr(self._holder, create=False)
        if r_pr is None:
            return None
        toggle = getattr(r_pr, field, None)
        if toggle is None:
            return None
        value = getattr(toggle, "val", None)
        return True if value is None else bool(value)

    def _set_bool(self, name: str, value: bool | None) -> None:
        field, twin = BOOL_PROPERTIES[name]
        if value is None and _r_pr(self._holder, create=False) is None:
            return
        with _write(self._story, f"docx.font.{name}"):
            r_pr = _r_pr(self._holder, create=True)
            assert r_pr is not None
            fields = (field,) if twin is None else (field, twin)
            for each in fields:
                if value is None:
                    setattr(r_pr, each, None)
                else:
                    toggle = _toggle(each, bool(value))
                    setattr(r_pr, each, toggle)
                    toggle.parent = r_pr

    @property
    def all_caps(self) -> bool | None:
        """``w:caps``: the text appears in capitals."""
        return self._get_bool("all_caps")

    @all_caps.setter
    def all_caps(self, value: bool | None) -> None:
        self._set_bool("all_caps", value)

    @property
    def bold(self) -> bool | None:
        """``w:b`` (and ``w:bCs``): the text is bold."""
        return self._get_bool("bold")

    @bold.setter
    def bold(self, value: bool | None) -> None:
        self._set_bool("bold", value)

    @property
    def complex_script(self) -> bool | None:
        """``w:cs``: treat the characters as complex script."""
        return self._get_bool("complex_script")

    @complex_script.setter
    def complex_script(self, value: bool | None) -> None:
        self._set_bool("complex_script", value)

    @property
    def cs_bold(self) -> bool | None:
        """``w:bCs``: the complex-script characters are bold."""
        return self._get_bool("cs_bold")

    @cs_bold.setter
    def cs_bold(self, value: bool | None) -> None:
        self._set_bool("cs_bold", value)

    @property
    def cs_italic(self) -> bool | None:
        """``w:iCs``: the complex-script characters are italic."""
        return self._get_bool("cs_italic")

    @cs_italic.setter
    def cs_italic(self, value: bool | None) -> None:
        self._set_bool("cs_italic", value)

    @property
    def double_strike(self) -> bool | None:
        """``w:dstrike``: double strikethrough."""
        return self._get_bool("double_strike")

    @double_strike.setter
    def double_strike(self, value: bool | None) -> None:
        self._set_bool("double_strike", value)

    @property
    def emboss(self) -> bool | None:
        """``w:emboss``: raised off the page in relief."""
        return self._get_bool("emboss")

    @emboss.setter
    def emboss(self, value: bool | None) -> None:
        self._set_bool("emboss", value)

    @property
    def hidden(self) -> bool | None:
        """``w:vanish``: hidden from display."""
        return self._get_bool("hidden")

    @hidden.setter
    def hidden(self, value: bool | None) -> None:
        self._set_bool("hidden", value)

    @property
    def imprint(self) -> bool | None:
        """``w:imprint``: pressed into the page."""
        return self._get_bool("imprint")

    @imprint.setter
    def imprint(self, value: bool | None) -> None:
        self._set_bool("imprint", value)

    @property
    def italic(self) -> bool | None:
        """``w:i`` (and ``w:iCs``): the text is italic."""
        return self._get_bool("italic")

    @italic.setter
    def italic(self, value: bool | None) -> None:
        self._set_bool("italic", value)

    @property
    def math(self) -> bool | None:
        """``w:oMath``: the run is Office Open XML Math."""
        return self._get_bool("math")

    @math.setter
    def math(self, value: bool | None) -> None:
        self._set_bool("math", value)

    @property
    def no_proof(self) -> bool | None:
        """``w:noProof``: not spell- or grammar-checked."""
        return self._get_bool("no_proof")

    @no_proof.setter
    def no_proof(self, value: bool | None) -> None:
        self._set_bool("no_proof", value)

    @property
    def outline(self) -> bool | None:
        """``w:outline``: each glyph outlined."""
        return self._get_bool("outline")

    @outline.setter
    def outline(self, value: bool | None) -> None:
        self._set_bool("outline", value)

    @property
    def rtl(self) -> bool | None:
        """``w:rtl``: right-to-left characteristics."""
        return self._get_bool("rtl")

    @rtl.setter
    def rtl(self, value: bool | None) -> None:
        self._set_bool("rtl", value)

    @property
    def shadow(self) -> bool | None:
        """``w:shadow``: each character shadowed."""
        return self._get_bool("shadow")

    @shadow.setter
    def shadow(self, value: bool | None) -> None:
        self._set_bool("shadow", value)

    @property
    def small_caps(self) -> bool | None:
        """``w:smallCaps``: lower case shown as small capitals."""
        return self._get_bool("small_caps")

    @small_caps.setter
    def small_caps(self, value: bool | None) -> None:
        self._set_bool("small_caps", value)

    @property
    def snap_to_grid(self) -> bool | None:
        """``w:snapToGrid``: lay out on the document grid."""
        return self._get_bool("snap_to_grid")

    @snap_to_grid.setter
    def snap_to_grid(self, value: bool | None) -> None:
        self._set_bool("snap_to_grid", value)

    @property
    def spec_vanish(self) -> bool | None:
        """``w:specVanish``: always hidden (§17.3.2.36)."""
        return self._get_bool("spec_vanish")

    @spec_vanish.setter
    def spec_vanish(self, value: bool | None) -> None:
        self._set_bool("spec_vanish", value)

    @property
    def strike(self) -> bool | None:
        """``w:strike``: a single line through the text."""
        return self._get_bool("strike")

    @strike.setter
    def strike(self, value: bool | None) -> None:
        self._set_bool("strike", value)

    @property
    def web_hidden(self) -> bool | None:
        """``w:webHidden``: hidden in web-page view."""
        return self._get_bool("web_hidden")

    @web_hidden.setter
    def web_hidden(self, value: bool | None) -> None:
        self._set_bool("web_hidden", value)

    # -- sub- and superscript ---------------------------------------------

    def _vert_align(self) -> str | None:
        r_pr = _r_pr(self._holder, create=False)
        value = getattr(getattr(r_pr, "vert_align", None), "val", None)
        value = getattr(value, "value", value)
        return None if value is None else str(value)

    def _set_vert_align(self, wanted: str, value: bool | None) -> None:
        current = self._vert_align()
        if value is None or (not value and current == wanted):
            if current is None:
                return
            with _write(self._story, f"docx.font.{wanted}"):
                r_pr = _r_pr(self._holder, create=False)
                if r_pr is not None:
                    r_pr.vert_align = None
            return
        if not value:
            return
        with _write(self._story, f"docx.font.{wanted}"):
            r_pr = _r_pr(self._holder, create=True)
            assert r_pr is not None
            r_pr.vert_align = el.vertAlign(val=wanted)
            r_pr.vert_align.parent = r_pr

    @property
    def subscript(self) -> bool | None:
        """``w:vertAlign="subscript"``; None when there is no ``w:vertAlign``."""
        value = self._vert_align()
        return None if value is None else value == "subscript"

    @subscript.setter
    def subscript(self, value: bool | None) -> None:
        self._set_vert_align("subscript", value)

    @property
    def superscript(self) -> bool | None:
        """``w:vertAlign="superscript"``; None when there is no ``w:vertAlign``."""
        value = self._vert_align()
        return None if value is None else value == "superscript"

    @superscript.setter
    def superscript(self, value: bool | None) -> None:
        self._set_vert_align("superscript", value)

    # -- the rest ----------------------------------------------------------

    @property
    def color(self) -> ColorFormat:
        """``font.color``: the :class:`ColorFormat` over ``w:color``."""
        return ColorFormat(self._holder, self._story)

    @property
    def highlight_color(self) -> Any:
        """``w:highlight`` as a ``WD_COLOR_INDEX`` member, or None."""
        r_pr = _r_pr(self._holder, create=False)
        value = getattr(getattr(r_pr, "highlight", None), "val", None)
        value = getattr(value, "value", value)
        if value is None:
            return None
        return _HIGHLIGHT_BY_XML.get(str(value))

    @highlight_color.setter
    def highlight_color(self, value: Any) -> None:
        with _write(self._story, "docx.font.highlight_color"):
            if value is None or value == WD_COLOR_INDEX.INHERITED:
                r_pr = _r_pr(self._holder, create=False)
                if r_pr is not None:
                    r_pr.highlight = None
                return
            name = WD_COLOR_INDEX(value).xml_value
            r_pr = _r_pr(self._holder, create=True)
            assert r_pr is not None
            r_pr.highlight = el.highlight(val=name)
            r_pr.highlight.parent = r_pr

    @property
    def name(self) -> str | None:
        """``w:rFonts/@ascii``: the typeface name, or None."""
        r_pr = _r_pr(self._holder, create=False)
        fonts = getattr(r_pr, "r_fonts", None) if r_pr is not None else None
        return None if fonts is None else (fonts.ascii or None)

    @name.setter
    def name(self, value: str | None) -> None:
        with _write(self._story, "docx.font.name"):
            r_pr = _r_pr(self._holder, create=value is not None)
            if r_pr is None:
                return
            if value is None:
                r_pr.r_fonts = None
                return
            fonts = r_pr.r_fonts
            if fonts is None:
                fonts = el.rFonts()
                r_pr.r_fonts = fonts
                fonts.parent = r_pr
            fonts.ascii = value
            fonts.h_ansi = value
            fonts.ascii_theme = None
            fonts.h_ansi_theme = None

    @property
    def size(self) -> Length | None:
        """``w:sz``, in half-points, as a :class:`Length`; None when unset."""
        r_pr = _r_pr(self._holder, create=False)
        value = getattr(getattr(r_pr, "sz", None), "val", None)
        return None if value is None else Pt(float(value) / 2)

    @size.setter
    def size(self, value: int | Length | None) -> None:
        with _write(self._story, "docx.font.size"):
            r_pr = _r_pr(self._holder, create=value is not None)
            if r_pr is None:
                return
            if value is None:
                r_pr.sz = None
                r_pr.sz_cs = None
                return
            half = round(Length(int(value)).pt * 2)
            for field, factory in (("sz", el.sz), ("sz_cs", el.szCs)):
                toggle = factory(val=half)
                setattr(r_pr, field, toggle)
                toggle.parent = r_pr

    @property
    def underline(self) -> bool | WD_UNDERLINE | None:
        """True, False, None or a ``WD_UNDERLINE`` member, as python-docx reports it."""
        r_pr = _r_pr(self._holder, create=False)
        u = getattr(r_pr, "u", None) if r_pr is not None else None
        if u is None:
            return None
        value = getattr(u, "val", None)
        value = getattr(value, "value", value)
        if value is None:
            return None
        try:
            member = WD_UNDERLINE.from_xml(str(value))
        except ValueError:
            return None
        if member is WD_UNDERLINE.SINGLE:
            return True
        if member is WD_UNDERLINE.NONE:
            return False
        return member

    @underline.setter
    def underline(self, value: bool | WD_UNDERLINE | None) -> None:
        with _write(self._story, "docx.font.underline"):
            r_pr = _r_pr(self._holder, create=value is not None)
            if r_pr is None:
                return
            if value is None:
                r_pr.u = None
                return
            member = (
                WD_UNDERLINE.SINGLE
                if value is True
                else WD_UNDERLINE.NONE
                if value is False
                else WD_UNDERLINE(value)
            )
            if member is WD_UNDERLINE.INHERITED:
                r_pr.u = None
                return
            toggle = el.u(val=member.xml_value)
            r_pr.u = toggle
            toggle.parent = r_pr


# ---------------------------------------------------------------------------
# ParagraphFormat
# ---------------------------------------------------------------------------

_LINE_SPACING_MULTIPLES = {
    WD_LINE_SPACING.SINGLE: 240,
    WD_LINE_SPACING.ONE_POINT_FIVE: 360,
    WD_LINE_SPACING.DOUBLE: 480,
}


class ParagraphFormat:
    """``paragraph.paragraph_format`` and ``style.paragraph_format``.

    Over the ``w:pPr`` of a ``w:p`` or a ``w:style``. Every property is None
    when the ``w:pPr`` says nothing, which is python-docx's "inherited".
    """

    __slots__ = ("_holder", "_story")

    def __init__(self, holder: Any, story: Story | None = None) -> None:
        """Wrap a ``w:p`` or ``w:style``."""
        self._holder = holder
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one holder are equal."""
        return isinstance(other, ParagraphFormat) and other._holder is self._holder

    def __hash__(self) -> int:
        """Hashes by the holder's identity."""
        return hash(id(self._holder))

    @property
    def element(self) -> Any:
        """The typed element whose ``w:pPr`` this is over."""
        return self._holder

    @property
    def part(self) -> Any:
        """The part this format is in, or None for a format on a style."""
        return None if self._story is None else self._story.part

    def _p_pr(self, *, create: bool) -> PPr | None:
        p_pr = getattr(self._holder, "p_pr", None)
        if p_pr is None and create:
            p_pr = PPr()
            self._holder.p_pr = p_pr
            p_pr.parent = self._holder
        return p_pr

    def _child(self, name: str, *, create: bool) -> Any:
        p_pr = self._p_pr(create=create)
        if p_pr is None:
            return None
        child = getattr(p_pr, name, None)
        if child is None and create:
            child = getattr(el, name if name != "ind" else "ind")()
            setattr(p_pr, name, child)
            child.parent = p_pr
        return child

    # -- the on/off toggles ------------------------------------------------

    def _get_toggle(self, field: str) -> bool | None:
        p_pr = self._p_pr(create=False)
        toggle = getattr(p_pr, field, None) if p_pr is not None else None
        if toggle is None:
            return None
        value = getattr(toggle, "val", None)
        return True if value is None else bool(value)

    def _set_toggle(self, field: str, name: str, value: bool | None) -> None:
        if value is None and self._p_pr(create=False) is None:
            return
        with _write(self._story, f"docx.paragraph_format.{field}"):
            p_pr = self._p_pr(create=True)
            assert p_pr is not None
            if value is None:
                setattr(p_pr, field, None)
                return
            toggle = getattr(el, name)() if value else getattr(el, name)(val=False)
            setattr(p_pr, field, toggle)
            toggle.parent = p_pr

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """``w:jc``, as a ``WD_ALIGN_PARAGRAPH`` member or None."""
        p_pr = self._p_pr(create=False)
        value = getattr(getattr(p_pr, "jc", None), "val", None)
        value = getattr(value, "value", value)
        if value is None:
            return None
        try:
            return WD_PARAGRAPH_ALIGNMENT.from_xml(str(value))
        except ValueError:
            return None

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT | None) -> None:
        if value is None and self._p_pr(create=False) is None:
            return
        member = None if value is None else WD_PARAGRAPH_ALIGNMENT(value)
        office_js = None if member is None else ALIGNMENT_TO_OFFICE_JS.get(member)
        if office_js is not None and self._story is not None and isinstance(self._holder, P):
            # the content API has this one, and its setter records the report
            ContentParagraph(
                self._holder, self._story.body._container_of(self._holder), self._story.body
            ).alignment = office_js
            return
        with _write(self._story, "docx.paragraph_format.alignment"):
            p_pr = self._p_pr(create=True)
            assert p_pr is not None
            if member is None:
                p_pr.jc = None
                return
            jc = el.jc(val=member.xml_value)
            p_pr.jc = jc
            jc.parent = p_pr

    @property
    def keep_together(self) -> bool | None:
        """``w:keepLines``: do not break the paragraph across pages."""
        return self._get_toggle("keep_lines")

    @keep_together.setter
    def keep_together(self, value: bool | None) -> None:
        self._set_toggle("keep_lines", "keepLines", value)

    @property
    def keep_with_next(self) -> bool | None:
        """``w:keepNext``: keep on the same page as the next paragraph."""
        return self._get_toggle("keep_next")

    @keep_with_next.setter
    def keep_with_next(self, value: bool | None) -> None:
        self._set_toggle("keep_next", "keepNext", value)

    @property
    def page_break_before(self) -> bool | None:
        """``w:pageBreakBefore``: start at the top of a page."""
        return self._get_toggle("page_break_before")

    @page_break_before.setter
    def page_break_before(self, value: bool | None) -> None:
        self._set_toggle("page_break_before", "pageBreakBefore", value)

    @property
    def widow_control(self) -> bool | None:
        """``w:widowControl``: keep the first and last lines with the rest."""
        return self._get_toggle("widow_control")

    @widow_control.setter
    def widow_control(self, value: bool | None) -> None:
        self._set_toggle("widow_control", "widowControl", value)

    # -- the indents and the spacing --------------------------------------

    def _ind(self, attribute: str) -> Length | None:
        p_pr = self._p_pr(create=False)
        ind = getattr(p_pr, "ind", None) if p_pr is not None else None
        if ind is None:
            return None
        value = getattr(ind, attribute, None)
        return None if value is None else Twips(int(value))

    def _set_ind(self, attribute: str, value: int | Length | None) -> None:
        if value is None and getattr(self._p_pr(create=False), "ind", None) is None:
            return
        with _write(self._story, f"docx.paragraph_format.{attribute}"):
            p_pr = self._p_pr(create=True)
            assert p_pr is not None
            ind = p_pr.ind
            if ind is None:
                ind = el.ind()
                p_pr.ind = ind
                ind.parent = p_pr
            setattr(ind, attribute, twips_of(value))
            if all(
                getattr(ind, each, None) is None
                for each in ("start", "end", "left", "right", "first_line", "hanging")
            ):
                p_pr.ind = None

    @property
    def left_indent(self) -> Length | None:
        """``w:ind/@left``: the space from the left margin."""
        return self._ind("left")

    @left_indent.setter
    def left_indent(self, value: int | Length | None) -> None:
        self._set_ind("left", value)

    @property
    def tab_stops(self) -> TabStops:
        """``w:tabs``: the custom tab stops, a :class:`TabStops` sequence."""
        return TabStops(self)

    @property
    def right_indent(self) -> Length | None:
        """``w:ind/@right``: the space from the right margin."""
        return self._ind("right")

    @right_indent.setter
    def right_indent(self, value: int | Length | None) -> None:
        self._set_ind("right", value)

    @property
    def first_line_indent(self) -> Length | None:
        """``w:ind/@firstLine``, or minus ``@hanging``, as python-docx reports it."""
        p_pr = self._p_pr(create=False)
        ind = getattr(p_pr, "ind", None) if p_pr is not None else None
        if ind is None:
            return None
        if getattr(ind, "hanging", None) is not None:
            return Twips(-int(ind.hanging))
        if getattr(ind, "first_line", None) is not None:
            return Twips(int(ind.first_line))
        return None

    @first_line_indent.setter
    def first_line_indent(self, value: int | Length | None) -> None:
        with _write(self._story, "docx.paragraph_format.first_line_indent"):
            if value is None and getattr(self._p_pr(create=False), "ind", None) is None:
                return
            p_pr = self._p_pr(create=True)
            assert p_pr is not None
            ind = p_pr.ind
            if ind is None:
                ind = el.ind()
                p_pr.ind = ind
                ind.parent = p_pr
            ind.first_line = None
            ind.hanging = None
            if value is not None:
                twips = twips_of(value) or 0
                if twips < 0:
                    ind.hanging = -twips
                else:
                    ind.first_line = twips
            if all(
                getattr(ind, each, None) is None
                for each in ("start", "end", "left", "right", "first_line", "hanging")
            ):
                p_pr.ind = None

    def _spacing(self, attribute: str) -> Length | None:
        p_pr = self._p_pr(create=False)
        spacing = getattr(p_pr, "spacing", None) if p_pr is not None else None
        if spacing is None:
            return None
        value = getattr(spacing, attribute, None)
        return None if value is None else Twips(int(value))

    def _set_spacing(self, attribute: str, value: int | Length | None) -> None:
        if value is None and getattr(self._p_pr(create=False), "spacing", None) is None:
            return
        with _write(self._story, f"docx.paragraph_format.{attribute}"):
            p_pr = self._p_pr(create=True)
            assert p_pr is not None
            spacing = p_pr.spacing
            if spacing is None:
                spacing = el.spacing()
                p_pr.spacing = spacing
                spacing.parent = p_pr
            setattr(spacing, attribute, twips_of(value))

    @property
    def space_before(self) -> Length | None:
        """``w:spacing/@before``."""
        return self._spacing("before")

    @space_before.setter
    def space_before(self, value: int | Length | None) -> None:
        self._set_spacing("before", value)

    @property
    def space_after(self) -> Length | None:
        """``w:spacing/@after``."""
        return self._spacing("after")

    @space_after.setter
    def space_after(self, value: int | Length | None) -> None:
        self._set_spacing("after", value)

    def _line_rule(self) -> Any:
        p_pr = self._p_pr(create=False)
        spacing = getattr(p_pr, "spacing", None) if p_pr is not None else None
        value = getattr(spacing, "line_rule", None) if spacing is not None else None
        value = getattr(value, "value", value)
        if value is None:
            return None
        try:
            return WD_LINE_SPACING.from_xml(str(value))
        except ValueError:
            return None

    @property
    def line_spacing(self) -> float | Length | None:
        """A float number of lines, a :class:`Length`, or None.

        ``w:spacing/@line`` with ``@lineRule="auto"`` is a multiple, which
        python-docx reports as a float; anything else is an absolute height.
        """
        line = self._spacing("line")
        if line is None:
            return None
        if self._line_rule() is WD_LINE_SPACING.MULTIPLE:
            return int(line) / Pt(12)
        return line

    @line_spacing.setter
    def line_spacing(self, value: float | Length | None) -> None:
        with _write(self._story, "docx.paragraph_format.line_spacing"):
            if value is None:
                if getattr(self._p_pr(create=False), "spacing", None) is None:
                    return
                spacing = self._spacing_element(create=False)
                if spacing is not None:
                    spacing.line = None
                    spacing.line_rule = None
                return
            spacing = self._spacing_element(create=True)
            assert spacing is not None
            if isinstance(value, Length):
                spacing.line = twips_of(value)
                rule = getattr(spacing, "line_rule", None)
                rule = getattr(rule, "value", rule)
                if str(rule) != "atLeast":
                    spacing.line_rule = "exact"
            else:
                spacing.line = twips_of(Emu(int(value * Twips(240))))
                spacing.line_rule = "auto"

    def _spacing_element(self, *, create: bool) -> Any:
        p_pr = self._p_pr(create=create)
        if p_pr is None:
            return None
        spacing = p_pr.spacing
        if spacing is None and create:
            spacing = el.spacing()
            p_pr.spacing = spacing
            spacing.parent = p_pr
        return spacing

    @property
    def line_spacing_rule(self) -> WD_LINE_SPACING | None:
        """The rule, with SINGLE / ONE_POINT_FIVE / DOUBLE recognised."""
        rule = self._line_rule()
        if rule is WD_LINE_SPACING.MULTIPLE:
            line = self._spacing("line")
            for member, twips in _LINE_SPACING_MULTIPLES.items():
                if line is not None and int(line) == int(Twips(twips)):
                    return member
        return rule

    @line_spacing_rule.setter
    def line_spacing_rule(self, value: WD_LINE_SPACING | None) -> None:
        with _write(self._story, "docx.paragraph_format.line_spacing_rule"):
            spacing = self._spacing_element(create=True)
            assert spacing is not None
            if value in _LINE_SPACING_MULTIPLES:
                spacing.line = _LINE_SPACING_MULTIPLES[WD_LINE_SPACING(value)]
                spacing.line_rule = "auto"
                return
            if value is None:
                spacing.line_rule = None
                return
            spacing.line_rule = WD_LINE_SPACING(value).xml_value


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------


class TabStops:
    """``paragraph_format.tab_stops``: the custom tab stops of a paragraph or a style.

    python-docx ``TabStops``, over ``w:pPr/w:tabs``. A sequence of
    :class:`TabStop` in document order, which ``add_tab_stop`` keeps sorted by
    position as python-docx's does; deleting the last one removes ``w:tabs``.
    The paragraph's *inherited* tab stops (its style's, the defaults) are not
    here, as they are not in python-docx's.
    """

    __slots__ = ("_format",)

    def __init__(self, paragraph_format: ParagraphFormat) -> None:
        """Wrap a :class:`ParagraphFormat`."""
        self._format = paragraph_format

    @property
    def element(self) -> Any:
        """The typed holder (``w:p`` or ``w:style``) whose ``w:tabs`` this is over."""
        return self._format.element

    @property
    def part(self) -> Any:
        """The part the paragraph is in, or None for a style's tab stops."""
        return self._format.part

    def _tabs(self, *, create: bool) -> Any:
        return self._format._child("tabs", create=create)

    def _elements(self) -> list[Any]:
        tabs = self._tabs(create=False)
        return list(tabs.tab or ()) if tabs is not None else []

    def __len__(self) -> int:
        """How many custom tab stops the paragraph states."""
        return len(self._elements())

    def __iter__(self) -> Iterator[TabStop]:
        """The stops in document order."""
        return (TabStop(self, element) for element in self._elements())

    def __getitem__(self, idx: int) -> TabStop:
        """The stop at `idx`; ``IndexError`` past the end, as python-docx."""
        elements = self._elements()
        if not elements:
            raise IndexError("TabStops object is empty")
        return TabStop(self, elements[idx])

    def __delitem__(self, idx: int) -> None:
        """Remove the stop at `idx`, and ``w:tabs`` itself when it was the last."""
        with _write(self._format._story, "docx.tab_stops.delete"):
            tabs = self._tabs(create=False)
            elements = list(tabs.tab or ()) if tabs is not None else []
            try:
                element = elements[idx]
            except IndexError:
                raise IndexError("tab index out of range") from None
            tabs.tab.remove(element)
            if not tabs.tab:
                self._format._p_pr(create=False).tabs = None

    def add_tab_stop(
        self,
        position: int | Length,
        alignment: WD_TAB_ALIGNMENT | int = WD_TAB_ALIGNMENT.LEFT,
        leader: WD_TAB_LEADER | int | None = WD_TAB_LEADER.SPACES,
    ) -> TabStop:
        """Add a stop at `position` (a :class:`Length`; negative is allowed).

        Inserted in position order, after any stop at the same position, as
        python-docx does. ``SPACES`` (or None) writes no ``w:leader``.
        """
        with _write(self._format._story, "docx.tab_stops.add"):
            tabs = self._tabs(create=True)
            element = CTTabStop(
                val=WD_TAB_ALIGNMENT.to_xml(alignment),
                leader=_leader_xml(leader),
                pos=twips_of(position),
            )
            return TabStop(self, _insert_in_order(tabs, element))

    def clear_all(self) -> None:
        """Remove every custom tab stop: ``w:tabs`` goes."""
        p_pr = self._format._p_pr(create=False)
        if p_pr is None or getattr(p_pr, "tabs", None) is None:
            return
        with _write(self._format._story, "docx.tab_stops.clear_all"):
            p_pr.tabs = None


def _leader_xml(leader: Any) -> str | None:
    if leader is None or WD_TAB_LEADER(leader) is WD_TAB_LEADER.SPACES:
        return None
    return WD_TAB_LEADER.to_xml(leader)


def _insert_in_order(tabs: Any, element: Any) -> Any:
    """Put a ``w:tab`` in ``w:tabs`` before the first one further right."""
    if tabs.tab is None:
        tabs.tab = ChildList([], owner=tabs)
    position = int(element.pos or 0)
    index = len(tabs.tab)
    for i, existing in enumerate(tabs.tab):
        if int(existing.pos or 0) > position:
            index = i
            break
    tabs.tab.insert(index, element)
    element.parent = tabs
    return element


class TabStop:
    """One custom tab stop, a ``w:tab`` in ``w:tabs``. python-docx ``TabStop``."""

    __slots__ = ("_element", "_stops")

    def __init__(self, stops: TabStops, element: Any) -> None:
        """Wrap a ``w:tab``."""
        self._stops = stops
        self._element = element

    @property
    def element(self) -> Any:
        """The typed ``w:tab``."""
        return self._element

    @property
    def part(self) -> Any:
        """The part the paragraph is in, or None for a style's tab stop."""
        return self._stops.part

    @property
    def alignment(self) -> WD_TAB_ALIGNMENT:
        """``w:tab/@val``, a ``WD_TAB_ALIGNMENT`` member."""
        value = getattr(self._element, "val", None)
        return WD_TAB_ALIGNMENT.from_xml(str(getattr(value, "value", value)))

    @alignment.setter
    def alignment(self, value: WD_TAB_ALIGNMENT | int) -> None:
        with _write(self._stops._format._story, "docx.tab_stop.alignment"):
            self._element.val = WD_TAB_ALIGNMENT.to_xml(value)

    @property
    def leader(self) -> WD_TAB_LEADER:
        """``w:tab/@leader``; ``SPACES`` when absent, as python-docx reports it."""
        value = getattr(self._element, "leader", None)
        value = getattr(value, "value", value)
        return WD_TAB_LEADER.SPACES if value is None else WD_TAB_LEADER.from_xml(str(value))

    @leader.setter
    def leader(self, value: WD_TAB_LEADER | int | None) -> None:
        with _write(self._stops._format._story, "docx.tab_stop.leader"):
            self._element.leader = _leader_xml(value)

    @property
    def position(self) -> Length:
        """``w:tab/@pos`` as a :class:`Length`; may be negative."""
        return Twips(int(self._element.pos or 0))

    @position.setter
    def position(self, value: int | Length) -> None:
        """Move the stop, keeping ``w:tabs`` in position order as python-docx does."""
        with _write(self._stops._format._story, "docx.tab_stop.position"):
            tabs = self._element.parent
            tabs.tab.remove(self._element)
            self._element.pos = twips_of(value)
            _insert_in_order(tabs, self._element)


class _Text:
    """A ``w:t`` element, as ``run.add_text`` returns it."""

    __slots__ = ("_t",)

    def __init__(self, t_elm: Any) -> None:
        """Wrap the ``w:t``."""
        self._t = t_elm


class RenderedPageBreak:
    """A ``w:lastRenderedPageBreak``: where Word broke the page when it saved.

    Never written by this library, as python-docx never writes one either.
    """

    __slots__ = ("_element", "_parent")

    def __init__(self, element: Any, parent: Any) -> None:
        """Wrap the ``w:lastRenderedPageBreak``."""
        self._element = element
        self._parent = parent

    @property
    def element(self) -> Any:
        """The typed ``w:lastRenderedPageBreak``."""
        return self._element

    @property
    def part(self) -> Any:
        """The part this break is in."""
        return self._parent.part


class Hyperlink:
    """A ``w:hyperlink``: read-only, as python-docx 1.2 has it."""

    __slots__ = ("_hyperlink", "_story")

    def __init__(self, hyperlink: Any, story: Story) -> None:
        """Wrap the ``w:hyperlink``."""
        self._hyperlink = hyperlink
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:hyperlink`` are equal."""
        return isinstance(other, Hyperlink) and other._hyperlink is self._hyperlink

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._hyperlink))

    @property
    def part(self) -> Any:
        """The part this hyperlink is in."""
        return self._story.part

    @property
    def address(self) -> str:
        """The relationship's target, or ``""`` for an internal jump."""
        rel_id = getattr(self._hyperlink, "id", None)
        if not rel_id:
            return ""
        part = self._story.part
        rels = getattr(part, "relationships_part", None)
        if rels is None:
            return ""
        relationship = rels.get_relationship_by_id(rel_id)
        return getattr(relationship, "target", "") or "" if relationship is not None else ""

    @property
    def fragment(self) -> str:
        """``w:hyperlink/@anchor``, without the ``#``."""
        return getattr(self._hyperlink, "anchor", None) or ""

    @property
    def runs(self) -> list[Run]:
        """The runs the link's visible text is in."""
        return [
            Run(item, self._story)
            for item in getattr(self._hyperlink, "content", None) or ()
            if isinstance(item, R)
        ]

    @property
    def text(self) -> str:
        """The link's visible text."""
        return holder_text(self._hyperlink)

    @property
    def url(self) -> str:
        """The address and the fragment joined by ``#``, or ``""``."""
        address, fragment = self.address, self.fragment
        if not address:
            return ""
        return f"{address}#{fragment}" if fragment else address

    @property
    def contains_page_break(self) -> bool:
        """True when Word rendered a page break inside the link."""
        return bool(_rendered_breaks(self._hyperlink))


def _rendered_breaks(element: Any) -> list[Any]:
    """Every ``w:lastRenderedPageBreak`` under `element`."""
    from docx4j_py.traversal import element_name, iter_nodes

    return [
        node for node in iter_nodes(element) if element_name(node) == f"{_W}lastRenderedPageBreak"
    ]


class Run:
    """A ``w:r``: python-docx's ``Run``."""

    __slots__ = ("_r", "_story")

    def __init__(self, r: R, story: Story) -> None:
        """Wrap the ``w:r``."""
        self._r = r
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:r`` are equal."""
        return isinstance(other, Run) and other._r is self._r

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._r))

    def __repr__(self) -> str:
        """``<docx4j_py.docx Run 'Hello'>``."""
        text = self.text
        preview = text[:30] + "…" if len(text) > 30 else text
        return f"<docx4j_py.docx Run {preview!r}>"

    @property
    def element(self) -> R:
        """The typed ``w:r``. python-docx's is an lxml element."""
        return self._r

    @property
    def _element(self) -> R:
        """python-docx's private name for :attr:`element`."""
        return self._r

    @property
    def part(self) -> Any:
        """The part this run is in."""
        return self._story.part

    @property
    def view_paragraph(self) -> ContentParagraph:
        """The content API :class:`~docx4j_py.model.content.Paragraph` this run is in.

        An extension, and what ``Document.add_comment`` turns a pair of runs
        into a :class:`~docx4j_py.model.content.Range` with.
        """
        paragraph = _paragraph_of(self._r)
        if paragraph is None:
            raise ValueError("this run is not in a w:p")
        body = self._story.body
        return ContentParagraph(paragraph, body._container_of(paragraph), body)

    # -- text --------------------------------------------------------------

    @property
    def text(self) -> str:
        """The run's text: ``w:tab`` as ``\\t``, a line ``w:br`` as ``\\n``. python-docx's, not the text model's (CR-011 Q5)."""
        return run_text(self._r)

    @text.setter
    def text(self, text: str) -> None:
        # python-docx 1.2.0's: a tab w:tab, a newline and a carriage return each a
        # w:br (\r\n two), any other control character refused (CR-013 Phase A)
        check_python_docx_text(text)
        with recording(self._story.body, "docx.run.text"):
            content = self._r.content
            content.clear()
            for item in r_builder((text or "").replace("\r", "\n")).content:
                content.append(item)

    def add_text(self, text: str) -> _Text:
        """Append a ``w:t`` holding `text`, and return it."""
        check_python_docx_text(text)
        with recording(self._story.body, "docx.run.add_text"):
            item = el.t(text)
            self._r.content.append(item)
            return _Text(item)

    def add_tab(self) -> None:
        """Append a ``w:tab``."""
        with recording(self._story.body, "docx.run.add_tab"):
            self._r.content.append(el.tab())

    def add_break(self, break_type: WD_BREAK = WD_BREAK.LINE) -> None:
        """Append a ``w:br`` of `break_type`."""
        kind, clear = {
            WD_BREAK.LINE: (None, None),
            WD_BREAK.PAGE: ("page", None),
            WD_BREAK.COLUMN: ("column", None),
            WD_BREAK.LINE_CLEAR_LEFT: ("textWrapping", "left"),
            WD_BREAK.LINE_CLEAR_RIGHT: ("textWrapping", "right"),
            WD_BREAK.LINE_CLEAR_ALL: ("textWrapping", "all"),
        }[break_type]
        with recording(self._story.body, "docx.run.add_break"):
            item = el.br()
            if kind is not None:
                item.type_value = kind
            if clear is not None:
                item.clear = clear
            self._r.content.append(item)

    def clear(self) -> Run:
        """Remove the run's content, keeping its formatting."""
        with recording(self._story.body, "docx.run.clear"):
            self._r.content.clear()
        return self

    def add_picture(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
    ) -> InlineShape:
        """Append a picture to this run and return its :class:`InlineShape`."""
        from docx4j_py.docx.document import read_image
        from docx4j_py.docx.shape import InlineShape

        data = read_image(image_path_or_stream)
        paragraph = _paragraph_of(self._r)
        if paragraph is None:
            raise ValueError("this run is not in a paragraph, so a picture has nowhere to go")
        with recording(self._story.body, "docx.add_picture"):
            view = ContentParagraph(
                paragraph, self._story.body._container_of(paragraph), self._story.body
            )
            picture = view.insert_inline_picture(
                data, width=points_of(width), height=points_of(height)
            )
            # python-docx puts the drawing in *this* run; the content API made
            # one of its own, so the drawing moves across and the run goes
            new_run = picture.run
            if new_run is not self._r:
                items = list(new_run.content)
                container = paragraph.content
                for item in items:
                    self._r.content.append(item)
                if any(item is new_run for item in container):
                    container.remove(new_run)
            return InlineShape(picture.inline, self._story)

    # -- formatting --------------------------------------------------------

    @property
    def font(self) -> Font:
        """The run's :class:`Font`."""
        return Font(self._r, self._story)

    @property
    def bold(self) -> bool | None:
        """``font.bold``."""
        return self.font.bold

    @bold.setter
    def bold(self, value: bool | None) -> None:
        self.font.bold = value

    @property
    def italic(self) -> bool | None:
        """``font.italic``."""
        return self.font.italic

    @italic.setter
    def italic(self, value: bool | None) -> None:
        self.font.italic = value

    @property
    def underline(self) -> bool | WD_UNDERLINE | None:
        """``font.underline``."""
        return self.font.underline

    @underline.setter
    def underline(self, value: bool | WD_UNDERLINE | None) -> None:
        self.font.underline = value

    @property
    def style(self) -> CharacterStyle | None:
        """The run's character style, or the document's default."""
        from docx4j_py.docx.styles import style_for

        r_pr = _r_pr(self._r, create=False)
        style_id = getattr(getattr(r_pr, "r_style", None), "val", None)
        return style_for(self._story.package, style_id, WD_STYLE_TYPE.CHARACTER)

    @style.setter
    def style(self, style_or_name: str | CharacterStyle | None) -> None:
        from docx4j_py.docx.styles import style_id_for

        style_id = style_id_for(self._story.package, style_or_name, WD_STYLE_TYPE.CHARACTER)
        with recording(self._story.body, "docx.run.style"):
            r_pr = _r_pr(self._r, create=style_id is not None)
            if r_pr is None:
                return
            if style_id is None:
                r_pr.r_style = None
                return
            r_pr.r_style = el.rStyle(val=style_id)
            r_pr.r_style.parent = r_pr

    # -- reading -----------------------------------------------------------

    @property
    def contains_page_break(self) -> bool:
        """True when Word rendered a page break in this run."""
        return bool(_rendered_breaks(self._r))

    def iter_inner_content(self) -> Iterator[Any]:
        """The run's content: text as ``str``, breaks and drawings as objects.

        Contiguous text-bearing items come out as one ``str``, a rendered page
        break as a :class:`RenderedPageBreak` and a ``w:drawing`` as a
        :class:`~docx4j_py.docx.shape.Drawing`, which is python-docx 1.2's rule.
        """
        from docx4j_py.docx.shape import Drawing
        from docx4j_py.traversal import element_name

        buffer: list[Any] = []

        def flush() -> Iterator[str]:
            if buffer:
                text = run_text(R(content=list(buffer)))
                buffer.clear()
                if text:
                    yield text

        for item in getattr(self._r, "content", None) or ():
            name = element_name(item)
            if name == f"{_W}lastRenderedPageBreak":
                yield from flush()
                yield RenderedPageBreak(item, self)
            elif name == f"{_W}drawing":
                yield from flush()
                yield Drawing(item, self)
            else:
                buffer.append(item)
        yield from flush()


def _paragraph_of(element: Any) -> P | None:
    """The ``w:p`` an element is in, from the parent pointers."""
    current = getattr(element, "parent", None)
    for _ in range(64):
        if current is None:
            return None
        if isinstance(current, P):
            return current
        current = getattr(current, "parent", None)
    return None


# ---------------------------------------------------------------------------
# Paragraph
# ---------------------------------------------------------------------------


class Paragraph:
    """A ``w:p``: python-docx's ``Paragraph``."""

    __slots__ = ("_p", "_story")

    def __init__(self, p: P, story: Story) -> None:
        """Wrap the ``w:p``."""
        self._p = p
        self._story = story

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:p`` are equal."""
        return isinstance(other, Paragraph) and other._p is self._p

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._p))

    def __repr__(self) -> str:
        """``<docx4j_py.docx Paragraph 'Chapter 2'>``."""
        text = self.text
        preview = text[:30] + "…" if len(text) > 30 else text
        return f"<docx4j_py.docx Paragraph {preview!r}>"

    # -- what is underneath ------------------------------------------------

    @property
    def element(self) -> P:
        """The typed ``w:p``."""
        return self._p

    @property
    def _element(self) -> P:
        """python-docx's private name for :attr:`element`."""
        return self._p

    @property
    def part(self) -> Any:
        """The part this paragraph is in."""
        return self._story.part

    @property
    def view(self) -> ContentParagraph:
        """The content API's :class:`~docx4j_py.model.content.Paragraph`.

        An extension, and the door out of the facade: tracked changes,
        comments, lists, addresses and ``to_api_script`` are all on it.
        """
        body = self._story.body
        return ContentParagraph(self._p, body._container_of(self._p), body)

    # -- text --------------------------------------------------------------

    @property
    def text(self) -> str:
        """The paragraph's text: its runs and its hyperlinks, as python-docx reads it; not the text model's (CR-011 Q5)."""
        return holder_text(self._p)

    @text.setter
    def text(self, text: str | None) -> None:
        with recording(self._story.body, "docx.paragraph.text"):
            self.clear()
            self.add_run(text)

    def clear(self) -> Paragraph:
        """Remove the paragraph's content, keeping its ``w:pPr``."""
        with recording(self._story.body, "docx.paragraph.clear"):
            p_pr = self._p.p_pr
            self._p.content.clear()
            self._p.p_pr = p_pr
        return self

    @property
    def runs(self) -> list[Run]:
        """The ``w:r`` children of this paragraph."""
        return [
            Run(item, self._story)
            for item in getattr(self._p, "content", None) or ()
            if isinstance(item, R)
        ]

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append a run holding `text`, with character style `style`."""
        with recording(self._story.body, "docx.add_run"):
            run = R()
            self._p.content.append(run)
            view = Run(run, self._story)
            if text:
                view.text = text
            if style:
                view.style = style
            return view

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """A new paragraph directly before this one."""
        with recording(self._story.body, "docx.insert_paragraph_before"):
            new = self.view.insert_paragraph("", location="Before")
            paragraph = Paragraph(new.element, self._story)
            if text:
                paragraph.add_run(text)
            if style is not None:
                paragraph.style = style
            return paragraph

    def add_picture(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
    ) -> InlineShape:
        """A picture in a new run at the end of this paragraph. An extension."""
        return self.add_run().add_picture(image_path_or_stream, width, height)

    # -- formatting --------------------------------------------------------

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """``w:jc``, as a ``WD_ALIGN_PARAGRAPH`` member or None."""
        return self.paragraph_format.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT | None) -> None:
        self.paragraph_format.alignment = value

    @property
    def paragraph_format(self) -> ParagraphFormat:
        """The paragraph's :class:`ParagraphFormat`."""
        return ParagraphFormat(self._p, self._story)

    @property
    def style(self) -> ParagraphStyle | None:
        """The paragraph's style object, or the document's default."""
        from docx4j_py.docx.styles import style_for

        p_pr = getattr(self._p, "p_pr", None)
        style_id = getattr(getattr(p_pr, "p_style", None), "val", None)
        return style_for(self._story.package, style_id, WD_STYLE_TYPE.PARAGRAPH)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None) -> None:
        from docx4j_py.docx.styles import style_id_for

        package = self._story.package
        style_id = style_id_for(package, style_or_name, WD_STYLE_TYPE.PARAGRAPH, define=True)
        with recording(self._story.body, "docx.paragraph.style"):
            if style_id is None:
                p_pr = getattr(self._p, "p_pr", None)
                if p_pr is not None:
                    p_pr.p_style = None
                return
            self.view.style_id = style_id

    # -- reading -----------------------------------------------------------

    @property
    def contains_page_break(self) -> bool:
        """True when Word rendered a page break in this paragraph."""
        return bool(_rendered_breaks(self._p))

    @property
    def rendered_page_breaks(self) -> list[RenderedPageBreak]:
        """Every ``w:lastRenderedPageBreak`` in this paragraph; usually none."""
        return [RenderedPageBreak(item, self) for item in _rendered_breaks(self._p)]

    @property
    def hyperlinks(self) -> list[Hyperlink]:
        """The ``w:hyperlink`` children of this paragraph."""
        from docx4j_py.traversal import element_name

        return [
            Hyperlink(item, self._story)
            for item in getattr(self._p, "content", None) or ()
            if element_name(item) == f"{_W}hyperlink"
        ]

    def iter_inner_content(self) -> Iterator[Run | Hyperlink]:
        """The runs and hyperlinks of this paragraph, in document order."""
        from docx4j_py.traversal import element_name

        for item in getattr(self._p, "content", None) or ():
            if isinstance(item, R):
                yield Run(item, self._story)
            elif element_name(item) == f"{_W}hyperlink":
                yield Hyperlink(item, self._story)
