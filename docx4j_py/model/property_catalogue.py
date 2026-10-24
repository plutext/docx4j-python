"""The property catalogue. docx4j ``org.docx4j.model.styles.PropertyCatalogue``.

CR-002 section 6.1. One table per properties element --- :data:`RUN` for
``w:rPr`` and the paragraph mark's ``w:pPr/w:rPr``, :data:`PARAGRAPH` for
``PPrBase``, :data:`TABLE` for ``CT_TblPrBase`` and :data:`CELL` for ``w:tcPr``
--- one entry per schema member in schema order, each saying how to read it,
write it, merge a more specific value over an inherited one, whether it is empty
and whether it formats the content.

``apply``, ``is_empty``, ``unset`` and ``has_direct_formatting`` in
:mod:`docx4j_py.model.style_util` are all derived from these tables by
iteration, so the four cannot drift apart as docx4j's hand-kept lists did before
its CR-015 phase 1.

Every carried leaf is a **copy** (:func:`copy_leaf`, which is
:func:`docx4j_py.child.deep_copy`), never the styles part's own object: an
effective ``w:rPr``'s ``w:b`` must not be the styles part's ``w:b``, or a caller
editing the effective object would edit the document.

Entries are keyed by the **Python attribute name** of the generated dataclass,
not by the schema's member name: ``w14:shadow`` and ``w:shadow`` are two members
of ``w:rPr`` and only the attribute names tell them apart. :attr:`Property.xml_name`
carries the schema's name, for messages and for tests.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Callable, Iterable, Sequence
from typing import Any

from docx4j_py.child import deep_copy
from docx4j_py.model.values import (
    emphasis_mark,
    highlight,
    is_on,
    states_colour,
    theme_colour,
    underline,
    vertical_alignment,
)

__all__ = [
    "CELL",
    "CELL_EXCLUDED",
    "IND",
    "PARAGRAPH",
    "PARAGRAPH_EXCLUDED",
    "RUN",
    "RUN_EXCLUDED",
    "TABLE",
    "TABLE_EXCLUDED",
    "TOGGLES",
    "TOGGLE_NAMES",
    "Property",
    "apply_catalogue",
    "apply_property",
    "copy_leaf",
    "has_direct_formatting_of",
    "is_empty_of",
    "is_true",
    "named_property",
    "unset_catalogue",
]


def copy_leaf(value: Any) -> Any:
    """A copy of a leaf value sharing nothing with it, and with no parent.

    :func:`docx4j_py.child.deep_copy` copies fields only and relinks the copy's
    descendants, which is exactly what a merged property wants: the effective
    object must not alias a style definition.
    """
    return deep_copy(value)


# ---------------------------------------------------------------- scalars


def _empty_scalar(value: Any) -> bool:
    """docx4j ``StyleUtil.isEmpty(String)``: absent, or the empty string.

    The model holds an ``ST_HexColor`` as ``bytes`` (``w:val=""`` parses to
    ``b""``), where docx4j holds a ``String``, so ``b""`` is the empty string too.
    """
    return value is None or value == "" or value == b""


def _attr(source: Any, destination: Any) -> Any:
    """docx4j ``StyleUtil.apply(String, String)``: the source where it states one."""
    return destination if source is None else source


def _empty_theme_colour(value: Any) -> bool:
    """docx4j ``isEmpty(STThemeColor)``: absent, ``none``, or outside the enumeration (:func:`theme_colour`)."""
    member = theme_colour(value)
    return member is None or member.value == "none"


def _stated_fill_or_border(value: Any) -> bool:
    """Whether a fill's or a border's colour value is stated: any value at all (CR-018 Phase B).

    Word reads one it rejects (``F00``, ``bogus``, ``""``) as **automatic**, and
    that overrides a style's colour (round six: Word saves ``w:fill="auto"``,
    ``w:color="auto"``), so every value present is stated. docx4j's
    ``statesColour`` holds ``""`` empty; here it is the automatic Word reads, as
    the normal form writes it (``auto``; CR-018 section 11.7).
    """
    return value is not None


def _states_colour(value: Any, theme_colour: Any, stated: Callable[[Any], bool] = _stated_fill_or_border) -> bool:
    """docx4j ``StyleUtil.statesColour``: whether a colour's value or theme reference is stated.

    Where it is, the colour inherits as one (:func:`_merge_colour`). ``auto`` ---
    the ``STHexColorAuto`` member --- is a stated value, as it is a non-empty
    ``String`` to docx4j; a theme reference of ``none`` is not a stated reference.
    `stated` reads the value: for a text or underline colour
    :func:`~docx4j_py.model.values.states_colour`, under which a value Word
    rejects is as if absent (docx4j's ``statesWordColour``, CR-060 phase 4); for
    a fill's or a border's, any value (:func:`_stated_fill_or_border`).
    """
    return stated(value) or not _empty_theme_colour(theme_colour)


# ---------------------------------------------------------------- emptiness


def is_empty_bool(value: Any) -> bool:
    """``<w:b/>`` and ``<w:b w:val="0"/>`` both count: an explicit false is applied."""
    return value is None


def is_true(value: Any) -> bool:
    """A ``BooleanDefaultTrue``'s value as Word reads it: an absent ``w:val`` means true (docx4j ``isVal``),
    and only exactly ``false``, ``off`` or ``0`` means false (CR-018, :func:`docx4j_py.model.values.is_on`)."""
    return is_on(value)


def is_empty_val(value: Any) -> bool:
    """A leaf whose only content is ``w:val``: absent, or with no value."""
    return value is None or _empty_scalar(value.val)


def _is_empty_enumerated(read: Callable[[Any], Any]) -> Callable[[Any], bool]:
    """A ``w:val`` leaf whose value is an enumeration: empty unless `read` (Word's reading) gives a member.

    CR-018 Phase B, rule 2: a ``w:highlight``, ``w:vertAlign`` or ``w:em`` whose
    value is missing or outside its enumeration (``YELLOW``, ``Superscript``,
    ``DOT``, ``bogus``, ``""``) is unset, as if absent, so a style's value
    beneath shows (round five). docx4j has it from JAXB, which reads such a value
    as null, and for ``w:highlight`` from ``WordRepairs.isHighlightColor``.
    """

    def is_empty(value: Any) -> bool:
        return value is None or read(value.val) is None

    return is_empty


def is_null(value: Any) -> bool:
    """Never empty in itself: absent only."""
    return value is None


def is_empty_r_fonts(value: Any) -> bool:
    """docx4j ``isEmpty(RFonts)``: insensitive to a ``w:hint``, which ``apply`` still carries."""
    return value is None or (
        _empty_scalar(value.ascii)
        and value.ascii_theme is None
        and _empty_scalar(value.cs)
        and value.cstheme is None
        and _empty_scalar(value.east_asia)
        and value.east_asia_theme is None
        and _empty_scalar(value.h_ansi)
        and value.h_ansi_theme is None
    )


def is_empty_color(value: Any) -> bool:
    """docx4j ``isEmpty(Color)``: no value and no theme attribute (since docx4j 17.3.3).

    A ``w:color`` stating only a theme colour is not empty, so a style's
    ``w:themeColor="accent1"`` with no ``w:val`` is inherited, where before
    docx4j's CR-034 only ``w:val`` counted and it was dropped. As in the Java,
    the reference counts when it is present at all, ``none`` included. A value
    Word rejects (``F00``, ``bogus``, ``Auto``, ``""``) is no value (CR-018 Phase
    B; docx4j 17.3.5, ``WordColors.stated``), so a style's colour beneath shows.
    """
    return value is None or (
        not states_colour(value.val)
        and value.theme_color is None
        and _empty_scalar(value.theme_tint)
        and _empty_scalar(value.theme_shade)
    )


def is_empty_u(value: Any) -> bool:
    """docx4j ``isEmpty(U)``: neither a kind nor a colour, the theme attributes counted (17.3.3).

    The kind and the colour as Word reads them (CR-018 Phase B): a ``w:val``
    outside the enumeration is no kind (the six on/off words are ``none``,
    :func:`~docx4j_py.model.values.underline`), and a colour Word rejects is no
    colour (:func:`~docx4j_py.model.values.states_colour`).
    """
    return value is None or (
        underline(value.val) is None
        and not states_colour(value.color)
        and value.theme_color is None
        and _empty_scalar(value.theme_tint)
        and _empty_scalar(value.theme_shade)
    )


def is_empty_lang(value: Any) -> bool:
    """docx4j ``isEmpty(CTLanguage)``, since its CR-015 phase 1 fix.

    A style stating only ``w:lang`` used to count as empty and was never applied.
    """
    return value is None or (
        _empty_scalar(value.val)
        and _empty_scalar(value.east_asia)
        and _empty_scalar(value.bidi)
    )


def is_empty_shd(value: Any) -> bool:
    """docx4j ``isEmpty(CTShd)``; a colour or a fill counts whatever its value (:func:`_stated_fill_or_border`)."""
    return value is None or (
        value.color is None
        and _empty_scalar(value.theme_tint)
        and _empty_scalar(value.theme_shade)
        and value.fill is None
        and _empty_scalar(value.theme_fill_tint)
        and _empty_scalar(value.theme_fill_shade)
        and value.theme_color is None
        and value.theme_fill is None
        and value.val is None
    )


def is_empty_border(value: Any) -> bool:
    """docx4j ``isEmpty(CTBorder)``; a colour counts whatever its value (:func:`_stated_fill_or_border`)."""
    return value is None or (
        value.color is None
        and value.space is None
        and value.sz is None
        and value.theme_color is None
        and _empty_scalar(value.theme_shade)
        and _empty_scalar(value.theme_tint)
        and value.val is None
    )


def is_empty_p_bdr(value: Any) -> bool:
    """docx4j ``isEmpty(PPrBase.PBdr)``: every edge empty."""
    return value is None or (
        is_empty_border(value.top)
        and is_empty_border(value.left)
        and is_empty_border(value.bottom)
        and is_empty_border(value.right)
        and is_empty_border(value.between)
        and is_empty_border(value.bar)
    )


def is_empty_ind(value: Any) -> bool:
    """docx4j ``isEmpty(Ind)``.

    ``w:start`` and ``w:end`` --- the strict-conformance spellings --- count too
    (docx4j 17.0.5); dropping them lost them in style resolution.
    """
    return value is None or (
        value.first_line is None
        and value.first_line_chars is None
        and value.hanging is None
        and value.hanging_chars is None
        and value.left is None
        and value.left_chars is None
        and value.right is None
        and value.right_chars is None
        and value.start is None
        and value.start_chars is None
        and value.end is None
        and value.end_chars is None
    )


def is_empty_spacing(value: Any) -> bool:
    """docx4j ``isEmpty(Spacing)``; an explicit autospacing of either value counts."""
    return value is None or (
        value.after is None
        and value.after_lines is None
        and value.before is None
        and value.before_lines is None
        and value.line is None
        and value.line_rule is None
        and value.before_autospacing is None
        and value.after_autospacing is None
    )


def _is_empty_tab_stop(value: Any) -> bool:
    return value is None or (value.pos is None and value.val is None and value.leader is None)


def is_empty_tabs(value: Any) -> bool:
    """docx4j ``isEmpty(Tabs)``: no stop that states anything."""
    if value is None or not value.tab:
        return True
    return not any(not _is_empty_tab_stop(stop) for stop in value.tab)


def is_empty_num_pr(value: Any) -> bool:
    """A ``w:numPr`` is empty only when it states neither ``w:numId`` nor ``w:ilvl``.

    Each inherits on its own: CR-015's probe ``styles-numpr-ilvl-only``.
    """
    return value is None or (
        (value.num_id is None or value.num_id.val is None)
        and (value.ilvl is None or value.ilvl.val is None)
    )


def is_empty_frame_pr(value: Any) -> bool:
    """docx4j ``isEmpty(CTFramePr)``: nothing stated, ``w:anchorLock`` included.

    docx4j reads ``w:anchorLock`` as stated (``FramePrAccess.getAnchorLock``,
    ``82c64bf14``), since CR-059 corrected its schema default to false: a bare
    ``<w:framePr/>`` is empty, and ``w:anchorLock="1"`` or ``"0"`` alone is not.
    Until then docx4j read an absent one as true, and so did this port.
    """
    if value is None:
        return True
    return (
        value.drop_cap is None
        and value.lines is None
        and value.w is None
        and value.h is None
        and value.v_space is None
        and value.h_space is None
        and value.wrap is None
        and value.h_anchor is None
        and value.v_anchor is None
        and value.x is None
        and value.x_align is None
        and value.y is None
        and value.y_align is None
        and value.h_rule is None
        and value.anchor_lock is None
    )


def _is_empty_tbl_width(value: Any) -> bool:
    # @w:type is ignored, as docx4j's isEmpty does
    return value is None or value.w is None


def _is_empty_tbl_cell_mar(value: Any) -> bool:
    # w:start / w:end are not read: docx4j's isEmpty / apply pair does not read them
    return value is None or (
        _is_empty_tbl_width(value.bottom)
        and _is_empty_tbl_width(value.left)
        and _is_empty_tbl_width(value.right)
        and _is_empty_tbl_width(value.top)
    )


def _is_empty_tbl_borders(value: Any) -> bool:
    return value is None or (
        is_empty_border(value.bottom)
        and is_empty_border(value.left)
        and is_empty_border(value.right)
        and is_empty_border(value.top)
        and is_empty_border(value.inside_h)
        and is_empty_border(value.inside_v)
    )


def _is_empty_tc_borders(value: Any) -> bool:
    return value is None or (
        is_empty_border(value.bottom)
        and is_empty_border(value.left)
        and is_empty_border(value.right)
        and is_empty_border(value.top)
        and is_empty_border(value.inside_h)
        and is_empty_border(value.inside_v)
        and is_empty_border(value.tl2br)
        and is_empty_border(value.tr2bl)
    )


def _is_empty_tc_mar(value: Any) -> bool:
    return value is None or (
        _is_empty_tbl_width(value.bottom)
        and _is_empty_tbl_width(value.left)
        and _is_empty_tbl_width(value.right)
        and _is_empty_tbl_width(value.top)
    )


def _is_empty_tbl_look(value: Any) -> bool:
    return value is None or (
        value.first_column is None
        and value.first_row is None
        and value.last_column is None
        and value.last_row is None
        and value.no_hband is None
        and value.no_vband is None
        and value.val is None
    )


def _is_empty_tblp_pr(value: Any) -> bool:
    return value is None or (
        value.left_from_text is None
        and value.right_from_text is None
        and value.top_from_text is None
        and value.bottom_from_text is None
        and value.vert_anchor is None
        and value.horz_anchor is None
        and value.tblp_xspec is None
        and value.tblp_x is None
        and value.tblp_yspec is None
        and value.tblp_y is None
    )


def _is_empty_tbl_layout(value: Any) -> bool:
    return value is None or value.type_value is None


# ---------------------------------------------------------------- merges


def merge_replace(source: Any, destination: Any) -> Any:
    """No per-attribute rule: a copy of the source where it has one, else the inherited."""
    return destination if source is None else copy_leaf(source)


def merge_bool(source: Any, destination: Any) -> Any:
    """docx4j ``apply(BooleanDefaultTrue, BooleanDefaultTrue)``."""
    return destination if source is None else copy_leaf(source)


def _not_carried(_source: Any, destination: Any) -> Any:
    """Never carried: docx4j's ``apply`` for the owning element does not mention it."""
    return destination


def _merge_val(
    cls: type, is_empty: Callable[[Any], bool], read: Callable[[Any], Any] | None = None
) -> Callable[[Any, Any], Any]:
    """A leaf whose only content is ``w:val``, taken whole when the source states it.

    `read`, where given, is what the value is carried as: an enumeration's
    member, as Word reads it (CR-018 Phase B).
    """

    def merge(source: Any, destination: Any) -> Any:
        if source is None or is_empty(source):
            return destination
        out = cls() if destination is None else destination
        out.val = source.val if read is None else read(source.val)
        return out

    return merge


def merge_r_fonts(source: Any, destination: Any) -> Any:
    """``w:rFonts``: a theme reference trumps the explicit face beside it.

    A source stating only a ``w:hint`` still says something, and a *null* source
    leaves the destination alone (docx4j's CR-015 phase 1 fix: before it, a null
    source was given an empty ``w:rFonts`` and every merged ``w:rPr`` then read
    as directly formatted).
    """
    from docx4j_py.wml import RFonts

    if source is None:
        return destination
    out = RFonts() if destination is None else destination
    if is_empty_r_fonts(source):
        if source.hint is not None:
            out.hint = source.hint
        # else: keep the (possibly empty) element, which RunFontSelector relies on
        return out
    if source.ascii is not None:
        out.ascii = source.ascii
        out.ascii_theme = None
    if source.cs is not None:
        out.cs = source.cs
        out.cstheme = None
    if source.east_asia is not None:
        out.east_asia = source.east_asia
        out.east_asia_theme = None
    if source.h_ansi is not None:
        out.h_ansi = source.h_ansi
        out.h_ansi_theme = None
    if source.ascii_theme is not None:
        out.ascii_theme = source.ascii_theme
        out.ascii = None
    if source.cstheme is not None:
        out.cstheme = source.cstheme
        out.cs = None
    if source.east_asia_theme is not None:
        out.east_asia_theme = source.east_asia_theme
        out.east_asia = None
    if source.h_ansi_theme is not None:
        out.h_ansi_theme = source.h_ansi_theme
        out.h_ansi = None
    if source.hint is not None:
        out.hint = source.hint
    return out


#: The four attributes of one colour, as :func:`_merge_colour` takes them: the
#: value, the theme reference, its tint and its shade.
_COLOR = ("val", "theme_color", "theme_tint", "theme_shade")
#: ``w:u``'s colour, ``w:shd``'s pattern colour and a border's colour.
_COLOR_ATTRIBUTE = ("color", "theme_color", "theme_tint", "theme_shade")
#: ``w:shd``'s fill.
_FILL = ("fill", "theme_fill", "theme_fill_tint", "theme_fill_shade")


def _merge_colour(
    source: Any,
    destination: Any,
    names: tuple[str, str, str, str],
    stated: Callable[[Any], bool] = _stated_fill_or_border,
) -> None:
    """A colour inherits **as one** (docx4j 17.3.3, its CR-034).

    Where the source states a value or a theme reference (:func:`_states_colour`),
    its value, reference, tint and shade replace the destination's together,
    ``None`` included; otherwise a tint or a shade stated alone still inherits on
    its own. Word draws a theme reference over the value written beside it, so
    taken attribute by attribute --- CR-015's merge, before 17.3.3 --- a run's
    ``w:val="auto"`` under a style's ``w:themeColor="accent1"`` kept the style's
    reference, and a character style's plain ``accent2`` kept a paragraph style's
    ``w:themeShade``, where Word draws the first black and the second in the
    plain accent.

    ``names`` is the colour's four attribute names, value first: :data:`_COLOR`,
    :data:`_COLOR_ATTRIBUTE` or :data:`_FILL`. ``stated`` reads the value
    (:func:`_states_colour`): a text or underline colour Word rejects is passed
    over, a fill's or a border's is a stated automatic (CR-018 Phase B).
    """
    value, theme_colour, tint, shade = names
    if _states_colour(getattr(source, value), getattr(source, theme_colour), stated):
        for name in names:
            setattr(destination, name, getattr(source, name))
    else:
        setattr(destination, tint, _attr(getattr(source, tint), getattr(destination, tint)))
        setattr(destination, shade, _attr(getattr(source, shade), getattr(destination, shade)))


def merge_color(source: Any, destination: Any) -> Any:
    """docx4j ``apply(Color, Color)``: ``w:color`` inherits as one (:func:`_merge_colour`)."""
    from docx4j_py.wml import Color

    if is_empty_color(source):
        return destination
    out = Color() if destination is None else destination
    _merge_colour(source, out, _COLOR, states_colour)
    return out


def merge_u(source: Any, destination: Any) -> Any:
    """docx4j ``apply(U, U)``: the kind inherits on its own, the colour as one.

    Both as Word reads them (CR-018 Phase B): the kind is
    :func:`~docx4j_py.model.values.underline`'s member, so a ``w:val`` outside
    the enumeration leaves the inherited kind (round five; S11 is the missing
    one) and the six on/off words are ``none``, which removes it (round six);
    a colour Word rejects leaves the inherited colour.
    """
    from docx4j_py.wml import U

    if is_empty_u(source):
        return destination
    out = U() if destination is None else destination
    out.val = _attr(underline(source.val), out.val)
    _merge_colour(source, out, _COLOR_ATTRIBUTE, states_colour)
    return out


def merge_lang(source: Any, destination: Any) -> Any:
    """``w:lang``: ``w:val``, ``w:eastAsia`` and ``w:bidi`` each name one script range."""
    from docx4j_py.wml import CTLanguage

    if is_empty_lang(source):
        return destination
    out = CTLanguage() if destination is None else destination
    out.val = _attr(source.val, out.val)
    out.east_asia = _attr(source.east_asia, out.east_asia)
    out.bidi = _attr(source.bidi, out.bidi)
    return out


def merge_shd(source: Any, destination: Any) -> Any:
    """docx4j ``apply(CTShd, CTShd)``: the pattern inherits on its own, each colour as one.

    The pattern colour and the fill are two colours (:func:`_merge_colour`), so a
    direct ``w:shd`` stating only a pattern keeps an inherited ``w:themeFill``
    (CR-015), and one stating a fill replaces the inherited fill's reference.
    """
    from docx4j_py.wml import CTShd

    if is_empty_shd(source):
        return destination
    out = CTShd() if destination is None else destination
    out.val = _attr(source.val, out.val)
    _merge_colour(source, out, _COLOR_ATTRIBUTE)
    _merge_colour(source, out, _FILL)
    return out


def merge_border(source: Any, destination: Any) -> Any:
    """docx4j ``apply(CTBorder, CTBorder)``, for ``w:bdr`` and every edge of a border set.

    ``w:val``, ``w:sz`` and ``w:space`` each inherit on their own, the colour as
    one (:func:`_merge_colour`).
    """
    from docx4j_py.wml import CTBorder

    if is_empty_border(source):
        return destination
    out = CTBorder() if destination is None else destination
    out.space = _attr(source.space, out.space)
    out.sz = _attr(source.sz, out.sz)
    out.val = _attr(source.val, out.val)
    _merge_colour(source, out, _COLOR_ATTRIBUTE)
    return out


def merge_p_bdr(source: Any, destination: Any) -> Any:
    """``w:pBdr``: edge by edge."""
    from docx4j_py.wml import CtPprBasePBdr

    if is_empty_p_bdr(source):
        return destination
    out = CtPprBasePBdr() if destination is None else destination
    out.top = merge_border(source.top, out.top)
    out.left = merge_border(source.left, out.left)
    out.bottom = merge_border(source.bottom, out.bottom)
    out.right = merge_border(source.right, out.right)
    out.between = merge_border(source.between, out.between)
    out.bar = merge_border(source.bar, out.bar)
    return out


def merge_num_pr(source: Any, destination: Any) -> Any:
    """``w:numPr``: ``w:numId`` and ``w:ilvl`` each inherit on their own."""
    from docx4j_py.wml import CtPprBaseNumPr

    if is_empty_num_pr(source):
        return destination
    out = CtPprBaseNumPr() if destination is None else destination
    if source.num_id is not None and source.num_id.val is not None:
        out.num_id = copy_leaf(source.num_id)
    if source.ilvl is not None and source.ilvl.val is not None:
        out.ilvl = copy_leaf(source.ilvl)
    return out


def merge_ind(source: Any, destination: Any) -> Any:
    """``w:ind``.

    ``w:firstLine`` and ``w:hanging`` are two spellings of one property
    (ECMA-376 17.3.1.12), so whichever of the two the more specific ``w:ind``
    states replaces both; the rest inherit per attribute.
    """
    from docx4j_py.wml import CtPprBaseInd

    if is_empty_ind(source):
        return destination
    out = CtPprBaseInd() if destination is None else destination
    source_first_line = source.first_line is not None or source.first_line_chars is not None
    source_hanging = source.hanging is not None or source.hanging_chars is not None
    if source_first_line and not source_hanging:
        out.hanging = None
        out.hanging_chars = None
    elif source_hanging and not source_first_line:
        out.first_line = None
        out.first_line_chars = None
    out.first_line = _attr(source.first_line, out.first_line)
    out.first_line_chars = _attr(source.first_line_chars, out.first_line_chars)
    out.hanging = _attr(source.hanging, out.hanging)
    out.hanging_chars = _attr(source.hanging_chars, out.hanging_chars)
    out.left = _attr(source.left, out.left)
    out.left_chars = _attr(source.left_chars, out.left_chars)
    out.right = _attr(source.right, out.right)
    out.right_chars = _attr(source.right_chars, out.right_chars)
    out.start = _attr(source.start, out.start)
    out.start_chars = _attr(source.start_chars, out.start_chars)
    out.end = _attr(source.end, out.end)
    out.end_chars = _attr(source.end_chars, out.end_chars)
    return out


def merge_spacing(source: Any, destination: Any) -> Any:
    """``w:spacing``.

    The autospacing attributes survive the merge, and ``w:lineRule`` qualifies
    ``w:line``: a source stating ``w:line`` states a rule too (``auto`` when it
    names none), and a source stating no ``w:line`` leaves both inherited
    (CR-015's probe ``styles-linerule``).
    """
    from docx4j_py.wml import CtPprBaseSpacing, STLineSpacingRule

    if source is None:
        return destination
    out = destination
    if source.before_autospacing is not None or source.after_autospacing is not None:
        if out is None:
            out = CtPprBaseSpacing()
        if source.before_autospacing is not None:
            out.before_autospacing = source.before_autospacing
        if source.after_autospacing is not None:
            out.after_autospacing = source.after_autospacing
    if not is_empty_spacing(source):
        if out is None:
            out = CtPprBaseSpacing()
        out.after = _attr(source.after, out.after)
        out.after_lines = _attr(source.after_lines, out.after_lines)
        out.before = _attr(source.before, out.before)
        out.before_lines = _attr(source.before_lines, out.before_lines)
        if source.line is not None:
            out.line = source.line
            out.line_rule = (
                STLineSpacingRule.AUTO if source.line_rule is None else source.line_rule
            )
        elif source.line_rule is not None:
            out.line_rule = source.line_rule
    return out


def merge_tabs(source: Any, destination: Any) -> Any:
    """``w:tabs``.

    The custom tab stops are the union of those declared and those inherited, a
    stop at the same position replaces the inherited one, and ``w:val="clear"``
    only removes (ECMA-376 17.3.1.38, 17.3.1.37). Word holds them in ascending
    position order.
    """
    from docx4j_py.wml import CTTabStop, STTabJc, Tabs

    if is_empty_tabs(source):
        return destination
    out = Tabs() if destination is None else destination
    stops = list(out.tab or ())
    for stop in source.tab:
        pos = stop.pos
        if pos is not None:
            stops = [existing for existing in stops if existing.pos != pos]
        if stop.val is STTabJc.CLEAR:
            continue
        stops.append(CTTabStop(val=stop.val, leader=stop.leader, pos=pos))
    stops.sort(key=lambda s: (s.pos is not None, s.pos if s.pos is not None else 0))
    out.tab[:] = stops
    return out


def merge_frame_pr(source: Any, destination: Any) -> Any:
    """``w:framePr``: each attribute inherits on its own, ``w:anchorLock`` too.

    docx4j's merge passes the stated values (``82c64bf14``): an omitted
    ``w:anchorLock`` keeps the style's, a stated one wins, and where neither
    layer states it none is written. Before that docx4j's merge read it through
    ``isAnchorLock()`` and wrote the old schema default, true, explicitly; this
    port did the same.
    """
    from docx4j_py.wml import CTFramePr

    if is_empty_frame_pr(source):
        return destination
    out = CTFramePr() if destination is None else destination
    out.drop_cap = _attr(source.drop_cap, out.drop_cap)
    out.lines = _attr(source.lines, out.lines)
    out.w = _attr(source.w, out.w)
    out.h = _attr(source.h, out.h)
    out.v_space = _attr(source.v_space, out.v_space)
    out.h_space = _attr(source.h_space, out.h_space)
    out.wrap = _attr(source.wrap, out.wrap)
    out.h_anchor = _attr(source.h_anchor, out.h_anchor)
    out.v_anchor = _attr(source.v_anchor, out.v_anchor)
    out.x = _attr(source.x, out.x)
    out.x_align = _attr(source.x_align, out.x_align)
    out.y = _attr(source.y, out.y)
    out.y_align = _attr(source.y_align, out.y_align)
    out.h_rule = _attr(source.h_rule, out.h_rule)
    out.anchor_lock = _attr(source.anchor_lock, out.anchor_lock)
    return out


def _merge_tbl_width(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import TblWidth

    if source is None or source.w is None:
        return destination
    out = TblWidth() if destination is None else destination
    out.w = source.w
    # "if @w:type is omitted, its value shall be assumed to be dxa"
    out.type_value = source.type_value
    return out


def _merge_tbl_cell_mar(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import CTTblCellMar

    if _is_empty_tbl_cell_mar(source):
        return destination
    out = CTTblCellMar() if destination is None else destination
    out.bottom = _merge_tbl_width(source.bottom, out.bottom)
    out.left = _merge_tbl_width(source.left, out.left)
    out.right = _merge_tbl_width(source.right, out.right)
    out.top = _merge_tbl_width(source.top, out.top)
    return out


def _merge_tbl_borders(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import TblBorders

    if _is_empty_tbl_borders(source):
        return destination
    out = TblBorders() if destination is None else destination
    out.bottom = merge_border(source.bottom, out.bottom)
    out.left = merge_border(source.left, out.left)
    out.right = merge_border(source.right, out.right)
    out.top = merge_border(source.top, out.top)
    out.inside_h = merge_border(source.inside_h, out.inside_h)
    out.inside_v = merge_border(source.inside_v, out.inside_v)
    return out


def _merge_tc_borders(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import CtTcPrInnerTcBorders

    if _is_empty_tc_borders(source):
        return destination
    out = CtTcPrInnerTcBorders() if destination is None else destination
    out.bottom = merge_border(source.bottom, out.bottom)
    out.left = merge_border(source.left, out.left)
    out.right = merge_border(source.right, out.right)
    out.top = merge_border(source.top, out.top)
    out.inside_h = merge_border(source.inside_h, out.inside_h)
    out.inside_v = merge_border(source.inside_v, out.inside_v)
    out.tl2br = merge_border(source.tl2br, out.tl2br)
    out.tr2bl = merge_border(source.tr2bl, out.tr2bl)
    return out


def _merge_tc_mar(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import TcMar

    if _is_empty_tc_mar(source):
        return destination
    out = TcMar() if destination is None else destination
    out.bottom = _merge_tbl_width(source.bottom, out.bottom)
    out.left = _merge_tbl_width(source.left, out.left)
    out.right = _merge_tbl_width(source.right, out.right)
    out.top = _merge_tbl_width(source.top, out.top)
    return out


def _merge_tbl_look(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import CTTblLook

    if _is_empty_tbl_look(source):
        return destination
    out = CTTblLook() if destination is None else destination
    out.first_column = _attr(source.first_column, out.first_column)
    out.first_row = _attr(source.first_row, out.first_row)
    out.last_column = _attr(source.last_column, out.last_column)
    out.last_row = _attr(source.last_row, out.last_row)
    out.no_hband = _attr(source.no_hband, out.no_hband)
    out.no_vband = _attr(source.no_vband, out.no_vband)
    out.val = _attr(source.val, out.val)
    return out


def _merge_tblp_pr(source: Any, destination: Any) -> Any:
    """``w:tblpPr``: the anchors and the ``*Spec`` enums only where the source states them."""
    from docx4j_py.wml import CTTblPPr

    if _is_empty_tblp_pr(source):
        return destination
    out = CTTblPPr() if destination is None else destination
    out.left_from_text = _attr(source.left_from_text, out.left_from_text)
    out.right_from_text = _attr(source.right_from_text, out.right_from_text)
    out.top_from_text = _attr(source.top_from_text, out.top_from_text)
    out.bottom_from_text = _attr(source.bottom_from_text, out.bottom_from_text)
    if source.vert_anchor is not None:
        out.vert_anchor = source.vert_anchor
    if source.horz_anchor is not None:
        out.horz_anchor = source.horz_anchor
    if source.tblp_xspec is not None:
        out.tblp_xspec = source.tblp_xspec
    out.tblp_x = _attr(source.tblp_x, out.tblp_x)
    if source.tblp_yspec is not None:
        out.tblp_yspec = source.tblp_yspec
    out.tblp_y = _attr(source.tblp_y, out.tblp_y)
    return out


def _merge_tbl_layout(source: Any, destination: Any) -> Any:
    from docx4j_py.wml import CTTblLayoutType

    if source is None or source.type_value is None:
        return destination
    out = CTTblLayoutType() if destination is None else destination
    out.type_value = source.type_value
    return out


# ---------------------------------------------------------------- the entries


@dataclasses.dataclass(frozen=True, slots=True)
class Property:
    """One member of a properties element.

    Attributes:
        name: the generated dataclass's attribute name (``r_fonts``).
        xml_name: the schema's member name (``rFonts``), for messages and tests.
        is_formatting: False for a style reference, a revision record or Word's
            condition cache; :func:`has_direct_formatting_of` leaves those out.
        counts_empty: False for the few members docx4j's hand-written
            ``isEmpty`` omits (``w:bidiVisual``, ``w:tblCaption``,
            ``w:tblDescription``), kept so that the tables stay
            one-entry-per-member and the quirk stays visible.
        merge: how a more specific value merges over an inherited one.
        is_empty: whether a value of this member says nothing.
    """

    name: str
    xml_name: str
    merge: Callable[[Any, Any], Any]
    is_empty: Callable[[Any], bool]
    is_formatting: bool = True
    counts_empty: bool = True

    def get(self, owner: Any) -> Any:
        """This member's value on `owner`, or None."""
        return getattr(owner, self.name, None)

    def set(self, owner: Any, value: Any) -> None:
        """Write this member's value on `owner`."""
        setattr(owner, self.name, value)


def _prop(
    name: str,
    xml_name: str,
    merge: Callable[[Any, Any], Any],
    is_empty: Callable[[Any], bool],
    *,
    formatting: bool = True,
    counts_empty: bool = True,
) -> Property:
    return Property(name, xml_name, merge, is_empty, formatting, counts_empty)


def _bool(
    name: str,
    xml_name: str | None = None,
    *,
    formatting: bool = True,
    counts_empty: bool = True,
) -> Property:
    """A ``BooleanDefaultTrue`` member."""
    return Property(name, xml_name or name, merge_bool, is_empty_bool, formatting, counts_empty)


def _w14(name: str, xml_name: str) -> Property:
    """A w14 member (Office 2010 text effects and OpenType features): replaced whole."""
    return Property(name, xml_name, merge_replace, is_null, True, True)


def _val(name: str, xml_name: str, cls_name: str, **kwargs: Any) -> Property:
    """A member whose only content is ``w:val``."""
    return _prop(name, xml_name, _lazy_merge_val(cls_name), is_empty_val, **kwargs)


def _enumerated(name: str, xml_name: str, cls_name: str, read: Callable[[Any], Any]) -> Property:
    """A ``w:val`` leaf whose value is an enumeration Word reads exactly (CR-018 Phase B, rule 2).

    Empty unless `read` gives a member (:func:`_is_empty_enumerated`), so a
    value outside the enumeration does not override a style's; the member is
    what is carried.
    """
    is_empty = _is_empty_enumerated(read)
    return _prop(name, xml_name, _lazy_merge_val(cls_name, is_empty, read), is_empty)


def _lazy_merge_val(
    cls_name: str,
    is_empty: Callable[[Any], bool] = is_empty_val,
    read: Callable[[Any], Any] | None = None,
) -> Callable[[Any, Any], Any]:
    """``_merge_val`` over a class named now and imported on first use.

    The model is not imported at import time: :mod:`docx4j_py.model` must not
    cost the whole object model to import (CR-001 section 14). `read`, where
    given, is what the value is carried as (an enumeration's member).
    """
    merger: list[Callable[[Any, Any], Any]] = []

    def merge(source: Any, destination: Any) -> Any:
        if not merger:
            from docx4j_py import wml

            merger.append(_merge_val(getattr(wml, cls_name), is_empty, read))
        return merger[0](source, destination)

    return merge


#: The members of ``w:rPr`` (17.3.2) and of the paragraph mark's ``w:pPr/w:rPr``,
#: in schema order, plus the w14 extensions the generated classes carry. Not
#: listed: ``w:rPrChange`` and the paragraph mark's ``w:ins`` / ``w:del`` /
#: ``w:moveFrom`` / ``w:moveTo``, which record revisions and are never inherited.
RUN: tuple[Property, ...] = (
    _val("r_style", "rStyle", "RStyle", formatting=False),
    _prop("r_fonts", "rFonts", merge_r_fonts, is_empty_r_fonts),
    _bool("b"),
    _bool("b_cs", "bCs"),
    _bool("i"),
    _bool("i_cs", "iCs"),
    _bool("caps"),
    _bool("small_caps", "smallCaps"),
    _bool("strike"),
    _bool("dstrike"),
    _bool("outline"),
    _bool("shadow"),
    _bool("emboss"),
    _bool("imprint"),
    _bool("no_proof", "noProof"),
    _bool("snap_to_grid", "snapToGrid"),
    _bool("vanish"),
    _bool("web_hidden", "webHidden"),
    _prop("color", "color", merge_color, is_empty_color),
    _val("spacing", "spacing", "CTSignedTwipsMeasure"),
    _val("w", "w", "CTTextScale"),
    _val("kern", "kern", "HpsMeasure"),
    _val("position", "position", "CTSignedHpsMeasure"),
    _val("sz", "sz", "HpsMeasure"),
    _val("sz_cs", "szCs", "HpsMeasure"),
    _enumerated("highlight", "highlight", "Highlight", highlight),
    _prop("u", "u", merge_u, is_empty_u),
    _val("effect", "effect", "CTTextEffect"),
    _prop("bdr", "bdr", merge_border, is_empty_border),
    _prop("shd", "shd", merge_shd, is_empty_shd),
    _prop("fit_text", "fitText", merge_replace, is_null),
    _enumerated("vert_align", "vertAlign", "CTVerticalAlignRun", vertical_alignment),
    _bool("rtl"),
    _bool("cs"),
    _enumerated("em", "em", "CTEm", emphasis_mark),
    _prop("lang", "lang", merge_lang, is_empty_lang),
    _prop("east_asian_layout", "eastAsianLayout", merge_replace, is_null),
    _bool("spec_vanish", "specVanish"),
    _bool("o_math", "oMath"),
    # w14 (Office 2010 text effects and OpenType features)
    _w14("glow", "glow"),
    _w14("schemas_microsoft_com_office_word_2010_wordml_shadow", "shadow"),
    _w14("reflection", "reflection"),
    _w14("text_outline", "textOutline"),
    _w14("text_fill", "textFill"),
    _w14("scene3d", "scene3d"),
    _w14("props3d", "props3d"),
    _w14("ligatures", "ligatures"),
    _w14("num_form", "numForm"),
    _w14("num_spacing", "numSpacing"),
    _w14("stylistic_sets", "stylisticSets"),
    _w14("cntxt_alts", "cntxtAlts"),
)

#: Members of ``w:rPr`` and ``w:pPr/w:rPr`` that are not in :data:`RUN`: the
#: revision records.
RUN_EXCLUDED: tuple[str, ...] = ("r_pr_change", "ins", "del_value", "move_from", "move_to")

#: The **toggle properties** of ``w:rPr``, in the order ECMA-376-1 17.7.3 lists
#: them. Twelve, and no more: ``w:dstrike``, ``w:noProof``, ``w:snapToGrid``,
#: ``w:webHidden``, ``w:rtl``, ``w:cs``, ``w:specVanish`` and ``w:oMath`` are
#: Boolean but are not toggles.
TOGGLE_NAMES: frozenset[str] = frozenset(
    {
        "b",
        "b_cs",
        "caps",
        "emboss",
        "i",
        "i_cs",
        "imprint",
        "outline",
        "shadow",
        "small_caps",
        "strike",
        "vanish",
    }
)

#: :data:`TOGGLE_NAMES` as members of :data:`RUN`, in schema order.
TOGGLES: tuple[Property, ...] = tuple(p for p in RUN if p.name in TOGGLE_NAMES)

#: The ``w:ind`` member's name: the paragraph merge applies it last, after the
#: numbering level's.
IND = "ind"

#: The members of ``w:pPr`` shared with a style's ``w:pPr`` (``PPrBase``,
#: 17.3.1), in schema order. Not formatting: ``w:pStyle`` (the style reference),
#: ``w:divId`` (an HTML import artefact), ``w:cnfStyle`` (Word's cache of the
#: table conditions that apply) and ``w14:collapsed`` (the outline view's state).
#: A ``w:pPr``'s own ``w:rPr`` (the paragraph mark) and ``w:sectPr`` are outside
#: ``PPrBase`` and outside this list.
PARAGRAPH: tuple[Property, ...] = (
    _val("p_style", "pStyle", "CtPprBasePStyle", formatting=False),
    _bool("keep_next", "keepNext"),
    _bool("keep_lines", "keepLines"),
    _bool("page_break_before", "pageBreakBefore"),
    _prop("frame_pr", "framePr", merge_frame_pr, is_empty_frame_pr),
    _bool("widow_control", "widowControl"),
    _prop("num_pr", "numPr", merge_num_pr, is_empty_num_pr),
    _bool("suppress_line_numbers", "suppressLineNumbers"),
    _prop("p_bdr", "pBdr", merge_p_bdr, is_empty_p_bdr),
    _prop("shd", "shd", merge_shd, is_empty_shd),
    _prop("tabs", "tabs", merge_tabs, is_empty_tabs),
    _bool("suppress_auto_hyphens", "suppressAutoHyphens"),
    _bool("kinsoku"),
    _bool("word_wrap", "wordWrap"),
    _bool("overflow_punct", "overflowPunct"),
    _bool("top_line_punct", "topLinePunct"),
    _bool("auto_space_de", "autoSpaceDE"),
    _bool("auto_space_dn", "autoSpaceDN"),
    _bool("bidi"),
    _bool("adjust_right_ind", "adjustRightInd"),
    _bool("snap_to_grid", "snapToGrid"),
    _prop("spacing", "spacing", merge_spacing, is_empty_spacing),
    _prop(IND, "ind", merge_ind, is_empty_ind),
    _bool("contextual_spacing", "contextualSpacing"),
    _bool("mirror_indents", "mirrorIndents"),
    _bool("suppress_overlap", "suppressOverlap"),
    _val("jc", "jc", "Jc"),
    _val("text_direction", "textDirection", "TextDirection"),
    _val("text_alignment", "textAlignment", "CtPprBaseTextAlignment"),
    _val("textbox_tight_wrap", "textboxTightWrap", "CTTextboxTightWrap"),
    _val("outline_lvl", "outlineLvl", "CtPprBaseOutlineLvl"),
    _prop("div_id", "divId", merge_replace, is_null, formatting=False),
    _val("cnf_style", "cnfStyle", "CTCnf", formatting=False),
    _bool("collapsed", "collapsed", formatting=False),
)

#: Members of ``w:pPr`` that are not in :data:`PARAGRAPH`: the mark, the section,
#: the revision.
PARAGRAPH_EXCLUDED: tuple[str, ...] = ("r_pr", "sect_pr", "p_pr_change")

#: The members of ``w:tblPr`` (``CT_TblPrBase``, 17.4.60), in schema order.
#:
#: Three quirks of docx4j's hand-written pair are reproduced and flagged:
#: ``w:bidiVisual`` is applied but not counted by ``isEmpty``, and
#: ``w:tblCaption`` and ``w:tblDescription`` are neither applied nor counted
#: (they are a table's accessibility text, not formatting a style carries).
TABLE: tuple[Property, ...] = (
    _val("tbl_style", "tblStyle", "CtTblPrBaseTblStyle", formatting=False),
    _prop("tblp_pr", "tblpPr", _merge_tblp_pr, _is_empty_tblp_pr),
    _val("tbl_overlap", "tblOverlap", "CTTblOverlap"),
    # applied by docx4j's apply(CTTblPrBase) but not counted by its isEmpty
    _bool("bidi_visual", "bidiVisual", counts_empty=False),
    _val("tbl_style_row_band_size", "tblStyleRowBandSize", "CtTblPrBaseTblStyleRowBandSize"),
    _val("tbl_style_col_band_size", "tblStyleColBandSize", "CtTblPrBaseTblStyleColBandSize"),
    _prop("tbl_w", "tblW", _merge_tbl_width, _is_empty_tbl_width),
    _val("jc", "jc", "Jc"),
    _prop("tbl_cell_spacing", "tblCellSpacing", _merge_tbl_width, _is_empty_tbl_width),
    _prop("tbl_ind", "tblInd", _merge_tbl_width, _is_empty_tbl_width),
    _prop("tbl_borders", "tblBorders", _merge_tbl_borders, _is_empty_tbl_borders),
    _prop("shd", "shd", merge_shd, is_empty_shd),
    _prop("tbl_layout", "tblLayout", _merge_tbl_layout, _is_empty_tbl_layout),
    _prop("tbl_cell_mar", "tblCellMar", _merge_tbl_cell_mar, _is_empty_tbl_cell_mar),
    _prop("tbl_look", "tblLook", _merge_tbl_look, _is_empty_tbl_look),
    _prop("tbl_caption", "tblCaption", _not_carried, is_null, counts_empty=False),
    _prop("tbl_description", "tblDescription", _not_carried, is_null, counts_empty=False),
)

TABLE_EXCLUDED: tuple[str, ...] = ("tbl_pr_change",)

#: The members of ``w:tcPr`` (``TcPr``, 17.4.68), in schema order. ``w:cellIns``,
#: ``w:cellDel``, ``w:cellMerge`` and ``w:tcPrChange`` record revisions and are
#: excluded.
CELL: tuple[Property, ...] = (
    _val("cnf_style", "cnfStyle", "CTCnf", formatting=False),
    _prop("tc_w", "tcW", _merge_tbl_width, _is_empty_tbl_width),
    _val("grid_span", "gridSpan", "CtTcPrInnerGridSpan"),
    _val("h_merge", "hMerge", "CtTcPrInnerHMerge"),
    _val("v_merge", "vMerge", "CtTcPrInnerVMerge"),
    _prop("tc_borders", "tcBorders", _merge_tc_borders, _is_empty_tc_borders),
    _prop("shd", "shd", merge_shd, is_empty_shd),
    _bool("no_wrap", "noWrap"),
    _prop("tc_mar", "tcMar", _merge_tc_mar, _is_empty_tc_mar),
    _val("text_direction", "textDirection", "TextDirection"),
    _bool("tc_fit_text", "tcFitText"),
    _val("v_align", "vAlign", "CTVerticalJc"),
    _bool("hide_mark", "hideMark"),
)

CELL_EXCLUDED: tuple[str, ...] = (
    "cell_ins_or_cell_del_or_cell_merge",
    "tc_pr_change",
)


# ---------------------------------------------------------------- generic operations


def named_property(catalogue: Sequence[Property], name: str) -> Property | None:
    """The member of that name, or None."""
    for entry in catalogue:
        if entry.name == name:
            return entry
    return None


def is_empty_of(catalogue: Sequence[Property], owner: Any) -> bool:
    """Whether every member of `owner` is empty."""
    if owner is None:
        return True
    for entry in catalogue:
        if not entry.counts_empty:
            continue
        if not entry.is_empty(entry.get(owner)):
            return False
    return True


def has_direct_formatting_of(catalogue: Sequence[Property], owner: Any) -> bool:
    """Whether any formatting member is present.

    An element that is there, even with no attributes, counts: it was stated.
    """
    if owner is None:
        return False
    return any(entry.is_formatting and entry.get(owner) is not None for entry in catalogue)


def apply_catalogue(
    catalogue: Sequence[Property],
    source: Any,
    destination: Any,
    except_names: Iterable[str] = (),
) -> None:
    """Apply every member of `source` over `destination`, in place, except those named."""
    skip = frozenset(except_names)
    for entry in catalogue:
        if entry.name in skip:
            continue
        apply_property(entry, source, destination)


def apply_property(entry: Property, source: Any, destination: Any) -> None:
    """Apply one member of `source` over `destination`."""
    entry.set(destination, entry.merge(entry.get(source), entry.get(destination)))


def unset_catalogue(catalogue: Sequence[Property], source: Any, destination: Any) -> None:
    """Unset in `destination` every member that `source` states."""
    for entry in catalogue:
        if entry.get(source) is not None:
            entry.set(destination, None)
