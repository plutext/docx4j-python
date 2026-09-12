"""The text-in sugar: ``p``, ``r``, ``t``, ``tbl``, ``br``, ``tab``.

CR-001 section 6.2: "The sugar is hand-written and small; it is the text-in
convenience, not a second factory." Everything here is built on
:mod:`docx4j_py.wml.el`, so there is one source of element names, and everything
it returns is a :class:`docx4j_py.child.Child` whose parents are linked as the
tree is assembled (``ChildList`` does that).

The run options are the ones ``docx4j-generated-objects-ts``'s
``builders/wml.mts`` supports, in the same vocabulary --- which is Office JS's
``Word.Font``, not the schema's, so that what a builder sets a view reads back
under the same name --- spelled ``snake_case``:

===================== =========================================================
``bold``              ``w:b`` and ``w:bCs``
``italic``            ``w:i`` and ``w:iCs``
``underline``         ``True`` is ``single``; a name from :data:`UNDERLINE`;
                      ``False``/``'None'`` removes ``w:u``
``strike_through``    ``w:strike``
``double_strike_through`` ``w:dstrike``
``subscript`` / ``superscript`` ``w:vertAlign``
``name`` (``font``)   ``w:rFonts`` ascii and hAnsi, clearing the theme fonts
``size``              points; ``w:sz`` and ``w:szCs`` are half-points
``color``             ``'#RRGGBB'`` or ``'auto'``
``highlight_color``   a highlight name or one of its ``#RRGGBB`` values
``style``             ``w:rStyle``
===================== =========================================================

``None`` leaves a property alone; ``False``, ``''`` and ``'None'`` remove it.
``font`` is accepted as an alias of ``name``, which reads better as a keyword
argument in Python.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.child import ChildList
from docx4j_py.wml import (
    RT,
    P,
    R,
    RPr,
    STBrType,
    Tbl,
    Tc,
    Tr,
    el,
)

__all__ = [
    "HIGHLIGHT_COLORS",
    "RUN_OPTIONS",
    "UNDERLINE",
    "apply_run_options",
    "br",
    "highlight_hex_value",
    "highlight_name_for_color",
    "p",
    "r",
    "read_run_options",
    "t",
    "tab",
    "tbl",
]

#: The underline names Office JS uses, and the ``w:u`` value each one writes.
#: The seventeen entries of ``builders/wml.mts``'s ``UNDERLINE_TO_WML``.
UNDERLINE: dict[str, str] = {
    "Single": "single",
    "Word": "words",
    "Double": "double",
    "Thick": "thick",
    "Dotted": "dotted",
    "DottedHeavy": "dottedHeavy",
    "DashLine": "dash",
    "DashLineHeavy": "dashedHeavy",
    "DashLineLong": "dashLong",
    "DashLineLongHeavy": "dashLongHeavy",
    "DotDashLine": "dotDash",
    "DotDashLineHeavy": "dashDotHeavy",
    "TwoDotDashLine": "dotDotDash",
    "TwoDotDashLineHeavy": "dashDotDotHeavy",
    "Wave": "wave",
    "WaveHeavy": "wavyHeavy",
    "WaveDouble": "wavyDouble",
}
_WML_TO_UNDERLINE = {v: k for k, v in UNDERLINE.items()}

#: ``w:highlight`` names and their sRGB values, as docx4j has them
#: (``Highlight.getHexVal``; ``darkYellow`` is gold).
HIGHLIGHT_COLORS: tuple[tuple[str, str], ...] = (
    ("black", "000000"),
    ("blue", "0000FF"),
    ("cyan", "00FFFF"),
    ("green", "008000"),
    ("magenta", "FF00FF"),
    ("red", "FF0000"),
    ("yellow", "FFFF00"),
    ("white", "FFFFFF"),
    ("darkBlue", "00008B"),
    ("darkCyan", "008B8B"),
    ("darkGreen", "006400"),
    ("darkMagenta", "8B008B"),
    ("darkRed", "8B0000"),
    ("darkYellow", "FFD700"),
    ("darkGray", "A9A9A9"),
    ("lightGray", "D3D3D3"),
)

#: Every run option :func:`apply_run_options` understands.
RUN_OPTIONS: frozenset[str] = frozenset(
    {
        "bold",
        "italic",
        "underline",
        "strike_through",
        "double_strike_through",
        "subscript",
        "superscript",
        "name",
        "size",
        "color",
        "highlight_color",
        "style",
    }
)


def highlight_hex_value(name: str | None) -> str | None:
    """``'#RRGGBB'`` for a highlight name (docx4j ``Highlight.getHexVal``)."""
    for entry, hex_value in HIGHLIGHT_COLORS:
        if entry == name:
            return f"#{hex_value}"
    return None


def highlight_name_for_color(color: str | None) -> str | None:
    """The highlight name for a colour, or None (docx4j ``Highlight.setVal``).

    A known name comes back as it is, ``#RRGGBB`` is looked up; anything else
    --- ``rgb(...)``, an arbitrary hex --- is not a highlight and gets None.
    """
    if color is None:
        return None
    if any(name == color for name, _ in HIGHLIGHT_COLORS):
        return color
    trimmed = color.strip()
    if trimmed.startswith("#"):
        wanted = trimmed[1:].upper()
        for name, hex_value in HIGHLIGHT_COLORS:
            if hex_value == wanted:
                return name
    return None


# ---------------------------------------------------------------------------
# run properties
# ---------------------------------------------------------------------------


def _on(value: Any) -> bool:
    """A ``BooleanDefaultTrue`` toggle is on when present and not ``val=0``."""
    return value is not None and getattr(value, "val", None) is not False


def apply_run_options(r_pr: RPr, **options: Any) -> RPr:
    """Write run options onto run properties, in place.

    The single mapping from the option names to ``w:rPr``, the counterpart of
    ``builders/wml.mts``'s ``applyRunOptions``; :func:`read_run_options` is its
    inverse. A content API (CR-003) calls the pair so that a property it sets is
    the property it reads back.
    """
    if "font" in options:
        options.setdefault("name", options.pop("font"))
    unknown = set(options) - RUN_OPTIONS
    if unknown:
        raise TypeError(f"unknown run option(s): {', '.join(sorted(unknown))}")

    bold = options.get("bold")
    if bold is not None:
        r_pr.b = el.b() if bold else None
        r_pr.b_cs = el.bCs() if bold else None

    italic = options.get("italic")
    if italic is not None:
        r_pr.i = el.i() if italic else None
        r_pr.i_cs = el.iCs() if italic else None

    underline = options.get("underline")
    if underline is not None:
        value = "Single" if underline is True else "None" if underline is False else underline
        if value in ("None", "Mixed"):
            r_pr.u = None
        else:
            try:
                r_pr.u = el.u(val=UNDERLINE[value])
            except KeyError:
                raise ValueError(
                    f"not an underline style: {value!r} (one of {', '.join(sorted(UNDERLINE))})"
                ) from None

    strike = options.get("strike_through")
    if strike is not None:
        r_pr.strike = el.strike() if strike else None

    dstrike = options.get("double_strike_through")
    if dstrike is not None:
        r_pr.dstrike = el.dstrike() if dstrike else None

    subscript = options.get("subscript")
    if subscript is not None:
        if subscript:
            r_pr.vert_align = el.vertAlign(val="subscript")
        elif getattr(r_pr.vert_align, "val", None) == "subscript":
            r_pr.vert_align = None

    superscript = options.get("superscript")
    if superscript is not None:
        if superscript:
            r_pr.vert_align = el.vertAlign(val="superscript")
        elif getattr(r_pr.vert_align, "val", None) == "superscript":
            r_pr.vert_align = None

    name = options.get("name")
    if name is not None:
        if name == "":
            r_pr.r_fonts = None
        else:
            fonts = r_pr.r_fonts or el.rFonts()
            fonts.ascii = name
            fonts.h_ansi = name
            fonts.ascii_theme = None
            fonts.h_ansi_theme = None
            r_pr.r_fonts = fonts

    size = options.get("size")
    if size is not None:
        half = round(size * 2)
        r_pr.sz = el.sz(val=half)
        r_pr.sz_cs = el.szCs(val=half)

    color = options.get("color")
    if color is not None:
        hex_value = color[1:] if color.startswith("#") else color
        if hex_value == "":
            r_pr.color = None
        else:
            r_pr.color = el.color(val="auto" if hex_value.lower() == "auto" else hex_value.upper())

    highlight = options.get("highlight_color")
    if "highlight_color" in options:
        if highlight in (None, ""):
            r_pr.highlight = None
        else:
            found = highlight_name_for_color(highlight)
            if found is None:
                raise ValueError(
                    f"not a highlight colour: {highlight!r} "
                    "(use a highlight name or one of its #RRGGBB values)"
                )
            r_pr.highlight = el.highlight(val=found)

    style = options.get("style")
    if style is not None:
        r_pr.r_style = None if style == "" else el.rStyle(val=style)

    return r_pr


def read_run_options(r_pr: RPr | None) -> dict[str, Any]:
    """Read the direct run formatting back, in the same vocabulary.

    The inverse of :func:`apply_run_options`, and the shape
    ``builders/wml.mts``'s ``readRunOptions`` returns: every option present,
    ``underline`` narrowed to a name, ``name``/``color``/``style`` empty strings
    and ``size`` zero where the run says nothing.
    """
    fonts = getattr(r_pr, "r_fonts", None)
    underline_value = getattr(getattr(r_pr, "u", None), "val", None)
    underline_value = getattr(underline_value, "value", underline_value)
    color_value = getattr(getattr(r_pr, "color", None), "val", None)
    highlight = getattr(getattr(r_pr, "highlight", None), "val", None)
    highlight = getattr(highlight, "value", highlight)
    vert = getattr(getattr(r_pr, "vert_align", None), "val", None)
    vert = getattr(vert, "value", vert)
    size = getattr(getattr(r_pr, "sz", None), "val", None)

    is_hex = (
        isinstance(color_value, str)
        and len(color_value) == 6
        and all(c in "0123456789abcdefABCDEF" for c in color_value)
    )
    return {
        "bold": _on(getattr(r_pr, "b", None)),
        "italic": _on(getattr(r_pr, "i", None)),
        "underline": (
            "None"
            if underline_value in (None, "none")
            else _WML_TO_UNDERLINE.get(underline_value, "Mixed")
        ),
        "strike_through": _on(getattr(r_pr, "strike", None)),
        "double_strike_through": _on(getattr(r_pr, "dstrike", None)),
        "subscript": vert == "subscript",
        "superscript": vert == "superscript",
        "name": (fonts.ascii or fonts.h_ansi or "") if fonts is not None else "",
        "size": 0 if size is None else size / 2,
        "color": (
            ""
            if color_value is None
            else f"#{color_value.upper()}"
            if is_hex
            else str(color_value)
        ),
        "highlight_color": (
            None if highlight in (None, "none") else (highlight_hex_value(highlight) or highlight)
        ),
        "style": getattr(getattr(r_pr, "r_style", None), "val", None) or "",
    }


# ---------------------------------------------------------------------------
# the builders
# ---------------------------------------------------------------------------


def t(text: str) -> RT:
    """``w:t``, with ``xml:space="preserve"`` where the whitespace needs it."""
    return el.t(text)


def br(type: str | STBrType | None = None) -> Any:  # noqa: A002 - the schema's name
    """``w:br``. A text-wrapping break carries no type attribute, as in Word."""
    if type in (None, "textWrapping", STBrType.TEXT_WRAPPING):
        return el.br()
    return el.br(type_value=type)


def tab() -> Any:
    """``w:tab``, the run-level one (``R.Tab`` in docx4j, ``RTab`` here)."""
    return el.tab()


def _run_content(pieces: tuple[Any, ...]) -> list[Any]:
    """Text becomes ``w:t``/``w:tab``/``w:br``; anything else is passed through."""
    content: list[Any] = []
    for piece in pieces:
        if not isinstance(piece, str):
            content.append(piece)
            continue
        buffer = ""
        for ch in piece:
            if ch in "\t\n":
                if buffer:
                    content.append(t(buffer))
                    buffer = ""
                content.append(tab() if ch == "\t" else br())
            else:
                buffer += ch
        if buffer:
            content.append(t(buffer))
    return content


def r(*text_or_children: Any, **options: Any) -> R:
    """A run.

    ``r('hello')`` is a run of text; ``r('a\\tb')`` turns the tab into ``w:tab``
    and a newline into ``w:br``, which docx4j's ``addParagraphOfText`` does not
    do but users expect. Anything that is not a string --- a ``w:drawing``, a
    ``w:sym``, another ``el`` result --- is put into the run as it is.

    Every run option of :func:`apply_run_options` is accepted as a keyword
    argument; a run with none gets no ``w:rPr`` at all.
    """
    run = R(content=ChildList(_run_content(text_or_children)))
    if options:
        r_pr = apply_run_options(RPr(), **options)
        if any(getattr(r_pr, f, None) is not None for f in _RPR_FIELDS):
            run.r_pr = r_pr
    return run


_RPR_FIELDS = (
    "r_style",
    "r_fonts",
    "b",
    "b_cs",
    "i",
    "i_cs",
    "strike",
    "dstrike",
    "color",
    "sz",
    "sz_cs",
    "highlight",
    "u",
    "vert_align",
)


def p(*runs_or_text: Any, style: str | None = None, **ppr: Any) -> P:
    """A paragraph.

    ``p('Hello')`` makes the run and the text; ``p(r('a', bold=True), r('b'))``
    takes runs; the two mix. ``style`` is the *paragraph* style (``w:pStyle``);
    a run style on the text form goes through ``run_style``. Anything else in
    ``**ppr`` is set on a ``w:pPr``, so ``p('x', keep_next=el.keepNext())``
    works and every ``PPr`` field is reachable without leaving the sugar.
    """
    run_options = {k: ppr.pop(k) for k in list(ppr) if k in RUN_OPTIONS or k == "font"}
    if "run_style" in ppr:
        run_options["style"] = ppr.pop("run_style")

    content: list[Any] = []
    for item in runs_or_text:
        content.append(r(item, **run_options) if isinstance(item, str) else item)

    paragraph = P(content=ChildList(content))
    if style is not None or ppr:
        p_pr = el.pPr(**ppr)
        if style is not None:
            p_pr.p_style = el.pStyle(val=style)
        paragraph.p_pr = p_pr
    return paragraph


#: The width of the text column on A4 with 2.54 cm margins, in twips: what
#: ``tbl`` spreads its columns over when it is told nothing else.
DEFAULT_TABLE_WIDTH = 9026


def tbl(
    rows: list[list[Any]],
    *,
    style: str | None = None,
    widths: list[int] | None = None,
    width: int = DEFAULT_TABLE_WIDTH,
) -> Tbl:
    """A table of cells, one paragraph per cell, with a grid.

    Args:
        rows: a row per list; a cell is a string, a :class:`P`, or a list of
            block-level objects to put into the cell as they are.
        style: the table style (``w:tblStyle``).
        widths: column widths in twips; equal columns over `width` otherwise,
            the last column taking the rounding remainder so that the grid sums
            to the table width exactly.
        width: the total width in twips when `widths` is absent.
    """
    columns = (
        max(1, len(widths or ()), *(len(row) for row in rows))
        if rows
        else max(1, len(widths or ()))
    )
    equal = width // columns
    resolved = [
        (
            widths[i]
            if widths is not None and i < len(widths)
            else (width - equal * (columns - 1) if i == columns - 1 else equal)
        )
        for i in range(columns)
    ]

    tbl_pr = el.tblPr(tbl_w=el.tblW(w=sum(resolved), type_value="dxa"))
    if style is not None:
        tbl_pr.tbl_style = el.tblStyle(val=style)

    table = Tbl(
        tbl_pr=tbl_pr,
        tbl_grid=el.tblGrid(grid_col=ChildList(el.gridCol(w=w) for w in resolved)),
    )
    for row in rows:
        cells: list[Any] = []
        for index in range(columns):
            cell = row[index] if index < len(row) else ""
            if isinstance(cell, str):
                blocks: list[Any] = [p(cell)]
            elif isinstance(cell, list):
                blocks = list(cell)
            else:
                blocks = [cell]
            cells.append(
                Tc(
                    tc_pr=el.tcPr(tc_w=el.tcW(w=resolved[index], type_value="dxa")),
                    content=ChildList(blocks),
                )
            )
        table.content.append(Tr(content=ChildList(cells)))
    return table
