"""The text-in sugar: ``p``, ``r``, ``t``, ``tbl``, ``tr``, ``tc``, ``br``, ``tab``.

CR-001 section 6.2: "The sugar is hand-written and small; it is the text-in
convenience, not a second factory." CR-003 Phase A added the row and cell
builders ``tr`` and ``tc`` (``tbl`` is written over them), the ``EG_RPrBase``
pair :func:`rpr_to_elements` / :func:`rpr_from_elements`, and, in the sibling
modules :mod:`docx4j_py.wml.pictures` and :mod:`docx4j_py.wml.sdt`, the inline
picture and content-control builders. Everything here is built on
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

import dataclasses
from typing import Any

from docx4j_py.child import ChildList, deep_copy
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
    "RPR_BASE_FIELDS",
    "RUN_OPTIONS",
    "UNDERLINE",
    "BuilderError",
    "RPrElement",
    "apply_run_options",
    "br",
    "highlight_hex_value",
    "highlight_name_for_color",
    "p",
    "r",
    "read_run_options",
    "rpr_from_elements",
    "rpr_to_elements",
    "t",
    "tab",
    "tbl",
    "tc",
    "tr",
]


class BuilderError(ValueError):
    """A builder was given something it cannot make an element of.

    CR-003 section 3.1 asks for one error hierarchy in which every message says
    what to do instead, and gives each error a stable :attr:`code` and a
    :attr:`hint` an agent can act on. Phase B builds that hierarchy
    (``Docx4JError`` / ``ContentError``) over the parts layer; Phase A is the
    tree only, so this is its stand-in, and it derives from ``ValueError`` so
    that code written against either spelling keeps working. Phase B re-roots
    it; the ``code`` strings do not change.

    Attributes:
        code: a stable string such as ``"sdt.form_mismatch"``.
        hint: one sentence saying what to do instead.
    """

    def __init__(self, message: str, *, code: str, hint: str) -> None:
        """Build the error from its message, its stable code and its hint."""
        super().__init__(f"{message} ({hint})")
        self.code = code
        self.hint = hint
        self.message = message


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


def _blocks_of(blocks: Any) -> list[Any]:
    """A cell's or a row's content as a list: a string is one paragraph."""
    if isinstance(blocks, str):
        return [p(blocks)]
    if isinstance(blocks, (list, tuple)):
        return list(blocks)
    return [blocks]


def tc(blocks: Any = "", *, width: int | None = None, span: int | None = None) -> Tc:
    """A table cell (``w:tc``).

    Args:
        blocks: a string (one paragraph of text), a block-level object
            (:class:`P`, :class:`Tbl`, a ``w:sdt``), or a list of them.
        width: the cell width in twips, written as ``w:tcW`` of type ``dxa``;
            none when absent.
        span: the number of grid columns the cell spans (``w:gridSpan``).

    An empty cell still gets the ``w:p`` Word requires: a ``w:tc`` whose last
    child is not a paragraph makes Word repair the document.
    """
    content = _blocks_of(blocks)
    if not content:
        content = [el.p()]
    cell = Tc(content=ChildList(content))
    if width is not None or span is not None:
        cell.tc_pr = el.tcPr()
        if width is not None:
            cell.tc_pr.tc_w = el.tcW(w=width, type_value="dxa")
        if span is not None:
            cell.tc_pr.grid_span = el.gridSpan(val=span)
    return cell


def tr(cells: list[Any], *, widths: list[int] | None = None, header: bool = False) -> Tr:
    """A table row (``w:tr``) of cells.

    Args:
        cells: one entry per cell: a string, a block-level object, a list of
            them (all three go through :func:`tc`), or a :class:`Tc` built
            already, which is taken as it is.
        widths: the cell widths in twips, one per column; a missing entry
            leaves that cell without a ``w:tcW``.
        header: repeat this row at the top of every page (``w:trPr`` with
            ``w:tblHeader``), which is what Word's "repeat header rows" sets.
    """
    row = Tr(
        content=ChildList(
            cell
            if isinstance(cell, Tc)
            else tc(cell, width=None if widths is None or i >= len(widths) else widths[i])
            for i, cell in enumerate(cells)
        )
    )
    if header:
        row.tr_pr = el.trPr(content=ChildList([el.tblHeader()]))
    return row


def tbl(
    rows: list[list[Any]],
    *,
    style: str | None = None,
    widths: list[int] | None = None,
    width: int = DEFAULT_TABLE_WIDTH,
) -> Tbl:
    """A table of cells, one paragraph per cell, with a grid.

    Built over :func:`tr` and :func:`tc`, whose output for a string cell of a
    known width is what this used to write by hand.

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
        table.content.append(
            tr([row[i] if i < len(row) else "" for i in range(columns)], widths=resolved)
        )
    return table


# ---------------------------------------------------------------------------
# run properties as a list of elements (EG_RPrBase)
# ---------------------------------------------------------------------------


def _rpr_base_fields() -> tuple[tuple[str, str], ...]:
    """The ``EG_RPrBase`` members in schema order, as (field name, qname).

    Taken from the model rather than from a list written here: the members are
    the fields ``RPr`` and ``CtRprChangeRPr`` have in common, in ``RPr``'s
    declaration order, which is the schema's. That is 51 members, the twelve
    w14 text effects included, and it excludes ``w:rPrChange`` itself (only
    ``RPr`` has it) and the four ``ParaRPr`` revision marks.
    """
    from docx4j_py.wml import CtRprChangeRPr

    shared = {f.name for f in dataclasses.fields(CtRprChangeRPr)}
    out = []
    for field in dataclasses.fields(RPr):
        if field.name not in shared:
            continue
        namespace = field.metadata.get("namespace")
        local = field.metadata.get("name") or field.name
        out.append((field.name, f"{{{namespace}}}{local}" if namespace else local))
    return tuple(out)


#: ``EG_RPrBase``: the 51 (field name, qualified name) pairs, in schema order.
RPR_BASE_FIELDS: tuple[tuple[str, str], ...] = _rpr_base_fields()

_RPR_FIELD_BY_NAME: dict[str, str] = {}
for _name, _qname in RPR_BASE_FIELDS:
    _RPR_FIELD_BY_NAME[_name] = _name
    _RPR_FIELD_BY_NAME[_qname] = _name
    _RPR_FIELD_BY_NAME[_qname.rpartition("}")[2]] = _name
del _name, _qname


@dataclasses.dataclass(frozen=True, slots=True)
class RPrElement:
    """One member of ``EG_RPrBase``: its field name, its element name, its value.

    The element name is carried beside the value because the value alone does
    not identify it: ``w:b``, ``w:i``, ``w:caps`` and fifteen more are all a
    :class:`BooleanDefaultTrue`, so a bare list of property objects could not
    be turned back into a ``w:rPr``.
    """

    name: str
    """The model's field name (``b_cs``, ``vert_align``, ``text_outline``)."""
    qname: str
    """The element name, ``{namespace}localName`` (``{…}bCs``)."""
    value: Any
    """The property object (``BooleanDefaultTrue``, ``RFonts``, ``CTGlow``…)."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the two names and the value's class."""
        return {
            "name": self.name,
            "qname": self.qname,
            "value": type(self.value).__name__,
        }


def rpr_to_elements(rpr: Any) -> list[RPrElement]:
    """Run properties as a list of ``EG_RPrBase`` elements, in schema order.

    Takes a ``w:rPr`` in any of the four shapes the model gives it ---
    :class:`RPr` (a run's), :class:`ParaRPr` (a paragraph mark's),
    ``CtRprChangeRPr`` (``w:rPrChange/w:rPr``) or ``CTParaRPrOriginal``
    (``w:pPr/w:rPr/w:rPrChange/w:rPr``) --- and returns each property that is
    present as an :class:`RPrElement`, deep-copied so that the result can be
    put straight into another tree. ``w:rPrChange`` itself is not a member of
    ``EG_RPrBase`` and is never returned; neither are ``ParaRPr``'s
    ``w:ins`` / ``w:del`` / ``w:moveFrom`` / ``w:moveTo`` revision marks.

    :func:`rpr_from_elements` is the inverse. The pair is what a tracked
    formatting change needs (CR-003 section 3.8: ``w:rPrChange`` records the
    run properties as they were).
    """
    out: list[RPrElement] = []
    if rpr is None:
        return out
    for name, qname in RPR_BASE_FIELDS:
        value = getattr(rpr, name, None)
        if value is None:
            continue
        items = value if isinstance(value, (list, tuple)) else (value,)
        for item in items:
            if item is not None:
                out.append(RPrElement(name, qname, deep_copy(item)))
    return out


def rpr_from_elements(elements: Any, *, cls: type = RPr) -> Any:
    """Build a ``w:rPr`` of `cls` from ``EG_RPrBase`` elements.

    The inverse of :func:`rpr_to_elements`. `elements` is what that returned,
    or a sequence of ``(name, value)`` pairs (the name may be the field name,
    the local name or the qualified name), or another ``w:rPr`` object of any
    of the four shapes, which is converted first. `cls` is the class to build:
    :class:`RPr` by default, ``CtRprChangeRPr`` for the inside of a
    ``w:rPrChange``, :class:`ParaRPr` or ``CTParaRPrOriginal`` for a paragraph
    mark's.

    Where the target class keeps a member as a list (the generated
    ``w:rPrChange/w:rPr`` classes keep every one of them as a list, because
    ``EG_RPrBase`` is an unbounded group there) the value is appended; where it
    keeps a single value the last one given wins. Order is not the caller's
    problem: the model declares these as named fields in schema order, so the
    serialiser writes them in ``EG_RPrBase`` order whatever order they arrive.

    Raises:
        BuilderError: if an item is not an :class:`RPrElement`, a pair, or a
            member of ``EG_RPrBase``.
    """
    if elements is None:
        return cls()
    if not isinstance(elements, (list, tuple)):
        elements = rpr_to_elements(elements)

    target = cls()
    for item in elements:
        if isinstance(item, RPrElement):
            name, value = item.name, item.value
        elif isinstance(item, tuple) and len(item) == 2:
            name, value = item
        else:
            raise BuilderError(
                f"not an EG_RPrBase element: {type(item).__name__}",
                code="rpr.not_an_element",
                hint="pass what rpr_to_elements returned, or (name, value) pairs",
            )
        field_name = _RPR_FIELD_BY_NAME.get(name)
        if field_name is None:
            raise BuilderError(
                f"not a member of EG_RPrBase: {name!r}",
                code="rpr.unknown_property",
                hint="RPR_BASE_FIELDS lists the 51 names, w:rPrChange is not one of them",
            )
        current = getattr(target, field_name, None)
        if isinstance(current, list):
            current.append(deep_copy(value))
        else:
            setattr(target, field_name, deep_copy(value))
    return target
