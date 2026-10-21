"""OMML to LaTeX (WD REQ-604). CR-011 section 6.

docx4j's ``org.docx4j.model.math.OmmlToLatex`` at ``198d015a3`` (its CR-046
D3): the reverse of docx4j-markdown's ``LatexToOmml``, for the same subset.
Anything outside it raises :class:`OmmlMathError`, and a reader then writes
the equation as its runs' text.

The LaTeX is normalised: braced arguments (``U^{3}``, ``\\frac{1}{2}``), the
shortest symbol command on a collision (``\\le``, not ``\\leq``), display math
on one line.

Characters are visited as Java visits them, a UTF-16 code unit at a time: a
character outside the Basic Multilingual Plane is never a letter for the space
a command needs after it, and never a one-character script base.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docx4j_py.model.math.latex_symbols import FUNCTION_NAMES, NARY, SYMBOLS

__all__ = ["OmmlMathError", "OmmlToLatex"]


class OmmlMathError(Exception):
    """An equation outside the converter's subset. docx4j ``OmmlMathException``."""


def _shorter(a: str, b: str) -> bool:
    return len(a) < len(b) or (len(a) == len(b) and a < b)


def _reverse(table: dict[str, str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for command, symbol in table.items():
        existing = out.get(symbol)
        if existing is None or _shorter(command, existing):
            out[symbol] = command
    return out


#: Unicode to the shortest LaTeX command.
REVERSE_SYMBOLS = _reverse(SYMBOLS)
REVERSE_NARY = _reverse(NARY)
REVERSE_ACCENTS = {
    "̂": "hat",
    "̃": "tilde",
    "̅": "bar",
    "⃗": "vec",
    "̇": "dot",
    "̈": "ddot",
    "̌": "check",
    "̆": "breve",
    "́": "acute",
    "̀": "grave",
}


def _is_letter(c: str) -> bool:
    """Java's ``Character.isLetter`` of a UTF-16 unit: a character beyond the BMP's first unit is a surrogate."""
    return ord(c) <= 0xFFFF and c.isalpha()


def _utf16_length(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2


def _java_trim(s: str) -> str:
    """Java's ``String.trim``: every character up to U+0020 off both ends."""
    start, end = 0, len(s)
    while start < end and s[start] <= " ":
        start += 1
    while end > start and s[end - 1] <= " ":
        end -= 1
    return s[start:end]


class _LatexBuilder:
    """Accumulates LaTeX, spacing a command from a letter that follows it."""

    __slots__ = ("command_open", "parts")

    def __init__(self) -> None:
        self.parts: list[str] = []
        #: The last append was a ``\\letters`` command.
        self.command_open = False

    def command(self, name: str) -> None:
        self.parts.append("\\" + name)
        self.command_open = _is_letter(name[-1])

    def raw(self, s: str) -> None:
        if not s:
            return
        if self.command_open and _is_letter(s[0]):
            self.parts.append(" ")
        self.parts.append(s)
        self.command_open = False

    def __str__(self) -> str:
        return "".join(self.parts)


_CLASSES: dict[str, type] | None = None


def _classes() -> dict[str, type]:
    global _CLASSES
    if _CLASSES is None:
        from docx4j_py import math, wml

        names = (
            "CTR", "CTF", "CTSSub", "CTSSup", "CTSSubSup", "CTRad", "CTNary", "CTD", "CTEqArr", "CTM", "CTGroupChr",
            "CTLimLow", "CTLimUpp", "CTBorderBox", "CTAcc", "CTBar", "CTOMath", "CTOMathPara", "CTText", "CTOnOff",
            "CTScript", "CTStyle",
        )  # fmt: skip
        found = {name: getattr(math, name) for name in names}
        found["CTMathRunTrackChange"] = wml.CTMathRunTrackChange
        found["texts"] = (wml.T, wml.InstrText, wml.DelInstrText)
        _CLASSES = found
    return _CLASSES


def _value(element: Any) -> Any:
    """An element's ``val``, an enumeration's value."""
    if element is None:
        return None
    val = getattr(element, "val", None)
    return getattr(val, "value", val)


def accepted(item: Any) -> bool:
    """The default reading: as Accept All leaves the equation (a ``w:del``'s content left out)."""
    c = _classes()
    return not (isinstance(item, c["CTMathRunTrackChange"]) and _is_deletion(item))


def _is_deletion(change: Any) -> bool:
    from docx4j_py.model.revisions import trees as rtrees

    return rtrees.name(change) == "del"


class OmmlToLatex:
    """OMML to LaTeX for docx4j's subset. docx4j ``OmmlToLatex``."""

    __slots__ = ("_keep",)

    def __init__(self, keep: Callable[[Any], bool] | None = None) -> None:
        """A converter reading the members `keep` admits: each math element, each piece of a run's
        content, each ``w:ins`` or ``w:del`` in a run. By default all but a ``w:del``'s content."""
        self._keep = keep or accepted

    def keep(self, keep: Callable[[Any], bool] | None) -> OmmlToLatex:
        """Which members of the equation's lists to read (the text model gives its view's rule)."""
        self._keep = keep or accepted
        return self

    def convert(self, value: Any) -> str:
        """An ``m:oMath`` or an ``m:oMathPara``."""
        if isinstance(value, _classes()["CTOMathPara"]):
            return self.convert_o_math_para(value)
        return self.convert_o_math(value)

    def convert_o_math(self, o_math: Any) -> str:
        builder = _LatexBuilder()
        self._elements(o_math.content, builder, False)
        return _java_trim(str(builder))

    def convert_o_math_para(self, para: Any) -> str:
        """All the para's equations, a space apart (display math re-exports on one line)."""
        return " ".join(self.convert_o_math(o_math) for o_math in para.o_math or ())

    # -- the walk ------------------------------------------------------------------

    def _elements(self, elements: list | None, builder: _LatexBuilder, in_eq_arr: bool) -> None:
        c = _classes()
        for u in elements or ():
            if isinstance(u, str):
                continue
            if not self._keep(u):
                continue
            if isinstance(u, c["CTR"]):
                self._run(u, builder, in_eq_arr)
            elif isinstance(u, c["CTF"]):
                builder.command("frac")
                self._braced(u.num, builder)
                self._braced(u.den, builder)
            elif isinstance(u, c["CTSSubSup"]):
                self._base(u.e, builder)
                builder.raw("_")
                self._braced(u.sub, builder)
                builder.raw("^")
                self._braced(u.sup, builder)
            elif isinstance(u, c["CTSSub"]):
                self._base(u.e, builder)
                builder.raw("_")
                self._braced(u.sub, builder)
            elif isinstance(u, c["CTSSup"]):
                self._base(u.e, builder)
                builder.raw("^")
                self._braced(u.sup, builder)
            elif isinstance(u, c["CTRad"]):
                self._rad(u, builder)
            elif isinstance(u, c["CTNary"]):
                self._nary(u, builder)
            elif isinstance(u, c["CTD"]):
                self._delimited(u, builder)
            elif isinstance(u, c["CTEqArr"]):
                self._rows("aligned", u, builder)
            elif isinstance(u, c["CTM"]):
                self._matrix("matrix", u, builder)
            elif isinstance(u, c["CTGroupChr"]):
                self._group_chr(u, builder)
            elif isinstance(u, c["CTLimLow"]):
                builder.command("underset")
                self._braced(u.lim, builder)
                self._braced(u.e, builder)
            elif isinstance(u, c["CTBorderBox"]):
                builder.command("boxed")
                self._braced(u.e, builder)
            elif isinstance(u, c["CTLimUpp"]):
                self._lim_upp(u, builder)
            elif isinstance(u, c["CTAcc"]):
                self._acc(u, builder)
            elif isinstance(u, c["CTBar"]):
                pr = u.bar_pr
                top = pr is not None and pr.pos is not None and _value(pr.pos) == "top"
                builder.command("overline" if top else "underline")
                self._braced(u.e, builder)
            elif isinstance(u, c["CTOMath"]):
                self._elements(u.content, builder, in_eq_arr)
            else:
                raise OmmlMathError(f"unsupported OMML element {type(u).__name__}")

    def _collect(self, r: Any) -> tuple[Any, Any, str]:
        """A run's ``m:rPr``, ``w:rPr`` and text, a ``w:ins`` or ``w:del`` in it read by the keep rule."""
        c = _classes()
        r_pr = r.r_pr
        wml_r_pr = r.r_pr_1
        text: list[str] = []

        def collect(content: list | None) -> None:
            nonlocal r_pr, wml_r_pr
            for u in content or ():
                if isinstance(u, str) or not self._keep(u):
                    continue
                if isinstance(u, c["CTMathRunTrackChange"]):
                    if r_pr is None:
                        r_pr = getattr(u, "r_pr", None)
                    if wml_r_pr is None:
                        wml_r_pr = getattr(u, "schemas_openxmlformats_org_wordprocessingml_2006_main_r_pr", None)
                    collect(u.content)
                elif isinstance(u, (c["CTText"], *c["texts"])):
                    if u.value:
                        text.append(u.value)
                # other run content (breaks and the like) is dropped

        collect(r.content)
        return r_pr, wml_r_pr, "".join(text)

    def _run(self, r: Any, builder: _LatexBuilder, in_eq_arr: bool) -> None:
        r_pr, wml_r_pr, s = self._collect(r)
        if not s:
            return
        c = _classes()
        nor = scr = sty = None
        if r_pr is not None:
            for item in r_pr.nor_or_scr_or_sty or ():
                if isinstance(item, c["CTScript"]):
                    scr = _value(item)
                elif isinstance(item, c["CTStyle"]):
                    sty = _value(item)
                elif isinstance(item, c["CTOnOff"]):
                    nor = item
        if nor is not None:
            command = "text"
            if wml_r_pr is not None and _on(getattr(wml_r_pr, "b", None)):
                command = "textbf"
            elif wml_r_pr is not None and _on(getattr(wml_r_pr, "i", None)):
                command = "textit"
            builder.command(command)
            builder.raw("{" + _escape_text(s) + "}")
            return
        if scr is not None:
            scripts = {"script": "mathcal", "double-struck": "mathbb", "fraktur": "mathfrak"}
            if scr not in scripts:
                raise OmmlMathError(f"unsupported m:scr {scr}")
            builder.command(scripts[scr])
            builder.raw("{" + _escape_text(s) + "}")
            return
        if sty == "p":
            if s in FUNCTION_NAMES:
                builder.command(s)
            else:
                builder.command("mathrm")
                builder.raw("{" + _escape_text(s) + "}")
            return
        if sty == "b":
            builder.command("mathbf")
            builder.raw("{" + _escape_text(s) + "}")
            return
        if sty == "i":
            builder.command("mathit")
            builder.raw("{" + _escape_text(s) + "}")
            return
        _math_chars(s, builder, in_eq_arr)

    def _rad(self, rad: Any, builder: _LatexBuilder) -> None:
        builder.command("sqrt")
        deg_hidden = rad.rad_pr is not None and rad.rad_pr.deg_hide is not None
        if not deg_hidden and rad.deg is not None and rad.deg.content:
            builder.raw("[")
            self._elements(rad.deg.content, builder, False)
            builder.raw("]")
        self._braced(rad.e, builder)

    def _nary(self, nary: Any, builder: _LatexBuilder) -> None:
        pr = nary.nary_pr
        chr_ = _value(pr.chr) if pr is not None and pr.chr is not None else "∫"  # OMML's default
        command = REVERSE_NARY.get(chr_)
        if command is None:
            raise OmmlMathError(f"unsupported n-ary operator {chr_}")
        builder.command(command)
        sub_hidden = pr is not None and pr.sub_hide is not None
        sup_hidden = pr is not None and pr.sup_hide is not None
        if not sub_hidden and nary.sub is not None:
            builder.raw("_")
            self._braced(nary.sub, builder)
        if not sup_hidden and nary.sup is not None:
            builder.raw("^")
            self._braced(nary.sup, builder)
        if nary.e is not None:
            self._elements(nary.e.content, builder, False)

    def _delimited(self, d: Any, builder: _LatexBuilder) -> None:
        c = _classes()
        pr = d.d_pr
        beg = _value(pr.beg_chr) if pr is not None and pr.beg_chr is not None else "("
        end = _value(pr.end_chr) if pr is not None and pr.end_chr is not None else ")"
        sep = _value(pr.sep_chr) if pr is not None and pr.sep_chr is not None else None
        args = list(d.e or ())
        only = None
        if len(args) == 1:
            content = [x for x in args[0].content or () if not isinstance(x, str)]
            if len(content) == 1:
                only = content[0]
        # a lone "{" holding a single equation array is \begin{cases}
        if beg == "{" and end == "" and sep is None and isinstance(only, c["CTEqArr"]):
            self._rows("cases", only, builder)
            return
        # () or [] holding a single matrix is \begin{pmatrix} / \begin{bmatrix}
        if sep is None and isinstance(only, c["CTM"]):
            if beg == "(" and end == ")":
                self._matrix("pmatrix", only, builder)
                return
            if beg == "[" and end == "]":
                self._matrix("bmatrix", only, builder)
                return
        # plain round parens re-export bare (the importer makes m:d of those)
        if beg == "(" and end == ")" and sep is None and len(args) == 1:
            builder.raw("(")
            self._elements(args[0].content, builder, False)
            builder.raw(")")
            return
        builder.command("left")
        builder.raw(_delimiter_for(beg))
        first = True
        for e in args:
            if not first:
                builder.command("middle")
                builder.raw(_delimiter_for("∣" if sep is None else sep))
            first = False
            self._elements(e.content, builder, False)
        builder.command("right")
        builder.raw(_delimiter_for(end))

    def _rows(self, environment: str, eq_arr: Any, builder: _LatexBuilder) -> None:
        builder.raw("\\begin{" + environment + "}")
        first = True
        for row in eq_arr.e or ():
            if not first:
                builder.raw(" \\\\ ")
            first = False
            self._elements(row.content, builder, True)
        builder.raw("\\end{" + environment + "}")

    def _matrix(self, environment: str, m: Any, builder: _LatexBuilder) -> None:
        """Rows of &-separated cells: ``\\begin{matrix}a&b \\\\ c&d\\end{matrix}``."""
        builder.raw("\\begin{" + environment + "}")
        first_row = True
        for row in m.mr or ():
            if not first_row:
                builder.raw(" \\\\ ")
            first_row = False
            first_cell = True
            for cell in row.e or ():
                if not first_cell:
                    builder.raw("&")
                first_cell = False
                self._elements(cell.content, builder, False)
        builder.raw("\\end{" + environment + "}")

    def _group_chr(self, group: Any, builder: _LatexBuilder) -> None:
        pr = group.group_chr_pr
        chr_ = _value(pr.chr) if pr is not None and pr.chr is not None else "⏟"  # OMML's default: a bottom brace
        if chr_ in ("⏟", "︸"):
            builder.command("underbrace")
        elif chr_ in ("⏞", "︷"):
            builder.command("overbrace")
        else:
            raise OmmlMathError(f"unsupported group character {chr_}")
        self._braced(group.e, builder)

    def _lim_upp(self, lim_upp: Any, builder: _LatexBuilder) -> None:
        """Content above a base: ``\\xrightarrow`` for arrows, else ``\\overset``."""
        c = _classes()
        base_text = None
        if lim_upp.e is not None:
            content = [x for x in lim_upp.e.content or () if not isinstance(x, str)]
            if len(content) == 1 and isinstance(content[0], c["CTR"]):
                base_text = "".join(
                    x.value or "" for x in content[0].content or () if isinstance(x, c["CTText"])
                )
        if base_text in ("⟶", "⟵"):
            builder.command("xrightarrow" if base_text == "⟶" else "xleftarrow")
            self._braced(lim_upp.lim, builder)
            return
        builder.command("overset")
        self._braced(lim_upp.lim, builder)
        self._braced(lim_upp.e, builder)

    def _acc(self, acc: Any, builder: _LatexBuilder) -> None:
        pr = acc.acc_pr
        chr_ = _value(pr.chr) if pr is not None and pr.chr is not None else "̂"  # OMML's default: a hat
        command = REVERSE_ACCENTS.get(chr_)
        if command is None:
            raise OmmlMathError(f"unsupported accent {chr_}")
        builder.command(command)
        self._braced(acc.e, builder)

    # -- arguments ----------------------------------------------------------------

    def _braced(self, arg: Any, builder: _LatexBuilder) -> None:
        builder.raw("{")
        if arg is not None:
            self._elements(arg.content, builder, False)
        builder.raw("}")

    def _base(self, arg: Any, builder: _LatexBuilder) -> None:
        """A script's base: a single-character plain run stays bare (``U^{3}``); anything else is braced."""
        c = _classes()
        if arg is not None:
            content = [x for x in arg.content or () if not isinstance(x, str)]
            if len(content) == 1:
                u = content[0]
                if isinstance(u, (c["CTD"], c["CTGroupChr"])):
                    # scripts bind to a (...) group, or to an underbrace, whose label is its subscript: no braces
                    self._elements(arg.content, builder, False)
                    return
                if isinstance(u, c["CTR"]):
                    plain = u.r_pr is None
                    text = None
                    for x in u.content or ():
                        if isinstance(x, c["CTText"]):
                            text = x.value
                    if (
                        plain
                        and text is not None
                        and _utf16_length(text) == 1
                        and text not in REVERSE_SYMBOLS
                        and text[0] not in "{}_^%$#&\\"
                    ):
                        builder.raw(text)
                        return
        self._braced(arg, builder)


def _on(flag: Any) -> bool:
    return flag is not None and getattr(flag, "val", None) is not False


def _math_chars(s: str, builder: _LatexBuilder, in_eq_arr: bool) -> None:
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        # a combining long solidus overlay after a symbol is \not before it
        if i + 1 < n and s[i + 1] == "̸":
            builder.command("not")
            i += 1  # the combining character is consumed; c is written below
        symbol = REVERSE_SYMBOLS.get(c)
        if symbol is not None:
            builder.command(symbol)
        elif c == "′":
            builder.raw("'")
        elif c == " ":
            builder.command(",")
        elif c == " ":
            builder.command(";")
        elif c == " ":
            builder.command("quad")
        elif c in "{}_^%$#":
            builder.raw("\\" + c)
        elif c == "&":
            # inside an equation array '&' is the alignment mark
            builder.raw("&" if in_eq_arr else "\\&")
        elif c == "\\":
            builder.command("backslash")  # not round-trippable; rare
        else:
            builder.raw(c)
        i += 1


def _escape_text(s: str) -> str:
    return (
        s.replace("\\", "\\backslash ")
        .replace("{", "\\{")
        .replace("}", "\\}")
        .replace("&", "\\&")
        .replace("%", "\\%")
        .replace("$", "\\$")
        .replace("#", "\\#")
        .replace("_", "\\_")
        .replace("^", "\\^")
    )


_DELIMITERS = {
    "{": "\\{",
    "}": "\\}",
    "∣": "|",
    "‖": "\\|",
    "⟨": "\\langle ",
    "⟩": "\\rangle ",
    "⌊": "\\lfloor ",
    "⌋": "\\rfloor ",
    "⌈": "\\lceil ",
    "⌉": "\\rceil ",
}


def _delimiter_for(chr_: str | None) -> str:
    if not chr_:
        return "."
    return _DELIMITERS.get(chr_, chr_)
