"""``to_api_script``: the content-API calls that would produce what you show it.

CR-003 section 3.11, Phase I. Given a :class:`~docx4j_py.model.content.Body`, a
:class:`~docx4j_py.model.content.Paragraph`, a
:class:`~docx4j_py.model.content.Range`, a
:class:`~docx4j_py.model.content.table.Table`, a ``w:p`` / ``w:tbl`` / ``w:sdt``
element or a sequence of them, this writes **Python source against a name**
``body`` --- the verbs where they exist, and ``body.insert_xml(...)`` with the
marshalled fragment for what they cannot express, with a comment saying which
construct forced it. For an agent it is reveal codes: shown a document, it
learns the calls that would make it.

The script executes as it stands::

    >>> from docx4j_py import create_package
    >>> pkg = create_package()
    >>> pkg.body.insert_paragraph("Chapter 1", style="Heading 1")   # doctest: +ELLIPSIS
    <Paragraph ...>
    >>> script = pkg.body.to_api_script()
    >>> target = create_package()
    >>> exec(script, {"body": target.body})
    >>> target.body.text
    'Chapter 1'

**The decision is taken per block, before a line is emitted**, so a paragraph
comes out whole either way: every run's formatting, every paragraph property
and every run-level child is checked first, and one ``_Unexpressible`` sends
the whole paragraph to ``insert_xml``.

What the verbs express, and what they do not, is section 19 of the CR. The
short form: a paragraph's text, style, alignment, indents, spacing and outline
level; a run's :class:`~docx4j_py.model.content.font.Font` vocabulary,
``w:rStyle`` included; a page or line break that is a run of its own; a plain
grid of single-paragraph text cells as ``insert_table``; an inline picture; a
list item as ``start_new_list`` / ``attach_to_list`` with the level setters for
a definition that is not one of docx4j's two defaults; and a comment thread as
``insert_comment`` / ``reply`` / ``resolved``. ``w:proofErr``,
``w:lastRenderedPageBreak`` and bookmarks are dropped. Everything else ---
a hyperlink, a field, a content control, a tracked change, ``mc:AlternateContent``
--- is one ``insert_xml``.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Sequence
from typing import Any, Literal

from docx4j_py.fragments import to_xml
from docx4j_py.traversal import element_name
from docx4j_py.wml import read_run_options

__all__ = ["PLACEHOLDER_PNG", "ScriptBlock", "script_blocks", "to_api_script"]

TWIPS_PER_POINT = 20

#: A 1x1 transparent PNG, the stand-in for a document's own images in the
#: default ``pictures="placeholder"`` mode: the script still executes, and an
#: agent reading it is not handed a megabyte of base64.
PLACEHOLDER_PNG = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: Run-level markers that carry no content: dropped, since no verb writes them
#: and none of them changes what Word shows (CR-003 section 4).
_DROPPED = frozenset(
    {
        f"{W}proofErr",
        f"{W}bookmarkStart",
        f"{W}bookmarkEnd",
        f"{W}lastRenderedPageBreak",
        f"{W}commentRangeStart",
        f"{W}commentRangeEnd",
        f"{W}commentReference",
        f"{W}annotationRef",
    }
)

#: ``w:pPr`` members the paragraph verbs express; anything else falls back.
_PPR_FIELDS = frozenset({"p_style", "jc", "ind", "spacing", "outline_lvl", "num_pr"})
#: ``w:ind`` attributes the four indent members express.
_IND_FIELDS = frozenset({"left", "right", "first_line", "hanging", "start", "end"})
#: ``w:spacing`` attributes ``space_before`` / ``space_after`` / ``line_spacing`` express.
_SPACING_FIELDS = frozenset({"before", "after", "line", "line_rule"})
#: ``w:rPr`` members the ``Font`` vocabulary expresses (``builders.RUN_OPTIONS``'s markup).
_RPR_FIELDS = frozenset(
    {
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
    }
)
#: Language and proofing hints: dropped wherever they appear, as ``w:proofErr``
#: is. No verb writes them, they change nothing Word paints, and Word rewrites
#: them on the next save (CR-003 section 19.3).
_IGNORED_RPR = frozenset({"lang", "no_proof"})

#: ``w:rFonts`` attributes ``Font.name`` expresses.
_RFONTS_FIELDS = frozenset({"ascii", "h_ansi"})
#: ``w:tblPr`` members ``insert_table`` reproduces.
_TBLPR_FIELDS = frozenset({"tbl_style", "tbl_w", "tbl_look"})

#: The ``Font`` members, in the order a script sets them.
_FONT_MEMBERS = (
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
)

_DEFAULT_FORMAT = read_run_options(None)


class _Unexpressible(Exception):
    """A construct no verb writes: the block falls back to ``insert_xml``."""


# ---------------------------------------------------------------------------
# what a call comes out as
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class ScriptBlock:
    """One block of the script: the lines, and what they came from.

    :func:`script_blocks` returns these and :func:`to_api_script` joins them.
    ``kind`` is ``"paragraph"``, ``"table"`` or ``"xml"``; ``exact`` says
    whether executing the lines reproduces `element` canonically (the test of
    CR-003 section 19.4 asserts it), and ``reason`` names the construct that
    forced the fallback or made the block inexact.
    """

    kind: str
    lines: tuple[str, ...]
    address: str = ""
    exact: bool = True
    reason: str | None = None
    element: Any = None

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: everything but the element."""
        return {
            "kind": self.kind,
            "address": self.address,
            "exact": self.exact,
            "reason": self.reason,
            "lines": list(self.lines),
        }


# ---------------------------------------------------------------------------
# the literals
# ---------------------------------------------------------------------------


def _quote(text: str) -> str:
    """A Python string literal, double quoted where that needs no escaping."""
    out = repr(text)
    if out.startswith("'") and '"' not in text:
        out = '"' + out[1:-1] + '"'
    return out


def _xml_literal(xml: str) -> str:
    """A literal for a marshalled fragment: triple quoted when it can be read."""
    if '"""' not in xml and "\\" not in xml and not xml.endswith('"'):
        return f'"""{xml}"""'
    return _quote(xml)


def _num(value: float) -> str:
    """A number rounded to two places, written as an int where it is one."""
    rounded = round(float(value) + 0.0, 2)
    if rounded == int(rounded):
        return str(int(rounded))
    return str(rounded)


def _literal(value: Any) -> str:
    """A Python literal for a ``Font`` value."""
    if value is None:
        return "None"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (int, float)):
        return _num(value)
    return _quote(str(value))


def _present(obj: Any, allowed: frozenset[str]) -> list[str]:
    """The dataclass fields of `obj` that are set and not in `allowed`."""
    if obj is None:
        return []
    out: list[str] = []
    for field in dataclasses.fields(obj):
        if field.name in allowed:
            continue
        value = getattr(obj, field.name, None)
        if value is None or value == [] or field.name == "parent":
            continue
        out.append(field.name)
    return out


def _value(obj: Any) -> Any:
    """``w:val`` as a plain value, whether the model made it an enum or not."""
    val = getattr(obj, "val", None)
    return getattr(val, "value", val)


def _refuse(obj: Any, allowed: frozenset[str], what: str) -> None:
    """Raise unless every member of `obj` that is set is in `allowed`."""
    extra = _present(obj, allowed)
    if extra:
        raise _Unexpressible(f"{what} ({', '.join(sorted(extra))})")


# ---------------------------------------------------------------------------
# a paragraph's runs, as the verbs would write them
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class _Piece:
    """One run of the source, and the call that would make it."""

    kind: str  # "text" | "break" | "picture"
    text: str = ""
    break_type: str = "Line"
    fmt: dict[str, Any] = dataclasses.field(default_factory=dict)
    picture: Any = None
    note: str | None = None
    #: whether the leading-run form (``insert_paragraph``'s text) writes it
    buildable: bool = False


def _run_format(run: Any) -> dict[str, Any]:
    """The ``Font`` vocabulary of a run; raises for anything outside it."""
    r_pr = getattr(run, "r_pr", None)
    if r_pr is None:
        return dict(_DEFAULT_FORMAT)
    _refuse(r_pr, _RPR_FIELDS | _IGNORED_RPR, "run properties")
    fonts = getattr(r_pr, "r_fonts", None)
    if fonts is not None:
        _refuse(fonts, _RFONTS_FIELDS, "w:rFonts beyond ascii and hAnsi")
        if fonts.ascii is not None and fonts.h_ansi is not None and fonts.ascii != fonts.h_ansi:
            raise _Unexpressible("w:rFonts with a different ascii and hAnsi")
    read = read_run_options(r_pr)
    if read["underline"] == "Mixed":
        raise _Unexpressible(f"w:u {_value(getattr(r_pr, 'u', None))!r}")
    return read


def _text_of_run(run: Any, pictures: Any) -> _Piece:
    """The piece one ``w:r`` becomes; raises for content no verb writes."""
    fmt = _run_format(run)
    items = [item for item in (run.content or ()) if element_name(item) not in _DROPPED]
    if not items:
        return _Piece("text", "", fmt=fmt, buildable=True)

    names = [element_name(item) for item in items]
    if names == [f"{W}drawing"]:
        found = pictures.get(id(run))
        if found is None:
            raise _Unexpressible("a w:drawing that is not an inline picture")
        return _Piece("picture", fmt=fmt, picture=found)

    if names == [f"{W}br"]:
        kind = _value_of_type(items[0])
        if kind not in (None, "textWrapping", "page"):
            raise _Unexpressible(f"w:br w:type={kind!r}")
        note = None
        if getattr(run, "r_pr", None) is not None:
            note = "insert_break writes no run properties"
        return _Piece("break", break_type="Page" if kind == "page" else "Line", fmt=fmt, note=note)

    text = ""
    buildable = True
    previous = ""
    for item, name in zip(items, names, strict=True):
        if name == f"{W}t":
            if previous == f"{W}t":
                # two adjacent w:t in one run: the r() builder writes one
                buildable = False
            text += str(getattr(item, "value", "") or "")
        elif name == f"{W}tab":
            text += "\t"
        elif name == f"{W}br":
            kind = _value_of_type(item)
            if kind not in (None, "textWrapping"):
                raise _Unexpressible(f"w:br w:type={kind!r} beside text in a run")
            text += "\n"
        else:
            raise _Unexpressible(f"{_short(name)} in a run")
        previous = name
    return _Piece("text", text=text, fmt=fmt, buildable=buildable)


def _value_of_type(item: Any) -> Any:
    """A ``w:br``'s ``w:type``, as a plain string."""
    kind = getattr(item, "type_value", None)
    return getattr(kind, "value", kind)


def _short(name: str) -> str:
    """``{ns}local`` as ``w:local`` where the namespace is WordprocessingML."""
    if name.startswith(W):
        return f"w:{name[len(W) :]}"
    return name.rpartition("}")[2] or name


def _pieces_of(element: Any, pictures: dict[int, Any]) -> list[_Piece]:
    """The pieces of a ``w:p``; raises for anything a verb cannot write."""
    out: list[_Piece] = []
    for item in element.content or ():
        name = element_name(item)
        if name in _DROPPED:
            continue
        if name != f"{W}r":
            raise _Unexpressible(f"{_short(name)} in the paragraph")
        piece = _text_of_run(item, pictures)
        if piece.kind == "text" and piece.text == "":
            # a run with nothing left in it once the markers are dropped --- a
            # comment reference, a bookmark's run --- writes nothing
            continue
        out.append(piece)
    return out


def _font_diff(previous: dict[str, Any], current: dict[str, Any]) -> list[tuple[str, Any]]:
    """The ``Font`` assignments that turn `previous` into `current`."""
    return [(name, current[name]) for name in _FONT_MEMBERS if previous[name] != current[name]]


# ---------------------------------------------------------------------------
# the emitter
# ---------------------------------------------------------------------------


class _Emitter:
    """One per call, so that the variable names are stable and readable."""

    def __init__(
        self,
        *,
        variable: str,
        location: str,
        pictures: str,
        addresses: bool,
    ) -> None:
        self.variable = variable
        self.location = location
        self.pictures = pictures
        self.addresses = addresses
        self.blocks: list[ScriptBlock] = []
        self.needs_author = False
        self.needs_placeholder = False
        self._counts: dict[str, int] = {}
        self._lists: dict[str, str] = {}
        self._author: tuple[str, str | None, str | None] | None = None

    # -- names -------------------------------------------------------------

    def name(self, prefix: str) -> str:
        """``p1``, ``r2``, ``t1``, ``list1``, ``c3``: the next name of a kind."""
        self._counts[prefix] = self._counts.get(prefix, 0) + 1
        return f"{prefix}{self._counts[prefix]}"

    def at(self) -> str:
        """``, location="Start"`` when the caller asked for one; nothing otherwise."""
        return "" if self.location == "End" else f", location={_quote(self.location)}"

    # -- the dispatch ------------------------------------------------------

    def element(self, element: Any, body: Any) -> None:
        """One block-level element: a paragraph, a table, or ``insert_xml``."""
        name = element_name(element)
        if name == f"{W}p":
            self.paragraph(element, body)
        elif name == f"{W}tbl":
            self.table(element, body)
        else:
            self.fallback(element, body, _short(name))

    def paragraph(self, element: Any, body: Any) -> None:
        """A ``w:p``, whole through the verbs or whole through ``insert_xml``."""
        try:
            lines, exact, reason = self._paragraph_lines(element, body)
        except _Unexpressible as error:
            self.fallback(element, body, str(error))
            return
        self.blocks.append(
            ScriptBlock(
                kind="paragraph",
                lines=tuple(lines),
                address=_address(body, element),
                exact=exact,
                reason=reason,
                element=element,
            )
        )

    # -- a paragraph -------------------------------------------------------

    def _paragraph_lines(self, element: Any, body: Any) -> tuple[list[str], bool, str | None]:
        """Every line a paragraph needs, decided before one is emitted."""
        view = _view_of(body, element)
        pictures = {id(picture.run): picture for picture in (view.inline_pictures if view else ())}
        properties = self._properties(element, view)
        pieces = _pieces_of(element, pictures)

        notes: list[str] = []
        exact = True

        first = pieces[0] if pieces else None
        leading = (
            first is not None
            and first.kind == "text"
            and first.buildable
            and first.fmt == _DEFAULT_FORMAT
        )
        rest = pieces[1:] if leading else pieces

        # every font change is checked before a line is emitted
        diffs: list[list[tuple[str, Any]]] = []
        previous = dict(_DEFAULT_FORMAT)
        for piece in rest:
            if piece.kind == "break":
                diffs.append([])
                previous = dict(_DEFAULT_FORMAT)
                continue
            diffs.append(_font_diff(previous, piece.fmt))
            previous = piece.fmt

        text = first.text if leading and first is not None else ""
        lines: list[str] = []
        if not properties and not rest and leading:
            # the whole paragraph is one plain run: one line, no variable
            lines.append(f"{self.variable}.insert_paragraph({_quote(text)}{self.at()})")
            return lines, exact, None

        name = self.name("p")
        lines.append(f"{name} = {self.variable}.insert_paragraph({_quote(text)}{self.at()})")
        for member, value in properties:
            lines.append(f"{name}.{member} = {value}")
        lines.extend(self._list_lines(name, view, notes))

        for piece, diff in zip(rest, diffs, strict=True):
            if piece.kind == "break":
                lines.append(f"{name}.insert_break({_quote(piece.break_type)})")
                if piece.note:
                    notes.append(piece.note)
                    exact = False
                continue
            if piece.kind == "picture":
                lines.extend(self._picture_lines(name, piece.picture))
                exact = False
                notes.append("an inline picture")
                continue
            if "\t" in piece.text or "\n" in piece.text:
                # insert_text writes the characters; the r() builder behind
                # insert_paragraph writes w:tab and w:br (section 19.3)
                notes.append("a w:tab or w:br outside the leading run")
                exact = False
            if not diff:
                lines.append(f"{name}.insert_text({_quote(piece.text)})")
                continue
            span = self.name("r")
            lines.append(f"{span} = {name}.insert_text({_quote(piece.text)})")
            for member, value in diff:
                lines.append(f"{span}.font.{member} = {_literal(value)}")

        comment_lines, commented = self._comment_lines(name, view)
        if commented:
            exact = False
            notes.append("the comment markers, which insert_comment places itself")
        lines.extend(comment_lines)

        if notes:
            lines.insert(0, f"# {'; '.join(dict.fromkeys(notes))}: not reproduced exactly")
        return lines, exact, "; ".join(dict.fromkeys(notes)) or None

    def _properties(self, element: Any, view: Any) -> list[tuple[str, str]]:
        """The paragraph members that carry ``w:pPr``; raises for anything else."""
        p_pr = getattr(element, "p_pr", None)
        if p_pr is None:
            return []
        _refuse(p_pr, _PPR_FIELDS | {"r_pr"}, "paragraph properties")
        # the paragraph mark's own run properties: no verb writes them
        # (``Paragraph.font`` is over the runs, CR-003 section 4), so a mark
        # that says anything beyond a language or proofing hint falls back
        _refuse(getattr(p_pr, "r_pr", None), _IGNORED_RPR, "the paragraph mark's w:rPr")
        out: list[tuple[str, str]] = []

        style = _value(getattr(p_pr, "p_style", None))
        if style is not None and style != "Normal":
            out.append(_style_assignment(str(style)))

        jc = _value(getattr(p_pr, "jc", None))
        if jc is not None:
            alignment = _ALIGNMENT.get(str(jc))
            if alignment is None:
                raise _Unexpressible(f"w:jc {jc}")
            out.append(("alignment", _quote(alignment)))

        ind = getattr(p_pr, "ind", None)
        if ind is not None:
            _refuse(ind, _IND_FIELDS, "w:ind")
            left = ind.left if ind.left is not None else ind.start
            right = ind.right if ind.right is not None else ind.end
            if left is not None:
                out.append(("left_indent", _num(float(left) / TWIPS_PER_POINT)))
            if right is not None:
                out.append(("right_indent", _num(float(right) / TWIPS_PER_POINT)))
            if ind.hanging is not None:
                out.append(("first_line_indent", _num(-float(ind.hanging) / TWIPS_PER_POINT)))
            elif ind.first_line is not None:
                out.append(("first_line_indent", _num(float(ind.first_line) / TWIPS_PER_POINT)))

        spacing = getattr(p_pr, "spacing", None)
        if spacing is not None:
            _refuse(spacing, _SPACING_FIELDS, "w:spacing")
            if spacing.before is not None:
                out.append(("space_before", _num(float(spacing.before) / TWIPS_PER_POINT)))
            if spacing.after is not None:
                out.append(("space_after", _num(float(spacing.after) / TWIPS_PER_POINT)))
            if spacing.line is not None:
                rule = getattr(spacing, "line_rule", None)
                rule = getattr(rule, "value", rule)
                if rule != "exact":
                    # line_spacing writes w:lineRule="exact", so only an exact
                    # rule round-trips (docx4j-core-ts CR-002 section 10)
                    raise _Unexpressible(f'w:spacing w:lineRule="{rule or "auto"}"')
                out.append(("line_spacing", _num(float(spacing.line) / TWIPS_PER_POINT)))
            if spacing.line is None and getattr(spacing, "line_rule", None) is not None:
                raise _Unexpressible("w:spacing w:lineRule without w:line")

        outline = _value(getattr(p_pr, "outline_lvl", None))
        if outline is not None:
            out.append(("outline_level", str(int(outline) + 1)))

        if getattr(p_pr, "num_pr", None) is not None and view is None:
            raise _Unexpressible("w:numPr outside a package")
        return out

    # -- lists (CR-003 section 18.7) ---------------------------------------

    def _list_lines(self, name: str, view: Any, notes: list[str]) -> list[str]:
        """``start_new_list`` for the first item of a ``w:numId``, then ``attach_to_list``.

        The ids are **the script's own**, held in a variable per list, because
        a literal would name the source document's ``w:numId`` and the target's
        allocator knows nothing of it (CR-003 section 18.7).
        """
        if view is None:
            return []
        p_pr = getattr(view.element, "p_pr", None)
        if getattr(p_pr, "num_pr", None) is None:
            return []
        item = view.list_item
        listing = view.list
        if item is None or listing is None:
            raise _Unexpressible("a w:numPr the paragraph is not numbered by")

        lines: list[str] = []
        key = str(listing.id)
        variable = self._lists.get(key)
        if variable is None:
            variable = self.name("list")
            self._lists[key] = variable
            kind = "Bullet" if listing.level_types[0] == "Bullet" else "Number"
            lines.append(f"{variable} = {name}.start_new_list(kind={_quote(kind)})")
            lines.extend(_definition_lines(variable, listing, kind, notes))
            if item.level:
                lines.append(f"{name}.attach_to_list({variable}.id, {item.level})")
        else:
            lines.append(f"{name}.attach_to_list({variable}.id, {item.level})")
        if _value(getattr(p_pr, "p_style", None)) is None:
            # start_new_list and attach_to_list apply Word's List Paragraph to a
            # paragraph that names no style; the source names none (section 19.3)
            lines.append(f'{name}.style_id = ""')
        return lines

    # -- pictures ----------------------------------------------------------

    def _picture_lines(self, name: str, picture: Any) -> list[str]:
        """``insert_inline_picture_from_base64``, with the bytes or a placeholder."""
        lines: list[str] = []
        arguments = [
            f"width={_num(picture.width)}",
            f"height={_num(picture.height)}",
        ]
        if picture.alt_text_description:
            arguments.append(f"alt_text_description={_quote(picture.alt_text_description)}")
        if self.pictures == "inline":
            data = _quote(picture.get_base64())
        else:
            self.needs_placeholder = True
            data = "PLACEHOLDER_PNG"
            part = picture.image_part
            where = str(part.part_name) if part is not None else "an image part"
            size = len(picture.get_bytes()) if part is not None else 0
            lines.append(f"# {where}, {size} bytes")
        lines.append(
            f"{name}.insert_inline_picture_from_base64({data}, {', '.join(arguments)})"
        )
        return lines

    # -- comments ----------------------------------------------------------

    def _comment_lines(self, name: str, view: Any) -> tuple[list[str], bool]:
        """``insert_comment`` on the anchor, then ``reply`` and ``resolved``."""
        if view is None:
            return [], False
        try:
            comments = [c for c in view.get_comments() if c.parent is None]
        except Exception:  # noqa: BLE001 - a document without comment parts has none
            return [], False
        if not comments:
            return [], False
        lines: list[str] = []
        for comment in comments:
            anchor = self._anchor(name, comment, view, lines)
            lines.extend(self._author_lines(comment))
            handle = self.name("c")
            lines.append(f"{handle} = {anchor}.insert_comment({_quote(comment.content)})")
            for reply in comment.replies:
                lines.extend(self._author_lines(reply))
                lines.append(f"{handle}.reply({_quote(reply.content)})")
            if comment.resolved:
                lines.append(f"{handle}.resolved = True")
        return lines, True

    def _anchor(self, name: str, comment: Any, view: Any, lines: list[str]) -> str:
        """The expression for the range a comment is on."""
        spans = [span for span in comment.get_range() if span.paragraph == view]
        if not spans:
            lines.append(f"# comment {comment.id}: the anchor is not a span in this paragraph")
            return f"{name}.get_range()"
        span = spans[0]
        text = span.text
        if not text or (span.start == 0 and span.end == len(view.text)):
            return f"{name}.get_range()"
        if view.text.count(text) != 1:
            lines.append(f"# comment {comment.id}: {text!r} is not unique in the paragraph")
            return f"{name}.get_range()"
        return f"{name}.search({_quote(text)})[0]"

    def _author_lines(self, comment: Any) -> list[str]:
        """``body.package.author = Author(...)`` when the author changes."""
        who = (comment.author_name, comment.initials or None, comment.author_email or None)
        if who == self._author:
            return []
        self._author = who
        self.needs_author = True
        arguments = [_quote(who[0])]
        if who[1]:
            arguments.append(f"initials={_quote(who[1])}")
        if who[2]:
            arguments.append(f"email={_quote(who[2])}")
        return [f"{self.variable}.package.author = Author({', '.join(arguments)})"]

    # -- tables ------------------------------------------------------------

    def table(self, element: Any, body: Any) -> None:
        """A ``w:tbl`` through ``insert_table`` when it is a plain grid of text."""
        try:
            grid = _table_grid(element)
        except _Unexpressible as error:
            self.fallback(element, body, str(error))
            return
        name = self.name("t")
        values = ", ".join(
            "[" + ", ".join(_quote(cell) for cell in row) + "]" for row in grid["values"]
        )
        arguments = [str(grid["rows"]), str(grid["columns"]), f"values=[{values}]"]
        lines = [f"{name} = {self.variable}.insert_table({', '.join(arguments)}{self.at()})"]
        if grid["style"] is not None:
            member, value = _style_assignment(grid["style"])
            lines.append(f"{name}.{member} = {value}")
        if grid["header_rows"]:
            lines.append(f"{name}.header_row_count = {grid['header_rows']}")
        self.blocks.append(
            ScriptBlock(
                kind="table",
                lines=tuple(lines),
                address=_address(body, element),
                exact=False,
                reason="insert_table sizes its own w:tblGrid",
                element=element,
            )
        )

    # -- the fallback ------------------------------------------------------

    def fallback(self, element: Any, body: Any, reason: str) -> None:
        """What no verb expresses: the marshalled fragment, which ``insert_xml`` takes back."""
        try:
            xml = to_xml(element)
        except Exception as error:  # noqa: BLE001 - a wildcard the serialiser cannot name
            self.blocks.append(
                ScriptBlock(
                    kind="xml",
                    lines=(f"# {reason}: cannot be marshalled ({error})",),
                    address=_address(body, element),
                    exact=False,
                    reason=reason,
                    element=element,
                )
            )
            return
        self.blocks.append(
            ScriptBlock(
                kind="xml",
                lines=(
                    f"# {reason}: as XML",
                    f"{self.variable}.insert_xml({_xml_literal(xml)}{self.at()})",
                ),
                address=_address(body, element),
                exact=True,
                reason=reason,
                element=element,
            )
        )


# ---------------------------------------------------------------------------
# the helpers the emitter reads
# ---------------------------------------------------------------------------

#: ``w:jc`` -> the ``Word.Alignment`` value ``Paragraph.alignment`` writes back.
_ALIGNMENT: dict[str, str] = {
    "left": "Left",
    "start": "Left",
    "center": "Centered",
    "right": "Right",
    "end": "Right",
    "both": "Justified",
}


def _style_assignment(style_id: str) -> tuple[str, str]:
    """``style_built_in`` for a built-in style, ``style_id`` for anything else.

    Office JS's ``style`` takes the *display name*, which only resolves against
    a document that defines the style; a script run against a fresh body would
    raise on a name the target has never heard of, so a style that is not
    built in is written by its id (CR-003 section 19.3).
    """
    from docx4j_py.model.content.styles import built_in_of

    built_in = built_in_of(style_id)
    if built_in == "Other":
        return ("style_id", _quote(style_id))
    return ("style_built_in", _quote(built_in))


def _view_of(body: Any, element: Any) -> Any:
    """The ``Paragraph`` view of a ``w:p``, or None when there is no body."""
    if body is None:
        return None
    try:
        return body.paragraph_for(element)
    except Exception:  # noqa: BLE001 - an element outside the body's tree has no view
        return None


def _address(body: Any, element: Any) -> str:
    """The **ordinal** address of a block (``body/3``), or ``""``.

    The ordinal rather than the ``w14:paraId`` :meth:`Body.address_of` prefers:
    a paraId is an id of the source document, and CR-003 section 3.11's
    determinism rule keeps those out of the script (section 19.3).
    """
    if body is None:
        return ""
    from docx4j_py.model.content.addresses import ordinal_of

    try:
        return ordinal_of(body, element) or ""
    except Exception:  # noqa: BLE001 - an element outside the body's tree has no address
        return ""


def _table_grid(element: Any) -> dict[str, Any]:
    """A table as ``insert_table`` would build it; raises when it is more."""
    _refuse(getattr(element, "tbl_pr", None), _TBLPR_FIELDS, "table properties")
    values: list[list[str]] = []
    header_rows = 0
    counting = True
    for row in element.content or ():
        name = element_name(row)
        if name != f"{W}tr":
            raise _Unexpressible(f"{_short(name)} in the table")
        tr_pr = getattr(row, "tr_pr", None)
        header = _has_header(tr_pr)
        _refuse(tr_pr, frozenset({"content"}), "row properties")
        if tr_pr is not None:
            _refuse_row(tr_pr)
        if header and counting:
            header_rows += 1
        else:
            counting = False
        cells: list[str] = []
        for cell in row.content or ():
            if element_name(cell) != f"{W}tc":
                raise _Unexpressible(f"{_short(element_name(cell))} in the row")
            cells.append(_cell_text(cell))
        values.append(cells)
    if not values:
        raise _Unexpressible("a table with no rows")
    columns = len(values[0])
    if columns == 0 or any(len(row) != columns for row in values):
        raise _Unexpressible("a table whose rows have different cell counts")
    style = _value(getattr(getattr(element, "tbl_pr", None), "tbl_style", None))
    return {
        "rows": len(values),
        "columns": columns,
        "values": values,
        "style": None if style is None else str(style),
        "header_rows": header_rows,
    }


def _has_header(tr_pr: Any) -> bool:
    """Whether a ``w:trPr`` carries ``w:tblHeader``."""
    for item in getattr(tr_pr, "content", None) or ():
        if element_name(item) == f"{W}tblHeader":
            return True
    return False


def _refuse_row(tr_pr: Any) -> None:
    """A row's properties may carry ``w:tblHeader`` and nothing else."""
    for item in getattr(tr_pr, "content", None) or ():
        name = element_name(item)
        if name != f"{W}tblHeader":
            raise _Unexpressible(f"{_short(name)} in w:trPr")


def _cell_text(cell: Any) -> str:
    """The text of a plain single-paragraph cell; raises when it is more."""
    blocks = [item for item in (cell.content or ()) if element_name(item) != f"{W}tcPr"]
    if len(blocks) != 1 or element_name(blocks[0]) != f"{W}p":
        raise _Unexpressible("a cell that is not one plain paragraph")
    paragraph = blocks[0]
    if getattr(paragraph, "p_pr", None) is not None:
        raise _Unexpressible("a cell paragraph with properties")
    pieces = _pieces_of(paragraph, {})
    if any(piece.kind != "text" or piece.fmt != _DEFAULT_FORMAT for piece in pieces):
        raise _Unexpressible("a cell with formatted text")
    return "".join(piece.text for piece in pieces)


# ---------------------------------------------------------------------------
# a list's definition (CR-003 section 18.7)
# ---------------------------------------------------------------------------

#: ``w:numFmt`` -> the ``Word.ListNumbering`` value ``set_level_numbering`` takes.
_NUMBERING_BACK: dict[str, str] = {
    "none": "None",
    "decimal": "Arabic",
    "upperRoman": "UpperRoman",
    "lowerRoman": "LowerRoman",
    "upperLetter": "UpperLetter",
    "lowerLetter": "LowerLetter",
}

_DEFAULT_LEVELS: dict[str, dict[int, tuple]] = {}


def _facts(level: Any) -> tuple:
    """What this generator compares of a level: what the three setters write, and the rest."""
    if level is None:
        return ()
    ind = level.ind
    return (
        level.num_fmt,
        level.level_text,
        bool(level.is_bullet),
        level.font,
        None if ind is None else ind.left,
        None if ind is None else ind.hanging,
        None if ind is None else ind.first_line,
        None if ind is None else ind.right,
        level.start_value,
        level.lvl_restart,
        bool(level.is_lgl),
        level.p_style,
    )


def _default_levels(kind: str) -> dict[int, tuple]:
    """The levels ``start_new_list(kind=...)`` itself produces, read once per process.

    Read from a throwaway package rather than from the resource, so that what
    the generator compares against is exactly what the verb writes.
    """
    found = _DEFAULT_LEVELS.get(kind)
    if found is not None:
        return found
    from docx4j_py import create_package

    package = create_package()
    package.id_seed = 0
    paragraph = package.body.insert_paragraph("x")
    listing = paragraph.start_new_list(kind=kind)
    definition = listing.definition
    found = {
        ilvl: _facts(None if definition is None else definition.level(ilvl)) for ilvl in range(9)
    }
    _DEFAULT_LEVELS[kind] = found
    return found


def _definition_lines(variable: str, listing: Any, kind: str, notes: list[str]) -> list[str]:
    """The level setters that turn the default definition into this list's.

    ``set_level_numbering``, ``set_level_bullet`` and ``set_level_indents`` are
    the three Office JS gives; what none of them says --- ``w:lvlJc``,
    ``w:isLgl``, ``w:suff``, ``w:start``, ``w:lvlRestart``, a picture bullet ---
    is named in a comment and left to the target document's own defaults
    (CR-003 section 19.3: there is no ``insert_xml`` for a numbering part).
    """
    definition = listing.definition
    if definition is None:
        return []
    defaults = _default_levels(kind)
    lines: list[str] = []
    for ilvl in range(9):
        level = definition.level(ilvl)
        if level is None:
            continue
        facts = _facts(level)
        if facts == defaults.get(ilvl):
            continue
        (fmt, text, bullet, font, left, hanging, first_line, right, start, restart, lgl, style) = (
            facts
        )
        was = defaults.get(ilvl) or ((None,) * 12)
        if (fmt, text, bullet, font) != (was[0], was[1], was[2], was[3]):
            if bullet:
                if text is not None and len(text) == 1:
                    arguments = [str(ilvl), '"Custom"', f"char_code={ord(text)}"]
                    if font:
                        arguments.append(f"font_name={_quote(font)}")
                    lines.append(f"{variable}.set_level_bullet({', '.join(arguments)})")
                else:
                    notes.append(f"level {ilvl}'s bullet {text!r}")
            else:
                word = _NUMBERING_BACK.get(str(fmt or ""))
                if word is None:
                    notes.append(f"level {ilvl}'s w:numFmt {fmt!r}")
                else:
                    arguments = [str(ilvl), _quote(word)]
                    if text is not None:
                        arguments.append(_quote(text))
                    lines.append(f"{variable}.set_level_numbering({', '.join(arguments)})")
                    if font:
                        notes.append(f"level {ilvl}'s w:rFonts")
        if (left, hanging) != (was[4], was[5]):
            lines.append(
                f"{variable}.set_level_indents({ilvl}, "
                f"{_num((left or 0) / TWIPS_PER_POINT)}, "
                f"{_num((hanging or 0) / TWIPS_PER_POINT)})"
            )
        beyond = []
        if (first_line, right) != (was[6], was[7]):
            beyond.append("w:ind/@w:firstLine or @w:right")
        if start != was[8]:
            beyond.append(f"w:start {start + 1}")
        if restart != was[9]:
            beyond.append("w:lvlRestart")
        if lgl != was[10]:
            beyond.append("w:isLgl")
        if style != was[11]:
            beyond.append("w:pStyle")
        if beyond:
            notes.append(f"level {ilvl}'s {', '.join(beyond)}")
    return lines


# ---------------------------------------------------------------------------
# the public calls
# ---------------------------------------------------------------------------


def _targets(target: Any) -> tuple[list[Any], Any, Any]:
    """`target` as (the block elements, the body, the range it was)."""
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range
    from docx4j_py.model.content.table import Table

    if isinstance(target, Body):
        return list(target.content), target, None
    if isinstance(target, Range):
        return [], target.paragraph.parent_body, target
    if isinstance(target, (Paragraph, Table)):
        return [target.element], target.parent_body, None
    if isinstance(target, (list, tuple)) or (
        isinstance(target, Sequence) and not isinstance(target, (str, bytes))
    ):
        elements: list[Any] = []
        body: Any = None
        for item in target:
            found, its_body, span = _targets(item)
            if span is not None:
                raise TypeError("to_api_script takes one Range, not a sequence of them")
            elements.extend(found)
            body = body or its_body
        return elements, body, None
    if element_name(target):
        return [target], _body_of(target), None
    raise TypeError(
        "to_api_script takes a Body, a Paragraph, a Range, a Table, "
        f"a w:p / w:tbl / w:sdt element or a sequence of them, not {type(target).__name__}"
    )


def _body_of(element: Any) -> Any:
    """The ``Body`` a bare element belongs to, walking its parents; None when loose."""
    from docx4j_py.model.content.body import body_of

    current = getattr(element, "parent", None)
    for _ in range(64):
        if current is None:
            return None
        if element_name(current) == f"{W}body":
            part = getattr(current, "parent", None)
            try:
                return body_of(part) if part is not None else None
            except Exception:  # noqa: BLE001 - a w:body outside a part has no view
                return None
        current = getattr(current, "parent", None)
    return None  # pragma: no cover - a tree deeper than the walk


def script_blocks(
    target: Any,
    *,
    variable: str = "body",
    location: str = "End",
    pictures: Literal["inline", "placeholder"] = "placeholder",
    limit: int | None = None,
    addresses: bool = False,
) -> list[ScriptBlock]:
    """The blocks :func:`to_api_script` joins, one per paragraph or table.

    Exposed so that a test (and CR-003 section 19.4's measurement) can ask
    which blocks the verbs expressed and which fell back, and why.
    """
    elements, body, span = _targets(target)
    emitter = _Emitter(
        variable=variable, location=location, pictures=pictures, addresses=addresses
    )
    if span is not None:
        emitter.blocks.append(_range_block(emitter, span))
        return emitter.blocks
    shown = elements if limit is None else elements[: max(0, limit)]
    for element in shown:
        emitter.element(element, body)
    if limit is not None and len(elements) > len(shown):
        more = len(elements) - len(shown)
        emitter.blocks.append(
            ScriptBlock(
                kind="note",
                lines=(
                    (
                        f"# ... {more} more blocks; call to_api_script(body, limit=None) "
                        "or pass a Paragraph/Table for one"
                    ),
                ),
                exact=False,
                reason="limit",
            )
        )
    return emitter.blocks


def _range_block(emitter: _Emitter, span: Any) -> ScriptBlock:
    """A range: its runs only, against `variable` as a ``Paragraph`` name."""
    from docx4j_py.fragments import wml

    element = wml(span.get_xml())
    try:
        pieces = _pieces_of(element, {})
    except _Unexpressible as error:
        return ScriptBlock(
            kind="xml",
            lines=(
                f"# {error}: as XML",
                f"{emitter.variable}.insert_xml({_xml_literal(to_xml(element))})",
            ),
            address=span.paragraph.address,
            exact=False,
            reason=str(error),
            element=element,
        )
    lines: list[str] = []
    previous = dict(_DEFAULT_FORMAT)
    exact = True
    for piece in pieces:
        if piece.kind != "text":
            lines.append(f"# {piece.kind}: not reproduced in a range")
            exact = False
            continue
        diff = _font_diff(previous, piece.fmt)
        previous = piece.fmt
        if not diff:
            lines.append(f"{emitter.variable}.insert_text({_quote(piece.text)})")
            continue
        name = emitter.name("r")
        lines.append(f"{name} = {emitter.variable}.insert_text({_quote(piece.text)})")
        for member, value in diff:
            lines.append(f"{name}.font.{member} = {_literal(value)}")
    return ScriptBlock(
        kind="paragraph",
        lines=tuple(lines),
        address=span.paragraph.address,
        exact=exact,
        reason=None if exact else "a range emits its runs only",
        element=element,
    )


def to_api_script(
    target: Any,
    *,
    variable: str = "body",
    location: str = "End",
    pictures: Literal["inline", "placeholder"] = "placeholder",
    limit: int | None = None,
    addresses: bool = False,
) -> str:
    """The content-API calls that reproduce what you show it, as Python source.

    The result is a sequence of statements over a name `variable` (``body`` by
    default) and **nothing else in scope**: ``exec(script, {"body": body})``
    runs it. One block per paragraph or table, a blank line between blocks, a
    variable per element the script refers back to (``p1``, ``r1``, ``t1``,
    ``list1``, ``c1``, allocated in document order), and ``import`` lines only
    where they are needed.

    Args:
        target: a :class:`~docx4j_py.model.content.Body`, a
            :class:`~docx4j_py.model.content.Paragraph`, a
            :class:`~docx4j_py.model.content.Range`, a
            :class:`~docx4j_py.model.content.table.Table`, a ``w:p`` / ``w:tbl``
            / ``w:sdt`` element, or a sequence of them.
        variable: the name the script inserts into. **For a ``Range`` it names
            a** :class:`~docx4j_py.model.content.Paragraph`, because a range
            emits its runs only, without its paragraph's properties.
        location: ``"End"`` (the default) or ``"Start"``, passed to every
            block-level insert.
        pictures: ``"placeholder"`` (the default) puts a 1x1 PNG constant in
            place of every image, with a comment giving the part name and its
            size, so that a script stays readable; ``"inline"`` writes the real
            base64.
        limit: at most this many blocks, then a comment saying how many were
            left and how to get them.
        addresses: a comment line naming each block's source address
            (``# body/3``), the addresses of CR-003 section 3.4.

    Returns:
        The script. The same document gives the same bytes: nothing in it comes
        from the wall clock or from an id of the source document.
    """
    blocks = script_blocks(
        target,
        variable=variable,
        location=location,
        pictures=pictures,
        limit=limit,
        addresses=addresses,
    )
    chunks: list[str] = []
    for block in blocks:
        lines = list(block.lines)
        if not lines:
            continue
        if addresses and block.address:
            lines.insert(0, f"# {block.address}")
        chunks.append("\n".join(lines))

    preamble: list[str] = []
    text = "\n".join(chunks)
    if "Author(" in text:
        preamble.append("from docx4j_py.model.content import Author")
    if "PLACEHOLDER_PNG" in text:
        preamble.append(
            "PLACEHOLDER_PNG = "
            f"{_quote(PLACEHOLDER_PNG)}  # a 1x1 PNG in place of the document's own"
        )
    if preamble:
        chunks.insert(0, "\n".join(preamble))
    return "\n\n".join(chunks)
