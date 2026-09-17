"""Markdown *out*: a document read as CommonMark plus GFM, with addresses.

CR-003 section 3.5, Phase K. The behavioural oracle is Java docx4j's
``docx4j-markdown`` module (``WmlToMarkdown``), whose semantics this follows:
headings from the outline level, bold / italic / strike / code from the run's
properties, links through the part's relationships, lists from ``w:numPr`` with
bullet-or-number decided by the numbering part's ``w:numFmt``, tables as GFM
pipe tables, images as ``![alt](media/imageN.png)``, footnotes as GFM
footnotes. What is *not* followed is Java's idioms: there is no commonmark AST
here, the renderer writes strings, and there is no ``PropertyResolver`` yet
(CR-002 Phase B), so formatting is read **direct** as CR-003 section 6 allows.

Two views, as docx4j-mcp's ``docx_to_markdown`` asks for:

``view="accepted"`` (the default)
    the document as if every tracked change were accepted: ``w:ins`` and
    ``w:moveTo`` contribute, ``w:del`` and ``w:moveFrom`` do not.
``view="markup"``
    CriticMarkup: an insertion is ``{++...++}``, a deletion ``{--...--}`` and a
    comment ``{>>...<<}`` after the run that carries its reference.

**Addresses.** ``to_markdown(addresses=True)`` puts each block's address ---
:attr:`~docx4j_py.model.content.paragraph.Paragraph.address`, which is the
``w14:paraId`` when the paragraph has one and the ordinal otherwise --- in an
HTML comment on **its own line immediately before the block**, indented as the
block is::

    <!-- w14:5A2B1C3D -->
    # Quarterly Report

    <!-- body/1 -->
    The quick brown fox.

:data:`ADDRESS_COMMENT` is the regex that recovers them, and
:meth:`~docx4j_py.model.content.body.Body.element_at` accepts what comes out of
it. The addressed output is for **reading and addressing**, not for feeding
back through ``insert_markdown``: a comment line between two list items ends
the list, as CommonMark says it must. The coarse "markdown in, markdown out"
workflow does not ask for addresses; the fine one does (section 3.5's last
paragraph, and the README).

Nothing is unmarshalled by a read: the numbering part, the footnotes part and
the comments part are read as **bytes** with lxml, the way
:mod:`docx4j_py.model.content.describe` does, so a document that is only
exported is still written back byte for byte.
"""

from __future__ import annotations

import dataclasses
import re
from enum import StrEnum
from typing import Any, Literal

from docx4j_py.model.content.reports import TextExcerpt, heading_level_of
from docx4j_py.model.content.text_model import (
    RUN_HOLDER_NAMES,
    block_children_of,
    item_text,
)
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of

__all__ = [
    "ADDRESS_COMMENT",
    "CODE_STYLE_IDS",
    "MONOSPACE_FONTS",
    "MarkdownView",
    "MarkdownViewValue",
    "address_comment",
    "body_markdown",
    "markdown_budget_of",
    "paragraph_markdown",
    "table_markdown",
]


# ---------------------------------------------------------------------------
# the values (CR-003 section 3.1: a Literal, with a StrEnum beside it)
# ---------------------------------------------------------------------------

#: Which view of a tracked document ``to_markdown`` renders.
MarkdownViewValue = Literal["accepted", "markup"]


class MarkdownView(StrEnum):
    """``docx_to_markdown``'s ``tracked_changes`` switch (docx4j-mcp)."""

    ACCEPTED = "accepted"
    """Every tracked change accepted: insertions in, deletions out."""
    MARKUP = "markup"
    """CriticMarkup: ``{++inserted++}``, ``{--deleted--}``, ``{>>comment<<}``."""


#: The exact form of an address comment, and how to recover one. The whole
#: line is the comment; ``match.group("address")`` is what ``element_at`` takes.
ADDRESS_COMMENT = re.compile(r"^[ \t]*<!--\s(?P<address>\S+)\s-->$", re.MULTILINE)

#: Character style ids that mean "this run is code" (docx4j ``ImportStyles``).
CODE_STYLE_IDS: frozenset[str] = frozenset({"CodeChar", "Code", "HTMLCode"})

#: Fonts that mean "this run is code" when nothing else says so. Java
#: ``WmlToMarkdown.MONO_FONTS``, lower-cased.
MONOSPACE_FONTS: frozenset[str] = frozenset(
    {
        "consolas",
        "courier new",
        "courier",
        "lucida console",
        "monaco",
        "menlo",
        "dejavu sans mono",
        "source code pro",
        "fira code",
    }
)

#: The paragraph style ids Java's importer and exporter agree on.
SOURCE_CODE_STYLE = "SourceCode"
QUOTE_STYLE_IDS: frozenset[str] = frozenset({"Quote", "IntenseQuote"})
#: Word's own list style, which a loose item's follow-on paragraph carries.
LIST_STYLE = "ListParagraph"


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


W_P = _w("p")
W_TBL = _w("tbl")
W_TR = _w("tr")
W_TC = _w("tc")
W_SDT = _w("sdt")
W_R = _w("r")
W_HYPERLINK = _w("hyperlink")
W_DRAWING = _w("drawing")
W_BR = _w("br")
W_DEL_TEXT = _w("delText")
W_FOOTNOTE_REF = _w("footnoteReference")
W_COMMENT_REF = _w("commentReference")

#: The four run-level revision holders and the view each contributes to.
_REVISIONS: dict[str, str] = {
    _w("ins"): "accepted",
    _w("moveTo"): "accepted",
    _w("del"): "original",
    _w("moveFrom"): "original",
}

_TWIPS_PER_LEVEL = 720


def address_comment(address: str, indent: str = "") -> str:
    """One address comment line, exactly as :data:`ADDRESS_COMMENT` matches it."""
    return f"{indent}<!-- {address} -->"


# ---------------------------------------------------------------------------
# escaping
# ---------------------------------------------------------------------------

#: The characters that would otherwise be read as markup inside a line.
_INLINE_ESCAPE = re.compile(r"([\\`*_\[\]<>])")
#: What would be read as a block marker at the start of a line.
_LEADING_ESCAPE = re.compile(r"^([#>+\-=|]|\d+[.)])")


def escape(text: str) -> str:
    """Escape what markdown would otherwise read as markup.

    Deliberately narrow: the six inline characters that start emphasis, code,
    a link or raw HTML. Block markers are escaped separately, once per line,
    by :func:`_escape_leading`, so that a paragraph beginning ``1. `` does not
    become a list and ``- `` does not become a bullet.
    """
    return _INLINE_ESCAPE.sub(r"\\\1", text)


def _escape_leading(line: str) -> str:
    """Escape a block marker at the very start of a rendered line."""
    return _LEADING_ESCAPE.sub(r"\\\1", line, count=1)


# ---------------------------------------------------------------------------
# the rendering context
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class _Context:
    """What the renderer needs beyond the element it is looking at."""

    body: Any
    addresses: bool = False
    view: str = "accepted"
    part: Any = None
    package: Any = None
    footnotes: bool = True
    #: >0 while a GFM header row is being built: its bold is convention, not
    #: markup, so it is not marked (Java ``WmlToMarkdown.tableHeaderDepth``).
    header_depth: int = 0
    #: ``w:footnoteReference/@w:id`` -> the GFM label, in reference order.
    footnote_labels: dict[str, str] = dataclasses.field(default_factory=dict)
    #: What was dropped, for the caller who wants to know.
    warnings: list[str] = dataclasses.field(default_factory=list)
    #: The story's list labels, counted once by the numbering emulator.
    _labels: dict[int, Any] | None = None
    #: The comments part's texts, read once as bytes.
    _comments: dict[str, str] | None = None

    def warn(self, message: str) -> None:
        """Record something the markdown could not carry."""
        if message not in self.warnings:
            self.warnings.append(message)


def _root(part: Any) -> Any:
    """A part's root element as lxml, from the bytes it would be saved as.

    The same trick :mod:`docx4j_py.model.content.describe` uses: nothing is
    unmarshalled, so an untouched part stays untouched.
    """
    if part is None:
        return None
    from lxml import etree

    try:
        return etree.fromstring(part.xml)
    except Exception:  # noqa: BLE001 - a part that will not parse says nothing
        return None


def _list_labels(context: _Context) -> dict[int, Any]:
    """The story's list labels, from the numbering emulator, counted once.

    CR-003 Phase H: the marker's number is what **Word** paints, so a restart, a
    ``w:startOverride``, a ``w:lvlRestart`` and a list numbered through a
    paragraph style all come out right, where the renderer's own count (Phase K,
    section 13.6) restarted at every markdown block.
    """
    if context._labels is not None:
        return context._labels
    from docx4j_py.model.listnumbering import labels_for

    body = context.body
    context._labels = labels_for(body) if body is not None else {}
    return context._labels


def _label_of(element: Any, context: _Context) -> Any:
    """The list label of a paragraph, or None when it is not a list item."""
    return _list_labels(context).get(id(element))


def _comment_texts(context: _Context) -> dict[str, str]:
    """``w:comment/@w:id`` -> its text, from the comments part's bytes."""
    if context._comments is not None:
        return context._comments
    out: dict[str, str] = {}
    part = None
    main = getattr(context.package, "main_document_part", None)
    for candidate in getattr(main, "parts_by_type", lambda _t: [])("comments") or ():
        part = candidate
    if part is None:
        part = _part_of_type(context.package, "CommentsPart")
    root = _root(part)
    if root is not None:
        for comment in root.findall(_w("comment")):
            identifier = comment.get(_w("id")) or ""
            text = " ".join(
                "".join(node.itertext()) for node in comment.iter(_w("t")) if node.text
            )
            out[identifier] = " ".join(text.split())
    context._comments = out
    return out


def _part_of_type(package: Any, class_name: str) -> Any:
    """The first part of a package whose class has that name, or None."""
    parts = getattr(package, "parts", None)
    if parts is None:
        return None
    for part in getattr(parts, "values", lambda: ())():
        if type(part).__name__ == class_name:
            return part
    return None


def _footnotes_part(package: Any) -> Any:
    """The document's ``/word/footnotes.xml`` part, or None."""
    return _part_of_type(package, "FootnotesPart")


# ---------------------------------------------------------------------------
# inline rendering
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class _Sink:
    """Accumulates same-formatted text and wraps it on flush.

    Java's ``WmlToMarkdown.InlineSink``: emphasis markers go round a *run* of
    equally formatted text rather than round each run, so two bold runs in a
    row are one ``**...**``.
    """

    out: list[str] = dataclasses.field(default_factory=list)
    buffer: str = ""
    bold: bool = False
    italic: bool = False
    strike: bool = False

    def text(self, value: str, *, bold: bool, italic: bool, strike: bool) -> None:
        """Add text with its formatting, flushing when the formatting changes."""
        if self.buffer and (bold, italic, strike) != (self.bold, self.italic, self.strike):
            self.flush()
        self.bold, self.italic, self.strike = bold, italic, strike
        self.buffer += value

    def node(self, markup: str) -> None:
        """Add already-rendered markup, which no emphasis wraps."""
        self.flush()
        self.out.append(markup)

    def flush(self) -> None:
        """Close the open emphasis and push what has accumulated."""
        if not self.buffer:
            return
        # a run of whitespace-only text carries no emphasis worth marking, and
        # ``** **`` is not emphasis in CommonMark anyway
        rendered = escape(self.buffer)
        if self.buffer.strip():
            lead = rendered[: len(rendered) - len(rendered.lstrip())]
            tail = rendered[len(rendered.rstrip()) :]
            core = rendered.strip()
            if self.bold:
                core = f"**{core}**"
            if self.italic:
                core = f"*{core}*"
            if self.strike:
                core = f"~~{core}~~"
            rendered = f"{lead}{core}{tail}"
        self.out.append(rendered)
        self.buffer = ""

    def value(self) -> str:
        """Everything, flushed."""
        self.flush()
        return "".join(self.out)


def _on(value: Any) -> bool:
    """A ``BooleanDefaultTrue`` toggle is on when present and not ``val=0``."""
    return value is not None and getattr(value, "val", None) is not False


def _style_id(r_pr: Any) -> str | None:
    style = getattr(r_pr, "r_style", None) if r_pr is not None else None
    value = getattr(style, "val", None) if style is not None else None
    return str(getattr(value, "value", value)) if value else None


def _is_code(r_pr: Any) -> bool:
    """A code run: a code character style, or a monospace ``w:rFonts``."""
    if _style_id(r_pr) in CODE_STYLE_IDS:
        return True
    fonts = getattr(r_pr, "r_fonts", None) if r_pr is not None else None
    ascii_font = getattr(fonts, "ascii", None) if fonts is not None else None
    return bool(ascii_font) and str(ascii_font).lower() in MONOSPACE_FONTS


def _hyperlink_target(element: Any, context: _Context) -> str | None:
    """``w:hyperlink``'s destination: its relationship's target, or ``#anchor``."""
    rel_id = getattr(element, "id", None)
    if rel_id:
        rels = getattr(context.part, "relationships_part", None)
        rel = rels.get_relationship_by_id(rel_id) if rels is not None else None
        target = getattr(rel, "target", None) if rel is not None else None
        if target:
            return str(target)
    anchor = getattr(element, "anchor", None)
    return f"#{anchor}" if anchor else None


def _image_markup(drawing: Any, context: _Context) -> str:
    """``![alt](media/imageN.png)`` for every picture in a ``w:drawing``."""
    out: list[str] = []
    for frame in getattr(drawing, "anchor_or_inline", None) or ():
        doc_pr = getattr(frame, "doc_pr", None)
        alt = getattr(doc_pr, "descr", None) or getattr(doc_pr, "name", None) or ""
        graphic = getattr(frame, "graphic", None)
        data = getattr(graphic, "graphic_data", None)
        embed = None
        for candidate in getattr(data, "any_element", None) or ():
            blip_fill = getattr(candidate, "blip_fill", None)
            blip = getattr(blip_fill, "blip", None)
            embed = getattr(blip, "embed", None)
            if embed:
                break
        if not embed:
            context.warn("a drawing that is not a picture was dropped")
            continue
        rels = getattr(context.part, "relationships_part", None)
        rel = rels.get_relationship_by_id(embed) if rels is not None else None
        part = rels.get_part(rel) if (rels is not None and rel is not None) else None
        name = str(getattr(part, "part_name", "") or "")
        if not name:
            context.warn(f"an image whose relationship {embed} does not resolve was dropped")
            continue
        # ``/word/media/image1.png`` is ``media/image1.png`` beside the markdown
        destination = name[len("/word/") :] if name.startswith("/word/") else name.lstrip("/")
        out.append(f"![{escape(str(alt))}]({destination})")
    return "".join(out)


def _footnote_label(identifier: Any, context: _Context) -> str:
    """The GFM label for a footnote id, allocated in reference order."""
    key = str(identifier)
    label = context.footnote_labels.get(key)
    if label is None:
        label = str(len(context.footnote_labels) + 1)
        context.footnote_labels[key] = label
    return label


def _render_run(run: Any, sink: _Sink, context: _Context, *, revision: str | None) -> None:
    """One ``w:r``: its text with its formatting, and what is not text."""
    r_pr = getattr(run, "r_pr", None)
    code = _is_code(r_pr)
    bold = _on(getattr(r_pr, "b", None)) and context.header_depth == 0
    italic = _on(getattr(r_pr, "i", None))
    strike = _on(getattr(r_pr, "strike", None))
    content = getattr(run, "content", None)
    if not isinstance(content, list):
        return
    for item in content:
        qname = element_name(item)
        if qname == W_DRAWING:
            markup = _image_markup(item, context)
            if markup:
                sink.node(markup)
            continue
        if qname == W_FOOTNOTE_REF:
            if context.footnotes:
                sink.node(f"[^{_footnote_label(getattr(item, 'id', ''), context)}]")
            continue
        if qname == W_COMMENT_REF:
            if context.view == "markup":
                text = _comment_texts(context).get(str(getattr(item, "id", "")))
                if text:
                    sink.node(f"{{>>{text}<<}}")
            continue
        if qname == W_BR:
            sink.node("  \n")
            continue
        if qname == W_DEL_TEXT:
            if revision not in ("del", "moveFrom"):
                continue
            value = getattr(item, "value", None)
            text = value if isinstance(value, str) else ""
        else:
            text = item_text(qname, item)
        if not text:
            continue
        if code:
            sink.node(f"`{text}`")
            continue
        sink.text(text, bold=bold, italic=italic, strike=strike)


def _render_inlines(items: Any, sink: _Sink, context: _Context, *, revision: str | None) -> None:
    """The run-level content of a paragraph, a hyperlink, a revision, a control."""
    for element in items or ():
        name = element_name(element)
        if name == W_R:
            _render_run(element, sink, context, revision=revision)
            continue
        shown = _REVISIONS.get(name or "")
        if shown is not None:
            kind = (name or "").rpartition("}")[2]
            if context.view == "markup":
                nested = _Sink()
                _render_inlines(run_items_of(element), nested, context, revision=kind)
                inner = nested.value()
                if inner:
                    if kind in ("ins", "moveTo"):
                        sink.node(f"{{++{inner}++}}")
                    else:
                        sink.node(f"{{--{inner}--}}")
                continue
            if shown != "accepted":
                continue
            _render_inlines(run_items_of(element), sink, context, revision=kind)
            continue
        if name == W_HYPERLINK:
            target = _hyperlink_target(element, context)
            nested = _Sink()
            _render_inlines(run_items_of(element), nested, context, revision=revision)
            inner = nested.value()
            sink.node(f"[{inner}]({target})" if target else inner)
            continue
        if name in RUN_HOLDER_NAMES:
            _render_inlines(run_items_of(element), sink, context, revision=revision)


def _inline_text(element: Any, context: _Context) -> str:
    """A paragraph's inline content as markdown, with no block marker."""
    sink = _Sink()
    _render_inlines(run_items_of(element), sink, context, revision=None)
    return sink.value()


# ---------------------------------------------------------------------------
# blocks
# ---------------------------------------------------------------------------


def _p_pr(element: Any) -> Any:
    return getattr(element, "p_pr", None)


def _paragraph_style_id(element: Any) -> str | None:
    p_pr = _p_pr(element)
    style = getattr(p_pr, "p_style", None) if p_pr is not None else None
    value = getattr(style, "val", None) if style is not None else None
    value = getattr(value, "value", value)
    return str(value) if value else None


def _has_num_pr(element: Any) -> bool:
    """Whether the paragraph states a ``w:numPr`` of its own."""
    p_pr = _p_pr(element)
    return getattr(p_pr, "num_pr", None) is not None if p_pr is not None else False


def _is_list_continuation(element: Any) -> bool:
    """Whether this paragraph is a loose list item's follow-on paragraph.

    Java ``WmlToMarkdown``'s rule, and the one the importer writes: a paragraph
    of the list style with no ``w:numPr`` of its own, immediately after an item
    (CR-003 section 13.4, decided in Phase H).
    """
    return _paragraph_style_id(element) == LIST_STYLE and not _has_num_pr(element)


def _quote_depth(element: Any) -> int:
    """How deep a ``Quote``-styled paragraph is nested, from its left indent."""
    p_pr = _p_pr(element)
    ind = getattr(p_pr, "ind", None) if p_pr is not None else None
    left = getattr(ind, "left", None) if ind is not None else None
    try:
        return 1 + int(left) // _TWIPS_PER_LEVEL if left is not None else 1
    except (TypeError, ValueError):  # pragma: no cover - malformed indent
        return 1


def _cell_alignment(cell: Any) -> str:
    """A GFM column alignment from the first paragraph's ``w:jc``."""
    for block in block_children_of(cell) or ():
        if element_name(block) != W_P:
            continue
        p_pr = _p_pr(block)
        jc = getattr(p_pr, "jc", None) if p_pr is not None else None
        value = getattr(jc, "val", None) if jc is not None else None
        value = str(getattr(value, "value", value) or "")
        if value == "center":
            return "center"
        if value in ("right", "end"):
            return "right"
        return "left"
    return "left"


def _rows_of(table: Any) -> list[Any]:
    """A table's rows, a row-level ``w:sdt`` or ``w:customXml`` unwrapped."""
    out: list[Any] = []
    for item in block_children_of(table) or ():
        name = element_name(item)
        if name == W_TR:
            out.append(item)
        elif name in (W_SDT, _w("customXml")):
            out.extend(_rows_of(item))
    return out


def _cells_of(row: Any) -> list[Any]:
    out: list[Any] = []
    for item in block_children_of(row) or ():
        name = element_name(item)
        if name == W_TC:
            out.append(item)
        elif name in (W_SDT, _w("customXml")):
            out.extend(_cells_of(item))
    return out


def _cell_markdown(cell: Any, context: _Context) -> str:
    """A cell's content as one line: GFM cells are inline only."""
    pieces: list[str] = []
    for block in block_children_of(cell) or ():
        name = element_name(block)
        if name == W_P:
            pieces.append(_inline_text(block, context))
        elif name == W_TBL:
            context.warn("a nested table was flattened to its text (GFM cells are inline)")
            pieces.append(escape(" ".join(_table_text(block))))
        elif name == W_SDT:
            pieces.append(_cell_markdown(block, context))
    text = "<br>".join(piece for piece in pieces if piece.strip())
    return text.replace("|", "\\|").replace("\n", " ")


def _table_text(table: Any) -> list[str]:
    from docx4j_py.traversal import text_of

    return [text_of(table).replace("\n", " ")]


def table_markdown(element: Any, context: Any = None, **options: Any) -> str:
    """A ``w:tbl`` as a GFM pipe table. Phase C hangs ``Table.to_markdown`` here.

    The first row is the header, as GFM requires; ``w:gridSpan`` pads with
    empty cells so the columns stay aligned, and a ``w:vMerge`` continuation
    contributes nothing because the merge's top cell already holds the content
    (Java ``WmlToMarkdown.table``).

    Args:
        element: the ``w:tbl``.
        context: the renderer's context; a :class:`Body` or None is accepted,
            so that a caller with only an element can render one.
        **options: ``addresses`` and ``view``, when no context is given.
    """
    ctx = context if isinstance(context, _Context) else _context_for(context, **options)
    rows = _rows_of(element)
    if not rows:
        return ""
    lines: list[str] = []
    alignments: list[str] = []
    width = 0
    for index, row in enumerate(rows):
        cells: list[str] = []
        if index == 0:
            ctx.header_depth += 1
        for cell in _cells_of(row):
            tc_pr = getattr(cell, "tc_pr", None)
            merge = getattr(tc_pr, "v_merge", None) if tc_pr is not None else None
            # ``w:vMerge`` with no ``w:val`` continues the merge; ``val`` is an
            # enum, so compare its value rather than its str() (which is the
            # member's qualified name)
            merge_value = getattr(merge, "val", None) if merge is not None else None
            merge_value = getattr(merge_value, "value", merge_value)
            if merge is not None and merge_value != "restart":
                cells.append("")
            else:
                cells.append(_cell_markdown(cell, ctx))
            if index == 0:
                alignments.append(_cell_alignment(cell))
            span = getattr(tc_pr, "grid_span", None) if tc_pr is not None else None
            span_value = getattr(span, "val", None) if span is not None else None
            try:
                extra = int(span_value) - 1 if span_value is not None else 0
            except (TypeError, ValueError):  # pragma: no cover - malformed span
                extra = 0
            for _ in range(max(0, extra)):
                cells.append("")
                if index == 0:
                    alignments.append("left")
        if index == 0:
            ctx.header_depth -= 1
        width = max(width, len(cells))
        lines.append(cells)  # type: ignore[arg-type]
    padded = [[*row, *([""] * (width - len(row)))] for row in lines]  # type: ignore[misc]
    while len(alignments) < width:
        alignments.append("left")
    rule = {"left": "---", "center": ":-:", "right": "--:"}
    out = ["| " + " | ".join(padded[0]) + " |"]
    out.append("| " + " | ".join(rule[a] for a in alignments[:width]) + " |")
    for row in padded[1:]:
        out.append("| " + " | ".join(row) + " |")
    return "\n".join(out)


def _context_for(body: Any, **options: Any) -> _Context:
    """A context from a body (or None), for a caller with only an element."""
    part = getattr(body, "part", None)
    return _Context(
        body=body,
        addresses=bool(options.get("addresses", False)),
        view=str(options.get("view", "accepted")),
        part=part,
        package=getattr(body, "package", None),
        footnotes=bool(options.get("footnotes", True)),
    )


@dataclasses.dataclass(slots=True)
class _OpenList:
    """The list being built: one markdown block per run of list paragraphs.

    It counts nothing: since Phase H the marker's number is the emulator's
    count, so what this holds is the block's lines, which top-level ``w:numId``
    it belongs to and where a loose item's follow-on paragraph is indented to
    (CR-003 sections 13.4 and 18).
    """

    lines: list[str] = dataclasses.field(default_factory=list)
    num_id: str | None = None
    #: The column an item's own continuation lines start at: the indent of the
    #: last item plus the width of its marker, as CommonMark asks.
    content_indent: str = "  "
    #: The column each level's items are indented to, by ``w:ilvl``: a nested
    #: item starts at its parent's *content* column, which is three characters
    #: under ``1. `` and two under ``- ``.
    indents: dict[int, str] = dataclasses.field(default_factory=dict)
    #: The two delimiters this block uses. A list that follows another list with
    #: nothing between them takes the *other* delimiter, because CommonMark
    #: reads two adjacent lists of one delimiter as a single list, and the
    #: second of them is a different ``w:numId`` (CR-003 section 18).
    ordered: str = "."
    bullet: str = "-"

    def indent(self, ilvl: int) -> str:
        """The prefix an item at this level is written at."""
        return self.indents.get(ilvl, "  " * ilvl)

    def marker(self, label: Any) -> str:
        """``- `` for a bullet item, ``N. `` for a numbered one.

        Markdown has one ordered marker, so a ``lowerLetter`` or ``upperRoman``
        level renders as ``N.`` --- but ``N`` is now the counter's own value,
        which is the half that was wrong before Phase H.
        """
        if label.result.is_bullet:
            return f"{self.bullet} "
        return f"{label.result.count}{self.ordered} "


def _render_blocks(items: Any, context: _Context, out: list[str]) -> None:
    """Every block of a container, in document order, as markdown blocks."""
    open_list: _OpenList | None = None
    #: ``(w:numId, the delimiter its top-level items used, "bullet" or
    #: "ordered")`` of the list block just rendered, cleared by anything else:
    #: what tells the next list whether it has to take the other delimiter to be
    #: a list of its own.
    previous: tuple[str, str, str] | None = None
    index = 0
    items = list(items or ())
    while index < len(items):
        element = items[index]
        name = element_name(element)

        if name == W_P:
            # the heading test comes first, as Java's does: a built-in Heading
            # style can carry a legacy w:numPr and is still a heading
            level = heading_level_of(element)
            label = None if (level is not None and level <= 6) else _label_of(element, context)
            if label is not None:
                num_id, ilvl = label.num_id, label.ilvl
                if open_list is not None and ilvl == 0 and open_list.num_id not in (None, num_id):
                    open_list = None  # a different top-level list: a new block
                if open_list is None:
                    open_list = _OpenList()
                    kind = "bullet" if label.result.is_bullet else "ordered"
                    if previous is not None and previous[0] != num_id and previous[2] == kind:
                        # CommonMark reads two adjacent lists of one delimiter as
                        # one list; the other delimiter makes this its own
                        if kind == "bullet":
                            open_list.bullet = "*" if previous[1] == "-" else "-"
                        else:
                            open_list.ordered = ")" if previous[1] == "." else "."
                    out.append("")
                if ilvl == 0:
                    open_list.num_id = num_id
                indent = open_list.indent(ilvl)
                marker = open_list.marker(label)
                open_list.content_indent = indent + " " * len(marker)
                open_list.indents[ilvl + 1] = open_list.content_indent
                if context.addresses:
                    open_list.lines.append(address_comment(_address_of(element, context), indent))
                text = _inline_text(element, context)
                open_list.lines.append(f"{indent}{marker}{text}".rstrip())
                out[-1] = "\n".join(open_list.lines)
                if ilvl == 0:
                    kind = "bullet" if label.result.is_bullet else "ordered"
                    delimiter = open_list.bullet if label.result.is_bullet else open_list.ordered
                    previous = (num_id, delimiter, kind)
                index += 1
                continue
            if open_list is not None and _is_list_continuation(element):
                # a loose item's follow-on paragraph: Java appends it to the
                # open item, and so does this since Phase H owns lists
                # (CR-003 section 13.4). An empty one contributes nothing and
                # does not close the list.
                text = _inline_text(element, context)
                if text.strip():
                    prefix = open_list.content_indent
                    open_list.lines.append("")
                    if context.addresses:
                        open_list.lines.append(
                            address_comment(_address_of(element, context), prefix)
                        )
                    open_list.lines.append(f"{prefix}{text}".rstrip())
                    out[-1] = "\n".join(open_list.lines)
                index += 1
                continue
            open_list = None
            previous = None
            index = _render_paragraph(items, index, context, out)
            continue

        open_list = None
        previous = None
        if name == W_TBL:
            rendered = table_markdown(element, context)
            if rendered:
                _emit(out, rendered, element, context)
        elif name == W_SDT:
            children = block_children_of(element)
            if children is not None:
                _render_blocks(children, context, out)
        elif name is not None:
            local = name.rpartition("}")[2]
            if local not in ("sectPr", "bookmarkStart", "bookmarkEnd", "proofErr"):
                context.warn(f"a w:{local} block has no markdown form and was dropped")
        index += 1


def _emit(out: list[str], markup: str, element: Any, context: _Context) -> None:
    """Push one block, with its address comment in front when asked for."""
    if context.addresses:
        out.append(f"{address_comment(_address_of(element, context))}\n{markup}")
    else:
        out.append(markup)


def _address_of(element: Any, context: _Context) -> str:
    """The block's address: its paraId when it has one, else its ordinal."""
    body = context.body
    if body is None:
        return "?"
    try:
        from docx4j_py.model.content.addresses import address_of

        return address_of(body, element)
    except Exception:  # noqa: BLE001 - an element outside the body has no address
        return "?"


def _render_paragraph(items: list, index: int, context: _Context, out: list[str]) -> int:
    """One paragraph (or a run of them, for a code block). Returns the next index."""
    element = items[index]
    style_id = _paragraph_style_id(element)

    if style_id == SOURCE_CODE_STYLE:
        lines: list[str] = []
        last = index
        for position in range(index, len(items)):
            candidate = items[position]
            if (
                element_name(candidate) != W_P
                or _paragraph_style_id(candidate) != SOURCE_CODE_STYLE
            ):
                break
            lines.extend(_code_lines(candidate))
            last = position
        body_text = "\n".join(lines)
        _emit(out, f"```\n{body_text}\n```", element, context)
        return last + 1

    level = heading_level_of(element)
    if level is not None and level <= 6:
        text = _inline_text(element, context)
        _emit(out, f"{'#' * level} {text}".rstrip(), element, context)
        return index + 1

    text = _inline_text(element, context)
    if style_id in QUOTE_STYLE_IDS:
        prefix = "> " * _quote_depth(element)
        _emit(out, f"{prefix}{text}".rstrip(), element, context)
        return index + 1

    if not text.strip():
        # markdown has no empty paragraphs; Java drops them too. An addressed
        # export keeps the comment, so that every block is still addressable.
        if context.addresses:
            out.append(address_comment(_address_of(element, context)))
        return index + 1
    _emit(out, _escape_leading(text), element, context)
    return index + 1


def _code_lines(element: Any) -> list[str]:
    """A ``SourceCode`` paragraph's text, one entry per ``w:br``."""
    lines: list[str] = []
    current = ""
    for run in run_items_of(element) or ():
        if element_name(run) != W_R:
            continue
        for item in getattr(run, "content", None) or ():
            qname = element_name(item)
            if qname == W_BR:
                lines.append(current)
                current = ""
                continue
            text = item_text(qname, item)
            if text:
                current += text
    lines.append(current)
    return lines


# ---------------------------------------------------------------------------
# footnotes
# ---------------------------------------------------------------------------


def _footnote_definitions(context: _Context) -> list[str]:
    """The referenced footnotes, as GFM ``[^1]: ...`` definitions, in order."""
    if not context.footnote_labels:
        return []
    part = _footnotes_part(context.package)
    if part is None:
        return []
    notes: dict[str, Any] = {}
    try:
        for note in getattr(part.contents, "footnote", None) or ():
            notes[str(getattr(note, "id", ""))] = note
    except Exception:  # noqa: BLE001 - a footnotes part that will not parse has none
        return []
    saved_part = context.part
    context.part = part
    out: list[str] = []
    for identifier, label in context.footnote_labels.items():
        note = notes.get(identifier)
        if note is None:
            context.warn(f"a reference to footnote {identifier}, which is not in the part")
            continue
        blocks: list[str] = []
        _render_blocks(getattr(note, "content", None) or (), context, blocks)
        text = "\n".join(block for block in blocks if block.strip()).strip()
        indented = text.replace("\n", "\n    ")
        out.append(f"[^{label}]: {indented}".rstrip())
    context.part = saved_part
    return out


# ---------------------------------------------------------------------------
# the public entry points
# ---------------------------------------------------------------------------


def _blocks(
    body: Any,
    *,
    addresses: bool,
    view: str,
    footnotes: bool,
) -> tuple[list[str], _Context]:
    context = _Context(
        body=body,
        addresses=addresses,
        view=view,
        part=getattr(body, "part", None),
        package=getattr(body, "package", None),
        footnotes=footnotes,
    )
    if view not in ("accepted", "markup"):
        from docx4j_py.model.content.errors import ContentError

        raise ContentError(
            f"to_markdown takes view='accepted' or 'markup', not {view!r}",
            code="markdown.view_invalid",
            hint="'accepted' reads the document as if every change were accepted",
        )
    out: list[str] = []
    _render_blocks(body.content, context, out)
    out = [block for block in out if block.strip()]
    if footnotes:
        out.extend(_footnote_definitions(context))
    return out, context


def _joined(blocks: list[str], max_chars: int | None) -> tuple[str, int, bool]:
    """The blocks joined, cut at a block boundary where the budget allows."""
    whole = "\n\n".join(blocks)
    if max_chars is None or len(whole) <= max_chars:
        return whole, len(whole), False
    kept: list[str] = []
    length = 0
    for block in blocks:
        extra = len(block) + (2 if kept else 0)
        if length + extra > max_chars:
            break
        kept.append(block)
        length += extra
    if kept:
        return "\n\n".join(kept), len(whole), True
    return whole[:max_chars], len(whole), True


def body_markdown(
    body: Any,
    *,
    addresses: bool = False,
    view: str = "accepted",
    max_chars: int | None = None,
    footnotes: bool = True,
) -> str:
    """A container of block content as markdown. What ``Body.to_markdown`` is."""
    blocks, context = _blocks(body, addresses=addresses, view=view, footnotes=footnotes)
    text, _chars, _truncated = _joined(blocks, max_chars)
    _record(body, context)
    return text


def markdown_budget_of(
    body: Any,
    max_chars: int | None = None,
    *,
    addresses: bool = False,
    view: str = "accepted",
    footnotes: bool = True,
) -> TextExcerpt:
    """:func:`body_markdown` with the flag: ``TextExcerpt(text, chars, truncated)``."""
    blocks, context = _blocks(body, addresses=addresses, view=view, footnotes=footnotes)
    text, chars, truncated = _joined(blocks, max_chars)
    _record(body, context)
    return TextExcerpt(text, chars, truncated)


def paragraph_markdown(
    paragraph: Any,
    *,
    addresses: bool = False,
    view: str = "accepted",
) -> str:
    """One paragraph as markdown: its block form, without footnote definitions.

    A footnote reference is still ``[^1]``; the definitions belong to the
    document, so only :func:`body_markdown` writes them.
    """
    body = getattr(paragraph, "parent_body", None)
    context = _Context(
        body=body,
        addresses=addresses,
        view=view,
        part=getattr(body, "part", None),
        package=getattr(body, "package", None),
        footnotes=True,
    )
    out: list[str] = []
    _render_blocks([paragraph.element], context, out)
    _record(body, context)
    return "\n\n".join(block for block in out if block.strip())


def _record(body: Any, context: _Context) -> None:
    """Put what was dropped on the open ``ChangeReport``, if there is one.

    A read opens none of its own --- ``to_markdown`` changes nothing --- so
    this only ever adds to a report a caller already opened, which is how
    ``insert_markdown``'s round trip in a test sees an export's warnings.
    """
    if not context.warnings:
        return
    from docx4j_py.model.content.reports import current_recorder

    recorder = current_recorder(getattr(body, "package", None))
    for message in context.warnings:
        recorder.warn(message)
