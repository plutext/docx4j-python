"""Markdown *out*: the content API's Markdown readers, over the text model. CR-003 section 3.5, CR-011 Phase D.

Since CR-011 Phase D (its section 14) ``Body.to_markdown``,
``Paragraph.to_markdown``, ``Table.to_markdown`` and ``pkg.to_markdown`` are
:func:`docx4j_py.model.text.markdown` over the text model's reading
(:func:`docx4j_py.model.text.extract`), the port of docx4j-markdown's
``MarkdownExporter``: the same Markdown as docx4j's, byte for byte, the text
the text model's (REQ-202) --- text boxes after their anchors, notes ``[^1]``
and ``[^e1]`` with their definitions after the body, hidden text left out,
fields' results, bound values, charts, equations as LaTeX --- and REQ-703's
rules for lists, links, pictures, code and tables.

Three views (REQ-401): ``"accepted"`` (the default; ``"current"`` is its
synonym), the document as Accept All leaves it, deleted paragraph marks joined;
``"original"``, as Reject All leaves it; and ``"markup"``, CriticMarkup written
flat: an insertion ``{++...++}``, a deletion ``{--...--}``, a comment
``{>>...<<}`` at its reference.

**Addresses.** ``to_markdown(addresses=True)`` puts each block's address in an
HTML comment on **its own line immediately before the block**, indented as the
block's first line is::

    <!-- w14:5A2B1C3D -->
    # Quarterly Report

    <!-- body/1 -->
    The quick brown fox.

A block of the body's own story has its content API address, the
``w14:paraId`` when it has one and the ordinal otherwise, which
:meth:`~docx4j_py.model.content.body.Body.element_at` accepts; a text box's, a
note's or a chart's block has the text model's (``body/3/textbox:0/1``,
``footnote:1/0``, REQ-519), which is for reading only (CR-011 Q9).
:data:`ADDRESS_COMMENT` is the regex that recovers them. The addressed output
is for **reading and addressing**, not for feeding back through
``insert_markdown``: a comment line between two list items ends the list, as
CommonMark says it must.

**Budgets.** ``max_chars`` cuts at a block boundary, by
:func:`docx4j_py.model.text.markdown_blocks` (core-ts ``MarkdownExporter.blocks``),
and in the middle of a block only when the first block is already over budget.

**Pictures** point at their part, relative to ``/word/``
(``media/image1.png``), where :func:`docx4j_py.model.text.markdown` itself
writes docx4j's data URI: a reader's Markdown goes into a context window
(CR-011 section 14).

Nothing is unmarshalled by a read beyond the body's own part: the text model
reads every other part with ``read_contents`` (REQ-204).
"""

from __future__ import annotations

import re
from enum import StrEnum
from typing import Any, Literal

from docx4j_py.model.content.reports import TextExcerpt

__all__ = [
    "ADDRESS_COMMENT",
    "CODE_STYLE_IDS",
    "MONOSPACE_FONTS",
    "MarkdownView",
    "MarkdownViewValue",
    "address_comment",
    "body_markdown",
    "escape",
    "markdown_budget_of",
    "package_markdown",
    "package_markdown_budget",
    "paragraph_markdown",
    "table_markdown",
]


# ---------------------------------------------------------------------------
# the values (CR-003 section 3.1: a Literal, with a StrEnum beside it)
# ---------------------------------------------------------------------------

#: Which view of a tracked document ``to_markdown`` renders (REQ-401).
MarkdownViewValue = Literal["accepted", "original", "markup"]


class MarkdownView(StrEnum):
    """``docx_to_markdown``'s ``tracked_changes`` switch (docx4j-mcp), and the text model's views."""

    ACCEPTED = "accepted"
    """Every tracked change accepted: insertions in, deletions out, deleted marks joined."""
    ORIGINAL = "original"
    """Every tracked change rejected: the document as it read before them."""
    MARKUP = "markup"
    """CriticMarkup: ``{++inserted++}``, ``{--deleted--}``, ``{>>comment<<}``."""


_VIEWS = ("accepted", "original", "markup", "current")

#: The exact form of an address comment, and how to recover one. The whole
#: line is the comment; ``match.group("address")`` is what ``element_at`` takes.
ADDRESS_COMMENT = re.compile(r"^[ \t]*<!--\s(?P<address>\S+)\s-->$", re.MULTILINE)

#: Character style ids that mean "this run is code" (docx4j ``ImportStyles``),
#: which the importer writes; the text model reads code from the font.
CODE_STYLE_IDS: frozenset[str] = frozenset({"CodeChar", "Code", "HTMLCode"})

#: Fonts that mean "this run is code" (docx4j ``WmlToMarkdown.MONO_FONTS`` before
#: 17.3.4, lower-cased). The text model's set is the WD's 25 (REQ-703).
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


def address_comment(address: str, indent: str = "") -> str:
    """One address comment line, exactly as :data:`ADDRESS_COMMENT` matches it."""
    return f"{indent}<!-- {address} -->"


#: The characters that would otherwise be read as markup inside a line.
_INLINE_ESCAPE = re.compile(r"([\\`*_\[\]<>])")


def escape(text: str) -> str:
    """Escape the six inline characters that start emphasis, code, a link or raw HTML.

    Kept for callers of CR-003's exporter; the readers' Markdown is escaped by
    :mod:`docx4j_py.model.text.commonmark`, as commonmark-java escapes it.
    """
    return _INLINE_ESCAPE.sub(r"\\\1", text)


# ---------------------------------------------------------------------------
# the readers
# ---------------------------------------------------------------------------


def _view_of(view: str) -> str:
    if view not in _VIEWS:
        from docx4j_py.model.content.errors import ContentError

        raise ContentError(
            f"to_markdown takes view='accepted', 'original' or 'markup', not {view!r}",
            code="markdown.view_invalid",
            hint="'accepted' reads the document as if every change were accepted, "
            "'original' as if every one were rejected, 'markup' shows them",
        )
    return "accepted" if view == "current" else str(view)


def _word_relative(image: Any) -> str:
    """A picture's destination for a reader: its part's name relative to ``/word/`` (``media/image1.png``)."""
    name = str(getattr(getattr(image.part, "part_name", None), "name", "") or getattr(image.part, "part_name", ""))
    return name[len("/word/") :] if name.startswith("/word/") else name.lstrip("/")


def _blocks(body: Any, view: str, *, footnotes: bool = True) -> tuple[Any, Any]:
    """The body's reading and its Markdown block by block."""
    from docx4j_py.model.content.text_model import reading_of

    return _written(reading_of(body, view=_view_of(view), notes=footnotes))


def _written(reading: Any) -> tuple[Any, Any]:
    from docx4j_py.model.text import markdown_blocks

    return reading, markdown_blocks(reading.result, reading.package, image_destination=_word_relative)


def _pieces(body: Any, view: str, addresses: bool, footnotes: bool) -> list[str]:
    """The body's Markdown, a piece a block: each block's lines, its address comment before them where asked."""
    reading, blocks = _blocks(body, view, footnotes=footnotes)
    return _pieces_of(reading, blocks, body.prefix, addresses)


def _pieces_of(reading: Any, blocks: Any, prefix: str, addresses: bool) -> list[str]:
    pieces: list[str] = [blocks.front_matter] if blocks.front_matter else []
    for item in blocks.blocks:
        if not addresses:
            if item.markdown:
                pieces.append(item.markdown)
            continue
        if item.merged:
            continue  # its lines are in the fence of the block before it, which carries that block's address
        address = reading.address_of(prefix, item.block)
        if not item.markdown:
            # Markdown has no empty paragraphs; the comment keeps the block addressable
            pieces.append(address_comment(address) + "\n\n")
            continue
        first = item.markdown.split("\n", 1)[0]
        indent = first[: len(first) - len(first.lstrip(" \t"))]
        pieces.append(address_comment(address, indent) + "\n" + item.markdown)
    if len(pieces) == (1 if blocks.front_matter else 0):
        # nothing written for any block (an empty document): the renderer's last line end, docx4j's "\n"
        pieces.append(blocks.markdown[len(blocks.front_matter) :])
    return pieces


def _package_pieces(package: Any, view: str, addresses: bool, footnotes: bool) -> list[str]:
    """The package's Markdown, read from the package itself, so that no part is unmarshalled."""
    from docx4j_py.model.content.text_model import Reading
    from docx4j_py.model.text import extract

    reading = Reading(extract(package, view=_view_of(view), notes=footnotes), "body", package)
    _reading, blocks = _written(reading)
    return _pieces_of(reading, blocks, "body", addresses)


def _joined(pieces: list[str], max_chars: int | None) -> tuple[str, int, bool]:
    """The pieces joined, cut at a block boundary where the budget allows."""
    whole = "".join(pieces)
    if max_chars is None or len(whole) <= max_chars:
        return whole, len(whole), False
    kept: list[str] = []
    length = 0
    for piece in pieces:
        if length + len(piece) > max_chars:
            break
        kept.append(piece)
        length += len(piece)
    if kept:
        return "".join(kept).rstrip(), len(whole), True
    return whole[:max_chars], len(whole), True


def body_markdown(
    body: Any,
    *,
    addresses: bool = False,
    view: str = "accepted",
    max_chars: int | None = None,
    footnotes: bool = True,
) -> str:
    """A body as Markdown: what ``Body.to_markdown`` is.

    The text model's reading of the body (the main document's as the
    package's: its story, text boxes and notes) written by
    :func:`docx4j_py.model.text.markdown`; with `footnotes` off, no note is
    read (``notes: false``), so neither the references nor the definitions
    are written.
    """
    text, _chars, _cut = _joined(_pieces(body, view, addresses, footnotes), max_chars)
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
    text, chars, truncated = _joined(_pieces(body, view, addresses, footnotes), max_chars)
    return TextExcerpt(text, chars, truncated)


def package_markdown(
    package: Any,
    *,
    addresses: bool = False,
    view: str = "accepted",
    max_chars: int | None = None,
    footnotes: bool = True,
) -> str:
    """A package as Markdown: what ``pkg.to_markdown`` is.

    The text model's reading of the package (its main story, text boxes and
    notes), as :func:`body_markdown` writes the main body's, but read from the
    package: no part is unmarshalled, so a save afterwards writes every part
    the caller has not touched byte for byte (CR-011 section 14).
    """
    text, _chars, _cut = _joined(_package_pieces(package, view, addresses, footnotes), max_chars)
    return text


def package_markdown_budget(
    package: Any,
    max_chars: int | None = None,
    *,
    addresses: bool = False,
    view: str = "accepted",
    footnotes: bool = True,
) -> TextExcerpt:
    """:func:`package_markdown` with the flag: ``TextExcerpt(text, chars, truncated)``."""
    text, chars, truncated = _joined(_package_pieces(package, view, addresses, footnotes), max_chars)
    return TextExcerpt(text, chars, truncated)


def _story_body(body: Any) -> Any:
    """The body of the story a nested body (a cell's, a control's) is in, where it is a part's."""
    part = body.part
    if part is None:
        return body
    try:
        from docx4j_py.model.content.body import body_of

        story = body_of(part)
    except Exception:  # noqa: BLE001 - a part with no body: the body itself
        return body
    return story


def _find(blocks: Any, element: Any) -> tuple[Any, bool]:
    """The block read from `element` (or holding its text, joined into another), and whether it is a top-level one."""
    from docx4j_py.model.text.result import TextBlock

    def nested(items: list[TextBlock]) -> TextBlock | None:
        for block in items:
            if block.source is element:
                return block
            for row in block.rows:
                for cell in row.cells:
                    found = nested(cell.blocks)
                    if found is not None:
                        return found
        return None

    def holding(items: list[TextBlock]) -> TextBlock | None:
        for block in items:
            if any(r.source is not None and r.source.paragraph is element for r in block.runs):
                return block
            for row in block.rows:
                for cell in row.cells:
                    found = holding(cell.blocks)
                    if found is not None:
                        return found
        return None

    for item in blocks.blocks:
        if item.block.source is element:
            return item, True
    top = [item.block for item in blocks.blocks]
    found = nested(top) or holding(top)
    for item in blocks.blocks:
        if item.block is found:
            return item, True
    return found, False


def _alone(block: Any, reading: Any) -> str:
    """One block written on its own: a nested block's, or a code paragraph's merged into a fence."""
    from docx4j_py.model.text import markdown
    from docx4j_py.model.text.result import TextResult

    alone = TextResult([block], reading.result.comments, [], {}, reading.result.options)
    return markdown(alone, reading.package, image_destination=_word_relative)


def _block_markdown(
    body: Any, element: Any, *, addresses: bool, view: str, address: str | None
) -> str:
    """The Markdown of one block of a body: its lines in the whole reading, or the block written alone."""
    story = _story_body(body)
    reading, blocks = _blocks(story, view)
    item, top = _find(blocks, element)
    if item is None:
        text = ""
    elif top and not item.merged and item.markdown:
        text = item.markdown
    else:
        block = item.block if top else item
        text = _alone(block, reading)
    text = text.rstrip()
    if not addresses:
        return text
    where = address
    if where is None and item is not None:
        where = reading.address_of(story.prefix, item.block if top else item)
    if where is None:
        return text
    first = text.split("\n", 1)[0]
    indent = first[: len(first) - len(first.lstrip(" \t"))]
    return address_comment(where, indent) + ("\n" + text if text else "")


def paragraph_markdown(
    paragraph: Any,
    *,
    addresses: bool = False,
    view: str = "accepted",
) -> str:
    """One paragraph as Markdown: its lines in its story's Markdown, without the footnote definitions.

    A list item is written with the marker and the indent it has in the
    document (``3. Plums``, ``4. 2. Quince``), a follow-on paragraph indented
    under its item; a paragraph in a table cell, or a code paragraph merged
    into a fence, is written on its own. A footnote reference is still
    ``[^1]``; the definitions belong to the document. Trailing white space is
    left out.
    """
    body = getattr(paragraph, "parent_body", None)
    if body is None:
        return ""
    return _block_markdown(
        body, paragraph.element, addresses=addresses, view=view, address=getattr(paragraph, "address", None)
    )


def table_markdown(element: Any, context: Any = None, **options: Any) -> str:
    """A ``w:tbl`` as a GFM pipe table: what ``Table.to_markdown`` is.

    The table's block in its body's Markdown, as docx4j writes it (its header
    the ``w:tblHeader`` rows or the first row; a span padded with empty
    cells; a vertical merge's top cell holding its content; a cell's
    paragraphs ``<br>`` apart). Trailing white space is left out.

    Args:
        element: the ``w:tbl``.
        context: the :class:`~docx4j_py.model.content.body.Body` it is in, or
            None for an element alone (read with no styles).
        **options: ``view`` and ``addresses``.
    """
    view = str(options.get("view", "accepted"))
    addresses = bool(options.get("addresses", False))
    if context is None or not hasattr(context, "container"):
        from docx4j_py.model.text import extract, markdown

        result = extract(element, view=_view_of(view))
        return markdown(result, None, image_destination=_word_relative).rstrip()
    return _block_markdown(context, element, addresses=addresses, view=view, address=None)
