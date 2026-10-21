"""The Markdown writer docx4j's ``MarkdownExporter`` serialises with. CR-011 Phase D (section 14).

A port of commonmark-java 0.30.0's ``org.commonmark.renderer.markdown``
(``MarkdownWriter``, ``CoreMarkdownNodeRenderer``, ``MarkdownRenderer``'s
special characters) and of the Markdown renderers of its GFM tables,
strikethrough and footnotes extensions, with docx4j-markdown's math nodes
(``org.docx4j.markdown.math.MathExtension``'s renderer), after core-ts
CR-012's ``commonmark.mts`` (its section 16.2). Only the nodes docx4j's
``WmlToMarkdown`` builds are here --- no indented code, HTML block, soft break
or inline footnote --- and each renders as commonmark-java renders it, so the
Markdown is docx4j's byte for byte. No dependency is added.

Each node may carry a ``tag``; :func:`render` records, for each tag, the
offset of the line its node's first character is written on, which is how
:func:`~docx4j_py.model.text.markdown.markdown_blocks` cuts the Markdown at
block boundaries (core-ts CR-012 section 16.4, item 5).

commonmark-java's licence, which its ported code keeps::

    Copyright (c) 2015, Atlassian Pty Ltd
    All rights reserved.

    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright notice, this
      list of conditions and the following disclaimer.

    * Redistributions in binary form must reproduce the above copyright notice,
      this list of conditions and the following disclaimer in the documentation
      and/or other materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
    AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
    IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
    DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
    FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
    DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
    SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
    CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
    OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
    OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

from __future__ import annotations

import re
from typing import Any

__all__ = [
    "BlockQuote",
    "BulletList",
    "Code",
    "DisplayMath",
    "Document",
    "Emphasis",
    "FencedCodeBlock",
    "FootnoteDefinition",
    "FootnoteReference",
    "HardLineBreak",
    "Heading",
    "HtmlInline",
    "Image",
    "InlineMath",
    "Link",
    "ListBlock",
    "ListItem",
    "Node",
    "OrderedList",
    "Paragraph",
    "Strikethrough",
    "StrongEmphasis",
    "TableBlock",
    "TableBody",
    "TableCell",
    "TableHead",
    "TableRow",
    "Text",
    "ThematicBreak",
    "render",
]


# ---------------------------------------------------------------- nodes (org.commonmark.node and the extensions')


class Node:
    """A node of the Markdown tree: its parent and children (commonmark-java ``Node``)."""

    __slots__ = ("children", "parent", "tag")

    def __init__(self) -> None:
        self.parent: Node | None = None
        self.children: list[Node] = []
        #: What the node was made for, where a caller wants to know where its Markdown begins.
        self.tag: Any = None

    @property
    def first_child(self) -> Node | None:
        return self.children[0] if self.children else None

    @property
    def last_child(self) -> Node | None:
        return self.children[-1] if self.children else None

    def _index(self) -> int:
        siblings = self.parent.children if self.parent is not None else []
        for i, sibling in enumerate(siblings):
            if sibling is self:
                return i
        return -1

    @property
    def previous(self) -> Node | None:
        i = self._index()
        return self.parent.children[i - 1] if self.parent is not None and i > 0 else None

    @property
    def next(self) -> Node | None:
        i = self._index()
        if self.parent is None or i < 0 or i + 1 >= len(self.parent.children):
            return None
        return self.parent.children[i + 1]

    def append_child(self, child: Node) -> None:
        child._unlink()
        child.parent = self
        self.children.append(child)

    def prepend_child(self, child: Node) -> None:
        child._unlink()
        child.parent = self
        self.children.insert(0, child)

    def _unlink(self) -> None:
        if self.parent is None:
            return
        i = self._index()
        if i >= 0:
            del self.parent.children[i]
        self.parent = None


class Document(Node):
    __slots__ = ()


class Paragraph(Node):
    __slots__ = ()


class BlockQuote(Node):
    __slots__ = ()


class HardLineBreak(Node):
    __slots__ = ()


class StrongEmphasis(Node):
    __slots__ = ()


class Emphasis(Node):
    __slots__ = ()


class TableBlock(Node):
    __slots__ = ()


class TableHead(Node):
    __slots__ = ()


class TableBody(Node):
    __slots__ = ()


class TableRow(Node):
    __slots__ = ()


class ListItem(Node):
    __slots__ = ()


class Heading(Node):
    __slots__ = ("level",)

    def __init__(self, level: int) -> None:
        super().__init__()
        self.level = level


class ListBlock(Node):
    """A bullet or ordered list: tight when its items are not separated by blank lines."""

    __slots__ = ("tight",)

    def __init__(self) -> None:
        super().__init__()
        self.tight = True


class BulletList(ListBlock):
    __slots__ = ("marker",)

    def __init__(self) -> None:
        super().__init__()
        #: ``-`` where unset.
        self.marker: str | None = None


class OrderedList(ListBlock):
    __slots__ = ("marker_delimiter", "marker_start_number")

    def __init__(self) -> None:
        super().__init__()
        #: 1 where unset.
        self.marker_start_number: int | None = None
        #: ``.`` where unset.
        self.marker_delimiter: str | None = None


class FencedCodeBlock(Node):
    __slots__ = ("literal",)

    def __init__(self, literal: str) -> None:
        super().__init__()
        self.literal = literal


class ThematicBreak(Node):
    __slots__ = ("literal",)

    def __init__(self, literal: str | None) -> None:
        super().__init__()
        self.literal = literal


class HtmlInline(Node):
    __slots__ = ("literal",)

    def __init__(self, literal: str) -> None:
        super().__init__()
        self.literal = literal


class Text(Node):
    __slots__ = ("literal",)

    def __init__(self, literal: str) -> None:
        super().__init__()
        self.literal = literal


class Code(Node):
    __slots__ = ("literal",)

    def __init__(self, literal: str) -> None:
        super().__init__()
        self.literal = literal


class Link(Node):
    __slots__ = ("destination", "title")

    def __init__(self, destination: str, title: str | None = None) -> None:
        super().__init__()
        self.destination = destination
        self.title = title


class Image(Node):
    __slots__ = ("destination", "title")

    def __init__(self, destination: str, title: str | None = None) -> None:
        super().__init__()
        self.destination = destination
        self.title = title


class TableCell(Node):
    """A GFM cell: its alignment (``"LEFT"``, ``"CENTER"``, ``"RIGHT"`` or None) and whether it is a header's."""

    __slots__ = ("alignment", "header")

    def __init__(self, header: bool) -> None:
        super().__init__()
        self.header = header
        self.alignment: str | None = None


class Strikethrough(Node):
    __slots__ = ("delimiter",)

    def __init__(self, delimiter: str) -> None:
        super().__init__()
        self.delimiter = delimiter


class FootnoteDefinition(Node):
    __slots__ = ("label",)

    def __init__(self, label: str) -> None:
        super().__init__()
        self.label = label


class FootnoteReference(Node):
    __slots__ = ("label",)

    def __init__(self, label: str) -> None:
        super().__init__()
        self.label = label


class InlineMath(Node):
    """Inline math, ``$...$`` (docx4j ``InlineMath``)."""

    __slots__ = ("display_hint", "literal")

    def __init__(self, literal: str, display_hint: bool) -> None:
        super().__init__()
        self.literal = literal
        self.display_hint = display_hint


class DisplayMath(Node):
    """Display math, a ``$$`` block (docx4j ``DisplayMath``)."""

    __slots__ = ("literal",)

    def __init__(self, literal: str) -> None:
        super().__init__()
        self.literal = literal


# ---------------------------------------------------------------- the writer (MarkdownWriter)


class _Writer:
    """commonmark-java's ``MarkdownWriter``, with the tag starts (``starts``)."""

    __slots__ = (
        "_block_separator",
        "_line_start",
        "_out",
        "_prefixes",
        "_raw_escapes",
        "_size",
        "_tight",
        "at_line_start",
        "last_char",
        "pending_tag",
        "starts",
    )

    def __init__(self, starts: dict[int, int]) -> None:
        self._out: list[str] = []
        self._block_separator = 0
        self.last_char = "\0"
        self.at_line_start = True
        self._prefixes: list[str] = []
        self._tight: list[bool] = []
        self._raw_escapes: list[frozenset[str]] = []
        #: Each tag's start (by ``id``), the offset of the line its node's first character is on.
        self.starts = starts
        #: A tag waiting for its node's first character.
        self.pending_tag: Any = None
        #: Where the line being written starts.
        self._line_start = 0
        #: How much has been written.
        self._size = 0

    def value(self) -> str:
        return "".join(self._out)

    def raw(self, s: str) -> None:
        self._flush_block_separator()
        self._write(s, None, True)

    def text(self, s: str, escape: frozenset[str] | None) -> None:
        if not s:
            return
        self._flush_block_separator()
        self._write(s, escape, True)

    def line(self) -> None:
        self._write("\n", None, False)
        self._write_prefixes()
        self.at_line_start = True

    def block(self) -> None:
        # whether this is a tight or loose separator is decided now, as tight can change before the flush
        self._block_separator = 1 if self._is_tight() else 2
        self.at_line_start = True

    def push_prefix(self, prefix: str) -> None:
        self._prefixes.append(prefix)

    def write_prefix(self, prefix: str) -> None:
        tmp = self.at_line_start
        self.raw(prefix)
        self.at_line_start = tmp

    def pop_prefix(self) -> None:
        self._prefixes.pop()

    def push_tight(self, tight: bool) -> None:
        self._tight.append(tight)

    def pop_tight(self) -> None:
        self._tight.pop()

    def push_raw_escape(self, escape: frozenset[str]) -> None:
        self._raw_escapes.append(escape)

    def pop_raw_escape(self) -> None:
        self._raw_escapes.pop()

    def _write(self, s: str, escape: frozenset[str] | None, content: bool) -> None:
        """Writes `s`, escaping what `escape` and the raw escapes match; content starts a pending tag at its line."""
        if not s:
            self.at_line_start = False
            return
        if content and self.pending_tag is not None:
            key = id(self.pending_tag)
            if key not in self.starts:
                self.starts[key] = self._line_start
            self.pending_tag = None
        chunk = s
        if self._raw_escapes or escape is not None:
            parts: list[str] = []
            for c in s:
                if (escape is not None and c in escape) or any(c in e for e in self._raw_escapes):
                    # a line end can't be escaped with \: a numeric character reference
                    parts.append("&#10;" if c == "\n" else "\\" + c)
                else:
                    parts.append(c)
            chunk = "".join(parts)
        newline = chunk.rfind("\n")
        if newline >= 0:
            self._line_start = self._size + newline + 1
        self._out.append(chunk)
        self._size += len(chunk)
        self.last_char = s[-1]
        self.at_line_start = False

    def _write_prefixes(self) -> None:
        for prefix in self._prefixes:
            self._write(prefix, None, False)

    def _flush_block_separator(self) -> None:
        if self._block_separator != 0:
            self._write("\n", None, False)
            self._write_prefixes()
            if self._block_separator > 1:
                self._write("\n", None, False)
                self._write_prefixes()
            self._block_separator = 0

    def _is_tight(self) -> bool:
        return bool(self._tight) and self._tight[-1]


# ---------------------------------------------------------------- the renderer (CoreMarkdownNodeRenderer and the extensions')

# The characters escaped in text: commonmark-java's own, and the extensions' special characters as docx4j
# registers them (``|`` the tables', ``~`` strikethrough's, ``$`` docx4j's math); ``#`` too in a heading.
_TEXT_ESCAPE = frozenset("[]<>`*_&\n\\|~$")
_TEXT_ESCAPE_IN_HEADING = frozenset("[]<>`*_&\n\\|~$#")
_LINK_DESTINATION_NEEDS_ANGLE_BRACKETS = frozenset(" ()<>\n\\")
_LINK_DESTINATION_ESCAPE_IN_ANGLE_BRACKETS = frozenset("<>\n\\")
_LINK_TITLE_ESCAPE_IN_QUOTES = frozenset('"\n\\')
_PIPE = frozenset("|")
_ORDERED_LIST_MARKER = re.compile(r"^([0-9]{1,9})([.)])")


class _ListHolder:
    __slots__ = ("delimiter", "marker", "number", "parent")

    def __init__(self, parent: _ListHolder | None, marker: str, number: int, delimiter: str | None) -> None:
        self.parent = parent
        self.marker = marker
        self.number = number
        #: None for a bullet list.
        self.delimiter = delimiter


def render(document: Document, starts: dict[int, int] | None = None) -> str:
    """A Markdown tree as commonmark-java's ``MarkdownRenderer`` renders it, with docx4j's extensions.

    `starts` receives, for each node with a ``tag``, the offset of the line
    its first character is written on, keyed by the tag's ``id``; a tag
    whose node writes nothing gives way to the next node with the same tag.
    """
    renderer = _Renderer(_Writer(starts if starts is not None else {}))
    renderer.render(document)
    return renderer.writer.value()


class _Renderer:
    __slots__ = ("_columns", "_list_holder", "writer")

    def __init__(self, writer: _Writer) -> None:
        self.writer = writer
        self._list_holder: _ListHolder | None = None
        #: The current table's column alignments, from its header cells.
        self._columns: list[str | None] = []

    def render(self, node: Node) -> None:
        w = self.writer
        # a tag whose node wrote nothing gives way to the next; a tag starts once
        tagged = node.tag is not None and id(node.tag) not in w.starts
        if tagged:
            w.pending_tag = node.tag
        self._render_node(node)
        if tagged and w.pending_tag is node.tag:
            w.pending_tag = None

    def _children(self, parent: Node) -> None:
        for child in list(parent.children):
            self.render(child)

    def _render_node(self, node: Node) -> None:
        # commonmark-java's visitors, one branch a node
        w = self.writer
        if isinstance(node, Text):
            self._text(node)
        elif isinstance(node, Paragraph):
            self._children(node)
            w.block()
        elif isinstance(node, Document):
            self._children(node)
            w.line()
        elif isinstance(node, ThematicBreak):
            # ___ is not ambiguous with a * or - list marker
            w.raw(node.literal if node.literal is not None else "___")
            w.block()
        elif isinstance(node, Heading):
            self._heading(node)
        elif isinstance(node, FencedCodeBlock):
            self._fenced_code_block(node)
        elif isinstance(node, BlockQuote):
            w.write_prefix("> ")
            w.push_prefix("> ")
            self._children(node)
            w.pop_prefix()
            w.block()
        elif isinstance(node, (BulletList, OrderedList)):
            w.push_tight(node.tight)
            if isinstance(node, BulletList):
                self._list_holder = _ListHolder(self._list_holder, node.marker or "-", 0, None)
            else:
                start = node.marker_start_number if node.marker_start_number is not None else 1
                self._list_holder = _ListHolder(self._list_holder, "", start, node.marker_delimiter or ".")
            self._children(node)
            self._list_holder = self._list_holder.parent
            w.pop_tight()
            w.block()
        elif isinstance(node, ListItem):
            self._list_item(node)
        elif isinstance(node, Code):
            self._code(node.literal)
        elif isinstance(node, Emphasis):
            # nested emphasis takes the other delimiter
            delimiter = "_" if w.last_char == "*" else "*"
            w.raw(delimiter)
            self._children(node)
            w.raw(delimiter)
        elif isinstance(node, StrongEmphasis):
            w.raw("**")
            self._children(node)
            w.raw("**")
        elif isinstance(node, Link):
            self._link_like(node.title, node.destination, node, "[")
        elif isinstance(node, Image):
            self._link_like(node.title, node.destination, node, "![")
        elif isinstance(node, HtmlInline):
            w.raw(node.literal)
        elif isinstance(node, HardLineBreak):
            w.raw("  ")
            w.line()
        elif isinstance(node, TableBlock):
            self._columns.clear()
            w.push_tight(True)
            self._children(node)
            w.pop_tight()
            w.block()
        elif isinstance(node, TableHead):
            self._children(node)
            for alignment in self._columns:
                w.raw("|")
                w.raw(_DELIMITER_ROW.get(alignment, "---"))
            w.raw("|")
            w.block()
        elif isinstance(node, TableBody):
            self._children(node)
        elif isinstance(node, TableRow):
            self._children(node)
            w.raw("|")  # the | at the end of the line
            w.block()
        elif isinstance(node, TableCell):
            if node.parent is not None and isinstance(node.parent.parent, TableHead):
                self._columns.append(node.alignment)
            w.raw("|")
            w.push_raw_escape(_PIPE)
            self._children(node)
            w.pop_raw_escape()
        elif isinstance(node, Strikethrough):
            w.raw(node.delimiter)
            self._children(node)
            w.raw(node.delimiter)
        elif isinstance(node, FootnoteReference):
            w.raw("[^")
            w.raw(node.label)  # a label is written as it is
            w.raw("]")
        elif isinstance(node, FootnoteDefinition):
            w.raw("[^")
            w.raw(node.label)
            w.raw("]: ")
            w.push_prefix("    ")
            self._children(node)
            w.pop_prefix()
        elif isinstance(node, InlineMath):
            delimiter = "$$" if node.display_hint else "$"
            w.raw(delimiter)
            w.raw(node.literal.replace("\n", " "))
            w.raw(delimiter)
        elif isinstance(node, DisplayMath):
            w.raw("$$")
            w.line()
            for line in node.literal.split("\n"):
                w.raw(line)
                w.line()
            w.raw("$$")
            w.block()
        else:
            raise TypeError(f"No Markdown renderer for {type(node).__name__}")

    def _heading(self, heading: Heading) -> None:
        w = self.writer
        if heading.level <= 2 and _has_line_break(heading):
            # a Setext heading can have several lines, but only at levels 1 and 2
            self._children(heading)
            w.line()
            w.raw("===" if heading.level == 1 else "---")
            w.block()
            return
        w.raw("#" * heading.level)
        w.raw(" ")
        self._children(heading)
        w.block()

    def _fenced_code_block(self, code: FencedCodeBlock) -> None:
        w = self.writer
        literal = code.literal
        # a fence longer than any run of backticks in the code
        fence = "`" * max(_max_run_length("`", literal) + 1, 3)
        w.raw(fence)
        w.line()
        if literal:
            for line in _lines_of(literal):
                w.raw(line)
                w.line()
        w.raw(fence)
        w.block()

    def _list_item(self, item: ListItem) -> None:
        w = self.writer
        holder = self._list_holder
        if holder is None:
            raise ValueError("A list item outside a list")
        if holder.delimiter is None:
            marker = holder.marker
        else:
            marker = f"{holder.number}{holder.delimiter}"
            holder.number += 1
        spaces = " "
        w.write_prefix(marker)
        w.write_prefix(spaces)
        w.push_prefix(" " * (len(marker) + len(spaces)))
        if item.first_child is None:
            w.block()  # an empty item
        else:
            self._children(item)
        w.pop_prefix()

    def _code(self, literal: str) -> None:
        w = self.writer
        # backticks in the code: one more around it
        backticks = _max_run_length("`", literal)
        w.raw("`" * (backticks + 1))
        # a space inside where the code starts or ends with a backtick, or starts and ends with a space
        add_space = (
            literal.startswith("`")
            or literal.endswith("`")
            or (literal.startswith(" ") and literal.endswith(" ") and literal.strip(" ") != "")
        )
        if add_space:
            w.raw(" ")
        w.raw(literal)
        if add_space:
            w.raw(" ")
        w.raw("`" * (backticks + 1))

    def _text(self, text: Text) -> None:
        # When in doubt, escape: a character special only at a line's start is escaped there alone.
        w = self.writer
        literal = text.literal
        if w.at_line_start and literal:
            c = literal[0]
            if c == "-":
                w.raw("\\-")  # a bullet list marker
                literal = literal[1:]
            elif c == "#":
                w.raw("\\#")  # an ATX heading
                literal = literal[1:]
            elif c == "=":
                if text.previous is not None:  # a Setext heading, unless the block's first line
                    w.raw("\\=")
                    literal = literal[1:]
            elif "0" <= c <= "9":
                m = _ORDERED_LIST_MARKER.match(literal)
                if m:
                    w.raw(m.group(1))
                    w.raw("\\" + m.group(2))
                    literal = literal[m.end() :]
            elif c == "\t":
                w.raw("&#9;")
                literal = literal[1:]
            elif c == " ":
                w.raw("&#32;")
                literal = literal[1:]
        escape = _TEXT_ESCAPE_IN_HEADING if isinstance(text.parent, Heading) else _TEXT_ESCAPE
        if literal.endswith("!") and isinstance(text.next, Link):
            # an unescaped ! would make the link an image
            w.text(literal[:-1], escape)
            w.raw("\\!")
        else:
            w.text(literal, escape)

    def _link_like(self, title: str | None, destination: str, node: Node, opener: str) -> None:
        w = self.writer
        w.raw(opener)
        self._children(node)
        w.raw("]")
        w.raw("(")
        if any(c in _LINK_DESTINATION_NEEDS_ANGLE_BRACKETS for c in destination):
            w.raw("<")
            w.text(destination, _LINK_DESTINATION_ESCAPE_IN_ANGLE_BRACKETS)
            w.raw(">")
        else:
            w.raw(destination)
        if title is not None:
            w.raw(" ")
            w.raw('"')
            w.text(title, _LINK_TITLE_ESCAPE_IN_QUOTES)
            w.raw('"')
        w.raw(")")


_DELIMITER_ROW = {"LEFT": ":---", "RIGHT": "---:", "CENTER": ":---:"}


def _has_line_break(node: Node) -> bool:
    return any(isinstance(c, HardLineBreak) or _has_line_break(c) for c in node.children)


def _max_run_length(needle: str, s: str) -> int:
    longest = 0
    pos = 0
    while pos < len(s):
        pos = s.find(needle, pos)
        if pos == -1:
            return longest
        run = 0
        while True:
            pos += len(needle)
            run += 1
            if not s.startswith(needle, pos):
                break
        longest = max(run, longest)
    return longest


def _lines_of(literal: str) -> list[str]:
    """A literal's lines, ``\\n`` being a terminator: ``"abc\\n"`` is one line."""
    parts = literal.split("\n")
    return parts[:-1] if parts[-1] == "" else parts
