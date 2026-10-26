"""Markdown from the block model (WD section 7.3, REQ-703 and REQ-704). CR-011 Phase D (section 14).

docx4j-markdown's ``MarkdownExporter``, ``MarkdownExportOptions`` and
``WmlToMarkdown`` at docx4j ``198d015a3`` (their main sources unchanged since
``d9acdf5e2``), after core-ts CR-012's phase 3 (its section 16): the block
model is written as a Markdown tree, which :mod:`.commonmark`, a port of
commonmark-java 0.30.0's Markdown renderer, serialises as docx4j does, so the
escaping, the list markers and the block separators are docx4j's byte for
byte::

    from docx4j_py.model.text import extract, markdown, plain_text

    result = extract(pkg, view="markup")
    plain_text(result)                 # the plain text and the Markdown of one reading
    markdown(result, pkg)

What text is written is the text model's (REQ-202): its view, its stories
(the main story, each text box after its anchor, the notes), hidden text left
out, fields' results, bound values, charts and equations. This module decides
only how a block is written (REQ-703):

* a heading from its outline level, not its style's name; a list from its
  numbering, two ``w:num`` on one ``w:abstractNum`` one list, kept apart from
  an adjacent list by the other delimiter or bullet, and going on at Word's
  count after a paragraph that interrupts it; a paragraph indented to an open
  item's text (or in ``ListParagraph``) a follow-on paragraph of it; a label
  level (``NOTE``) a bold label;
* bold, italic, struck-through and code text where a run adds them to its
  paragraph's style; a code-styled paragraph, or one in a mono paragraph
  style, a fenced code block, consecutive ones one fence;
* links from a ``w:hyperlink`` or a HYPERLINK field; pictures with the
  describer's text, their image a data URI unless ``image_destination`` says
  otherwise; notes ``[^1]`` and ``[^e1]``, their definitions where the model
  reads them (after the body);
* GFM tables, the header the ``w:tblHeader`` rows or the first row, a cell's
  paragraphs ``<br>`` apart; a chart's note, title and table; block quotes from
  the Quote styles, nested by the indent beyond the style's own; thematic
  breaks; display math;
* in the ``markup`` view CriticMarkup, ``{++...++}``, ``{--...--}`` and
  ``{>>...<<}``, written flat, with ``markup_detail``;
* under ``metadata: true`` the document properties as YAML front matter
  (REQ-704).

The package is never changed (REQ-204). Imported on first use.
"""

from __future__ import annotations

import base64
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

from docx4j_py.model.text import commonmark as md
from docx4j_py.model.text.options import TextOptions, TextView
from docx4j_py.model.text.result import (
    TextBlock,
    TextCell,
    TextComment,
    TextNumbering,
    TextResult,
    TextRevision,
    TextRow,
    TextRun,
)
from docx4j_py.model.values import integer

__all__ = [
    "MarkdownBlock",
    "MarkdownBlocks",
    "MarkdownExportOptions",
    "MarkdownExporter",
    "MarkdownImage",
    "MarkdownImageDestination",
    "front_matter",
    "markdown",
    "markdown_blocks",
]

_SVG = "image/svg+xml"
_TWIPS_PER_LEVEL = 720

# the paragraph styles docx4j's importer writes, which the export reads back (docx4j ``ImportStyles``)
_QUOTE = "Quote"
_INTENSE_QUOTE = "IntenseQuote"
_LIST_PARAGRAPH = "ListParagraph"
_SOURCE_CODE = "SourceCode"

#: Paragraph styles which are code blocks by name, whatever their font.
_CODE_STYLES = frozenset({_SOURCE_CODE, "Code", "HTMLPreformatted", "MacroText"})


@dataclass(frozen=True, slots=True)
class MarkdownImage:
    """A picture's image, for a :data:`MarkdownImageDestination`."""

    #: The picture's run in the block model.
    run: TextRun
    #: The image part.
    part: Any
    data: bytes
    content_type: str
    #: The picture's number among the export's images with an image part, from 1.
    index: int
    #: docx4j's file name for it, ``image1.png``: ``image``, the index, and its content type's extension.
    file_name: str


#: Where a picture's Markdown image points: a data URI by default, or what a host makes of the image.
MarkdownImageDestination = Callable[[MarkdownImage], str]


@dataclass(slots=True)
class MarkdownExportOptions:
    """How Markdown is written. docx4j ``MarkdownExportOptions``.

    What text is written is the text model's, under its :class:`TextOptions`
    (:attr:`text_options`); the options here are how it is written. docx4j's
    ``imageDirPath`` and ``imageTargetUri``, which write each image to a
    directory, are :attr:`image_destination` here, as core-ts has it (its
    CR-012 section 16.4): the exporter writes no file, and a host that wants
    the files writes them in the callable, which is given docx4j's file name,
    and returns the link.
    """

    #: The text model's options a package is read with: the default is the
    #: text model's, the ``accepted`` view, the main story with its text boxes
    #: and notes, hidden text left out. A result is written under the options
    #: it was read with.
    text_options: TextOptions = field(default_factory=TextOptions)
    #: Where each picture points; a data URI (``data:image/png;base64,...``) where None.
    image_destination: MarkdownImageDestination | None = None

    @property
    def view(self) -> str:
        """The revision view: ``accepted`` (the default), ``original`` or ``markup``."""
        return self.text_options.view

    @view.setter
    def view(self, view: TextView) -> None:
        self.text_options.set("view", view)


@dataclass(frozen=True, slots=True)
class MarkdownBlock:
    """A block's Markdown (:func:`markdown_blocks`)."""

    block: TextBlock
    #: The lines from the start of the line its first character is on to the
    #: next block's, the blank line after it included. Empty for a block that
    #: writes nothing: an empty paragraph, or a code paragraph merged into the
    #: fence before it.
    markdown: str
    #: A code paragraph merged into the fence of a block before it (its
    #: Markdown is that block's). This package's, beside core-ts's two fields.
    merged: bool = False


@dataclass(frozen=True, slots=True)
class MarkdownBlocks:
    """A result's Markdown block by block (:func:`markdown_blocks`)."""

    #: The YAML front matter, under ``metadata: true`` (REQ-704); empty otherwise.
    front_matter: str
    #: Each block of the result, in its order, with its Markdown.
    blocks: tuple[MarkdownBlock, ...]
    #: The whole: the front matter and the blocks' Markdown joined, what :func:`markdown` gives.
    markdown: str


class MarkdownExporter:
    """Exports WordprocessingML to Markdown: CommonMark with GFM's tables and strikethrough, footnotes and math.

    docx4j ``MarkdownExporter``. :meth:`export` takes a package (read under
    ``options.text_options``) or a result the text model has made (written
    under the options it was read with); a result's block quotes are nested
    by their indent beyond the quote style's own where the package it was
    read from is given, by their indent alone where it is not.
    """

    __slots__ = ("options",)

    def __init__(self, options: MarkdownExportOptions | None = None) -> None:
        self.options = options if options is not None else MarkdownExportOptions()

    def export(self, source: Any, package: Any = None) -> str:
        """The Markdown of a package or of a :class:`TextResult`."""
        return self.blocks(source, package).markdown

    def blocks(self, source: Any, package: Any = None) -> MarkdownBlocks:
        """The Markdown of each block, the front matter apart (core-ts CR-012 section 4.4, the editor's ask).

        Joined, the front matter and the blocks' Markdown are :meth:`export`'s.
        """
        if isinstance(source, TextResult):
            result, styles = source, package
        else:
            from docx4j_py.model.text import extract

            result, styles = extract(source, self.options.text_options), source
        writer = _WmlToMarkdown(result, _resolver_of(styles))
        document = writer.convert()
        index = 0

        def next_index() -> int:
            nonlocal index
            index += 1
            return index

        for node, run in writer.images:
            node.destination = self._destination(run, next_index)
        starts: dict[int, int] = {}
        body = md.render(document, starts)
        front = front_matter(result) if result.options.metadata else ""
        return MarkdownBlocks(front, _sliced(result.blocks, starts, body, writer.merged), front + body)

    def _destination(self, run: TextRun, next_index: Callable[[], int]) -> str:
        """A picture's destination: its image part's bytes as a data URI or as ``image_destination`` says.

        Empty where its image cannot be had, so that its alt text is kept.
        """
        part = run.image.part if run.image is not None else None
        data: bytes | None = None
        if part is not None:
            try:
                data = _bytes_of(part)
            except Exception:  # noqa: BLE001 - its alt text, with no destination
                data = None
        if part is None or data is None:
            return ""
        content_type = str(getattr(part, "content_type", "") or "")
        if self.options.image_destination is None:
            return f"data:{content_type};base64,{base64.b64encode(data).decode('ascii')}"
        n = next_index()
        image = MarkdownImage(run, part, data, content_type, n, f"image{n}.{_extension_for(content_type, part)}")
        return self.options.image_destination(image)


def markdown(result: TextResult, package: Any = None, *, image_destination: MarkdownImageDestination | None = None) -> str:
    """A result as Markdown (WD section 7.3, REQ-703, REQ-704). docx4j ``new MarkdownExporter().export(result, pkg)``.

    Args:
        result: what :func:`~docx4j_py.model.text.extract` made; it is
            written under the options it was read with (its view, its
            metadata as YAML front matter under ``metadata: true``).
        package: the package it was read from, for the quote style's indent
            (a block quote's depth); without it the depth is the indent alone.
        image_destination: where each picture points; a data URI where None.
    """
    return markdown_blocks(result, package, image_destination=image_destination).markdown


def markdown_blocks(
    result: TextResult, package: Any = None, *, image_destination: MarkdownImageDestination | None = None
) -> MarkdownBlocks:
    """:func:`markdown` block by block: the front matter apart, each block of the result with its Markdown.

    core-ts ``MarkdownExporter.blocks``. A host that writes something before
    each block (an address), or stops after some (a budget), has the pieces
    :func:`markdown` joins.
    """
    exporter = MarkdownExporter(MarkdownExportOptions(image_destination=image_destination))
    return exporter.blocks(result, package)


def front_matter(result: TextResult) -> str:
    """The document properties as a YAML front-matter block (WD REQ-704, REQ-515). docx4j ``MarkdownExporter.frontMatter``.

    A mapping for each kind of property the result's metadata has, ``core``,
    ``app`` and ``custom`` in that order, each property by its element's
    local name (core, application) or its name (custom), its value a
    double-quoted string. Empty where it has none.
    """
    from docx4j_py.model.text.metadata import shown

    out: list[str] = []
    for kind in (TextResult.CORE, TextResult.APP, TextResult.CUSTOM):
        properties = result.metadata.get(kind)
        if not isinstance(properties, dict) or not properties:
            continue
        out.append(f"{kind}:\n")
        for key, value in properties.items():
            out.append(f"  {_yaml_key(str(key))}: {_yaml_string(shown(value))}\n")
    return "" if not out else "---\n" + "".join(out) + "---\n\n"


# ---------------------------------------------------------------- the exporter's helpers


def _resolver_of(package: Any) -> Any:
    """The resolver a block quote's style indent is asked of: the package's, where it has a styles part."""
    if package is None:
        return None
    try:
        # asking a package with no styles part for a resolver would add one to it
        if getattr(package, "style_definitions_part", None) is None:
            return None
        return package.get_property_resolver()
    except Exception:  # noqa: BLE001 - no styles to ask
        return None


def _bytes_of(part: Any) -> bytes | None:
    """An image part's bytes, or an SVG picture's SVG as stored."""
    data = getattr(part, "data", None)
    if isinstance(data, bytes):
        return data
    if str(getattr(part, "content_type", "")) == _SVG:
        xml = getattr(part, "xml", None)
        if isinstance(xml, bytes):
            return xml
    return None


def _sliced(blocks: list[TextBlock], starts: dict[int, int], body: str, merged: set[int]) -> tuple[MarkdownBlock, ...]:
    """Each block's Markdown: from its start to the next start, the last to the end; a start is one block's."""
    ordered = sorted(set(starts.values()))
    end = {s: (ordered[i + 1] if i + 1 < len(ordered) else len(body)) for i, s in enumerate(ordered)}
    taken: set[int] = set()
    out: list[MarkdownBlock] = []
    for block in blocks:
        start = starts.get(id(block))
        if start is None or start in taken:
            out.append(MarkdownBlock(block, "", id(block) in merged))
            continue
        taken.add(start)
        out.append(MarkdownBlock(block, body[start : end[start]]))
    return tuple(out)


def _yaml_key(key: str) -> str:
    import re

    return key if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.-]*", key) else _yaml_string(key)


def _yaml_string(value: str) -> str:
    """A YAML double-quoted scalar: ``\\``, ``"`` and the control characters escaped (Java's UTF-16 units)."""
    out = ['"']
    for c in value:
        if c == '"':
            out.append('\\"')
        elif c == "\\":
            out.append("\\\\")
        elif c == "\n":
            out.append("\\n")
        elif c == "\r":
            out.append("\\r")
        elif c == "\t":
            out.append("\\t")
        else:
            code = ord(c)
            if code < 0x20 or code in (0x7F, 0x85, 0x2028, 0x2029):
                out.append(f"\\u{code:04x}")
            else:
                out.append(c)
    out.append('"')
    return "".join(out)


def _extension_for(content_type: str, part: Any) -> str:
    known = {
        "image/png": "png",
        "image/jpeg": "jpg",
        "image/gif": "gif",
        "image/bmp": "bmp",
        "image/tiff": "tif",
        "image/svg+xml": "svg",
        "image/x-emf": "emf",
        "image/x-wmf": "wmf",
    }
    if content_type in known:
        return known[content_type]
    name = str(getattr(getattr(part, "part_name", None), "name", "") or getattr(part, "part_name", ""))
    dot = name.rfind(".")
    return name[dot + 1 :] if dot >= 0 else "bin"


# ---------------------------------------------------------------- WmlToMarkdown


@dataclass(slots=True, eq=False)
class _OpenList:
    """An open list: it spans consecutive numbered paragraphs."""

    node: md.ListBlock
    #: The list: its abstractNum where the text model gives one, else its numId.
    num_id: str
    last_item: md.ListItem | None = None
    #: The last item's left indent, where its text starts (twips).
    item_indent: int = 0
    #: The count the next item must show for this list to go on, else a new list starts at it; None for a bullet list.
    next_value: int | None = None


class _WmlToMarkdown:
    """The block model as a Markdown tree (WD REQ-703). docx4j ``WmlToMarkdown``.

    Each node made for a block carries it as its ``tag``, so that
    :func:`markdown_blocks` can say where each block's Markdown begins.
    """

    def __init__(self, result: TextResult, resolver: Any) -> None:
        self.result = result
        self.text_options = result.options
        #: For a style's own indent (a block quote's depth); None where there are no styles to ask.
        self.resolver = resolver
        #: The open lists, the deepest last.
        self.open_lists: list[_OpenList] = []
        #: The count each list level last showed (``numId/ilvl``), for a parent item reopened after an interruption.
        self.last_values: dict[str, int] = {}
        # the open quote: it spans consecutive Quote paragraphs
        self.open_quote: md.BlockQuote | None = None
        self.open_quote_depth = 0
        #: Above 0 while a GFM header row is built, whose bold is convention, not markup.
        self.table_header_depth = 0
        self.style_indents: dict[str, int] = {}
        #: The images, in the order written, for their destinations.
        self.images: list[tuple[md.Image, TextRun]] = []
        #: The ``id`` of each code paragraph merged into the fence of a block before it.
        self.merged: set[int] = set()

    def convert(self) -> md.Document:
        document = md.Document()
        blocks = self.result.blocks
        i = 0
        while i < len(blocks):
            block = blocks[i]
            if _is_note(block):
                # a note's blocks are its definition, at the place the text model reads it (after the main story)
                self._close_lists()
                self._close_quote()
                last = i
                while last + 1 < len(blocks) and _same_note(block, blocks[last + 1]):
                    last += 1
                definition = _tagged(md.FootnoteDefinition(_note_label(block.note_type, block.note_number, block.note_id)), block)
                self._body(blocks[i : last + 1], definition)
                self._close_lists()
                self._close_quote()
                _trim_leading_marker_space(definition)
                document.append_child(definition)
                i = last
            else:
                i = self._block(blocks, i, document)
            i += 1
        self._close_lists()
        self._close_quote()
        return document

    # -- blocks -------------------------------------------------------------------

    def _body(self, blocks: list[TextBlock], container: md.Node) -> None:
        i = 0
        while i < len(blocks):
            i = self._block(blocks, i, container) + 1

    def _block(self, blocks: list[TextBlock], index: int, container: md.Node) -> int:
        """The index of the last block consumed (a code block spans several)."""
        block = blocks[index]
        if block.kind == "TABLE":
            self._close_lists()
            self._close_quote()
            self._table(block, container)
            return index
        if block.kind == "CHART":
            self._close_lists()
            self._close_quote()
            self._chart(block, container)
            return index
        return self._paragraph(blocks, index, container)

    def _paragraph(self, blocks: list[TextBlock], index: int, container: md.Node) -> int:
        block = blocks[index]
        style = block.style

        # a thematic break: an empty paragraph carrying a bottom border of its own
        if block.bottom_border and _is_empty(block):
            self._close_lists()
            self._close_quote()
            container.append_child(_tagged(md.ThematicBreak("---"), block))  # else the renderer picks its own idiom
            return index

        # display math: a paragraph which is one m:oMathPara
        display_math = _sole_display_math(block)
        if display_math is not None:
            self._close_lists()
            self._close_quote()
            container.append_child(_tagged(md.DisplayMath(_latex(display_math)), block))
            return index

        # a code block: a code-styled paragraph, or one whose paragraph style's own font is mono;
        # consecutive ones merged into one fence, one line each
        if _is_code_paragraph(block):
            self._close_lists()
            self._close_quote()
            literal: list[str] = []
            last = index
            for j in range(index, len(blocks)):
                b = blocks[j]
                if b.kind != "PARAGRAPH" or not _is_code_paragraph(b) or not _same_story(block, b):
                    break
                literal.append(self._code_text(b) + "\n")
                last = j
                if j > index:
                    self.merged.add(id(b))
            container.append_child(_tagged(md.FencedCodeBlock("".join(literal)), block))
            return last

        # a heading by its effective outlineLvl (0 to 5 are # to ######); before the numbering, since
        # the built-in Heading styles can carry a legacy numPr
        outline_level = block.outline_level
        if outline_level is not None and outline_level <= 5:
            self._close_lists()
            self._close_quote()
            heading = _tagged(md.Heading(outline_level + 1), block)
            self._inlines(block, heading, True)
            container.append_child(heading)
            return index

        # a list paragraph
        numbering = block.numbering
        if numbering is not None:
            self._close_quote()
            if numbering.label:
                # a level with numFmt none, or a lvlText with no %n, is a label ("NOTE", "EXAMPLE"),
                # not a list: it leads the paragraph in bold
                self._close_lists()
                label = "" if numbering.level_text is None else numbering.level_text.strip()
                labelled = self._inline_paragraph(block, False)
                if label:
                    if labelled.first_child is not None:
                        labelled.prepend_child(md.Text(" "))
                    strong = md.StrongEmphasis()
                    strong.append_child(md.Text(label))
                    labelled.prepend_child(strong)
                if labelled.first_child is not None:
                    container.append_child(labelled)
                return index
            self._list_paragraph(block, numbering, container)
            return index

        # an unlabelled paragraph inside a list (WD REQ-703): Markdown has no empty paragraphs, so an
        # empty one neither ends the list nor joins it; one indented at least to an open item's text (or
        # in the ListParagraph style) is a follow-on paragraph of the deepest such item; any other ends
        # the lists, and the list goes on at its own count where it resumes (_list_paragraph)
        if self.open_lists and style != _QUOTE and style != _INTENSE_QUOTE:
            if _is_empty(block):
                return index
            into = self._follow_on_item(block, style)
            if into is not None:
                into.node.tight = False
                assert into.last_item is not None
                into.last_item.append_child(self._inline_paragraph(block, True))
                return index

        self._close_lists()

        # a block quote, nested by the indent beyond the quote style's own
        if style in (_QUOTE, _INTENSE_QUOTE):
            beyond = _value(block.indent_left) - self._style_indent_left(style)
            depth = 1 + max(0, beyond) // _TWIPS_PER_LEVEL
            self._append_to_quote(depth, self._inline_paragraph(block, True), container, block)
            return index

        self._close_quote()

        paragraph = self._inline_paragraph(block, True)
        if paragraph.first_child is not None:
            container.append_child(paragraph)  # Markdown has no empty paragraphs: dropped
        return index

    def _inline_paragraph(self, block: TextBlock, with_list_label: bool) -> md.Paragraph:
        paragraph = _tagged(md.Paragraph(), block)
        self._inlines(block, paragraph, with_list_label)
        return paragraph

    def _code_text(self, block: TextBlock) -> str:
        """A code paragraph's text, unescaped; in ``markup``, its revisions marked."""
        out: list[str] = []
        runs = block.runs
        i = 0
        while i < len(runs):
            revision = runs[i].revision
            j = i
            while j < len(runs) and self._same_marker(revision, runs[j].revision):
                j += 1
            if revision is not None:
                out.append(self._open(revision))
            for k in range(i, j):
                r = runs[k]
                if r.kind in ("TEXT", "FIELD_CODE"):
                    out.append(r.text)
            if revision is not None:
                out.append(_close(revision))
            i = j
        return "".join(out)

    def _style_indent_left(self, style_id: str) -> int:
        """A paragraph style's own left indent, in twips: 0 where it has none, or there are no styles to ask."""
        cached = self.style_indents.get(style_id)
        if cached is not None:
            return cached
        left = 0
        if self.resolver is not None:
            try:
                effective = self.resolver.get_effective_ppr(style_id)
                ind = getattr(effective, "ind", None) if effective is not None else None
                if ind is not None:
                    value = ind.left if ind.left is not None else ind.start
                    left = integer(value) or 0  # as Word reads it (CR-021); a refused form is unset
            except Exception:  # noqa: BLE001 - no indent for the style
                left = 0
        self.style_indents[style_id] = left
        return left

    # -- quotes -------------------------------------------------------------------

    def _append_to_quote(self, depth: int, paragraph: md.Paragraph, container: md.Node, block: TextBlock) -> None:
        if self.open_quote is None or depth != self.open_quote_depth:
            self._close_quote()
            outer = _tagged(md.BlockQuote(), block)
            inner = outer
            for _ in range(1, depth):
                nested = md.BlockQuote()
                inner.append_child(nested)
                inner = nested
            container.append_child(outer)
            self.open_quote = inner
            self.open_quote_depth = depth
        self.open_quote.append_child(paragraph)

    def _close_quote(self) -> None:
        self.open_quote = None
        self.open_quote_depth = 0

    # -- lists --------------------------------------------------------------------

    def _follow_on_item(self, block: TextBlock, style: str | None) -> _OpenList | None:
        """The open list whose last item an unlabelled paragraph continues, the lists deeper than it closed.

        The deepest whose item's text the paragraph is indented to (its left
        indent at least the item's, and not at the margin), or the deepest open
        list for a ListParagraph-styled paragraph; None where none: the
        paragraph ends the lists.
        """
        indent = _value(block.indent_left)
        list_paragraph = style == _LIST_PARAGRAPH
        for k in range(len(self.open_lists) - 1, -1, -1):  # deepest first
            open_list = self.open_lists[k]
            if open_list.last_item is not None and (list_paragraph or (indent > 0 and indent >= open_list.item_indent)):
                del self.open_lists[k + 1 :]
                return open_list
        return None

    def _list_paragraph(self, block: TextBlock, numbering: TextNumbering, container: md.Node) -> None:
        ilvl = numbering.level
        # the list: two w:num on one w:abstractNum are one list, counting on from each other
        num_id = numbering.abstract_num_id if numbering.abstract_num_id is not None else numbering.num_id
        count = numbering.value
        bullet = numbering.bullet or numbering.format is None
        lists = self.open_lists

        # close any deeper (or same-level-but-different-list) contexts
        del lists[ilvl + 1 :]
        if len(lists) == ilvl + 1 and lists[-1].num_id != num_id and ilvl == 0:
            lists.pop()  # a different top-level list: restart
        if len(lists) == ilvl + 1 and isinstance(lists[-1].node, md.BulletList) != bullet:
            lists.pop()  # a level drawn as the other kind (a w:num overriding it): a list of that kind

        tight = block.contextual_spacing is not False

        # the list's items show consecutive counts from its start; an item whose count is not the next one
        # (the list resumed after a paragraph that ended it, a w:num sharing the count, a start override)
        # begins a new list at its count
        if len(lists) == ilvl + 1 and count is not None and lists[-1].next_value is not None and count != lists[-1].next_value:
            lists.pop()

        # open the missing levels (a level skipped over takes the paragraph's kind of list); a parent level
        # reopened after an interruption is an empty item showing the parent's last count
        while len(lists) < ilvl + 1:
            level = len(lists)
            list_block = _tagged(
                _create_list_block(numbering, level == ilvl, tight, count if level == ilvl else self.last_values.get(f"{num_id}/{level}")),
                block,
            )
            if not lists:
                _append_list(container, list_block)
            else:
                parent = lists[-1]
                if parent.last_item is None:
                    parent.last_item = _tagged(md.ListItem(), block)
                    parent.node.append_child(parent.last_item)
                _append_list(parent.last_item, list_block)
            lists.append(_OpenList(list_block, num_id))

        target = lists[-1]
        if not tight:
            target.node.tight = False
        item = _tagged(md.ListItem(), block)
        item.append_child(self._inline_paragraph(block, False))
        target.node.append_child(item)
        target.last_item = item
        target.item_indent = _value(block.indent_left)
        if isinstance(target.node, md.OrderedList) and count is not None:
            target.next_value = count + 1
            self.last_values[f"{num_id}/{ilvl}"] = count

    def _close_lists(self) -> None:
        self.open_lists.clear()

    # -- tables -------------------------------------------------------------------

    def _table(self, table: TextBlock, container: md.Node) -> None:
        """GFM pipes: the header the ``w:tblHeader`` rows, else the first; a cell's paragraphs ``<br>`` apart.

        Several ``w:tblHeader`` rows are one header, their cells' texts `` / ``
        apart, as plain text's listed format joins them; ``w:vMerge`` and
        ``w:gridSpan`` degrade (top-left wins, empty cells elsewhere), GFM
        having no spans. The table style's formatting is left out on purpose
        (docx4j CR-030 decision 4): a header row the style makes bold is not
        written ``**bold**``.
        """
        table_block = _tagged(md.TableBlock(), table)
        head = md.TableHead()
        body = md.TableBody()
        table_block.append_child(head)

        rows = table.rows
        header_rows = 0
        while header_rows < len(rows) and rows[header_rows].header:
            header_rows += 1
        if header_rows == 0 and rows:
            header_rows = 1

        if header_rows > 0:
            head.append_child(self._header_row(rows[:header_rows]))
        for r in range(header_rows, len(rows)):
            row = md.TableRow()
            for cell in rows[r].cells:
                row.append_child(self._table_cell(cell, False))
                # gridSpan: empty cells after it, so that the columns stay aligned
                for _ in range(1, cell.grid_span):
                    row.append_child(md.TableCell(False))
            body.append_child(row)
        if body.first_child is not None:
            table_block.append_child(body)
        container.append_child(table_block)

    def _header_row(self, rows: list[TextRow]) -> md.TableRow:
        """The header: one row's cells, or several rows' cells by grid column, `` / `` apart."""
        row = md.TableRow()
        if len(rows) == 1:
            for cell in rows[0].cells:
                row.append_child(self._table_cell(cell, True))
                for _ in range(1, cell.grid_span):
                    row.append_child(md.TableCell(True))
            return row
        # each grid column's cells, from the top: a cell is in the column it starts in
        width = 0
        for r in rows:
            width = max(width, sum(max(1, c.grid_span) for c in r.cells))
        columns: list[list[TextCell]] = [[] for _ in range(width)]
        for r in rows:
            g = 0
            for c in r.cells:
                if not c.merge_continue:
                    columns[g].append(c)
                g += max(1, c.grid_span)
        for column in columns:
            cell = md.TableCell(True)
            # the column's alignment is its topmost header cell's, as with one header row
            alignment: str | None = None
            if column:
                first_paragraphs = _paragraphs_of(column[0])
                if first_paragraphs:
                    alignment = _cell_alignment(first_paragraphs[0])
            sink = _InlineSink(cell)
            self.table_header_depth += 1
            try:
                first = True
                for c in column:
                    if _is_empty_cell(c):
                        continue
                    if not first:
                        sink.text(" / ", False, False, False)
                    first = False
                    self._cell_content(c, cell, sink)
            finally:
                self.table_header_depth -= 1
            sink.flush()
            cell.alignment = alignment
            row.append_child(cell)
        return row

    def _table_cell(self, text_cell: TextCell, header: bool) -> md.TableCell:
        cell = md.TableCell(header)
        # a vMerge continuation: the merge's top cell already holds the content
        if text_cell.merge_continue:
            return cell
        sink = _InlineSink(cell)
        if header:
            self.table_header_depth += 1
        try:
            self._cell_content(text_cell, cell, sink)
        finally:
            if header:
                self.table_header_depth -= 1
        sink.flush()
        return cell

    def _cell_content(self, text_cell: TextCell, cell: md.TableCell, sink: _InlineSink) -> None:
        """A cell's paragraphs, a nested table's included, ``<br>`` apart; its empty paragraphs left out.

        A cell's alignment is its first paragraph's effective ``w:jc``, empty
        or not; GFM aligns a column, from its header cell's, so a body cell's
        is not seen.
        """
        first = True
        paragraphs = _paragraphs_of(text_cell)
        if cell.alignment is None and paragraphs:
            cell.alignment = _cell_alignment(paragraphs[0])
        for b in paragraphs:
            if _is_empty(b):
                continue
            if not first:
                sink.node(md.HtmlInline("<br>"))
            first = False
            self._runs(b, sink)

    # -- charts -------------------------------------------------------------------

    def _chart(self, block: TextBlock, container: md.Node) -> None:
        """A chart: its note (docx4j CR-047), its title and its data as a table, the categories down, the series across."""
        if block.chart is not None:
            note = _tagged(md.Paragraph(), block)
            note.append_child(md.Text(block.chart.note))
            container.append_child(note)
        if block.chart_title:
            title = _tagged(md.Paragraph(), block)
            title.append_child(md.Text(block.chart_title))
            container.append_child(title)
        if not block.chart_header:
            return
        table = _tagged(md.TableBlock(), block)
        head = md.TableHead()
        head.append_child(_chart_row(block.chart_header, True, block.chart_header))
        table.append_child(head)
        if block.chart_rows:
            body = md.TableBody()
            for r in block.chart_rows:
                body.append_child(_chart_row(r, False, block.chart_header))
            table.append_child(body)
        container.append_child(table)

    # -- inlines ------------------------------------------------------------------

    def _inlines(self, block: TextBlock, parent: md.Node, with_list_label: bool) -> None:
        sink = _InlineSink(parent)
        # under list_labels: true, the label the numbering model computes leads a paragraph that isn't
        # written as a list item (a numbered heading's "1.2")
        if with_list_label and block.list_label is not None and block.list_label.strip():
            sink.text(block.list_label.replace("\t", " "), False, False, False)
        self._runs(block, sink)
        sink.flush()

    def _runs(self, block: TextBlock, sink: _InlineSink) -> None:
        """A block's runs, in ``markup`` each revision's between its CriticMarkup marks (REQ-703, REQ-702)."""
        runs = block.runs
        i = 0
        while i < len(runs):
            revision = runs[i].revision
            j = i
            while j < len(runs) and self._same_marker(revision, runs[j].revision):
                j += 1
            if revision is not None:
                sink.node(md.HtmlInline(self._open(revision)))
            self._linked(runs[i:j], sink)
            if revision is not None:
                sink.node(md.HtmlInline(_close(revision)))
            i = j

    def _linked(self, runs: list[TextRun], sink: _InlineSink) -> None:
        """Runs, those with the same link target in one link."""
        i = 0
        while i < len(runs):
            target = runs[i].link
            j = i
            while j < len(runs) and runs[j].link == target:
                j += 1
            if target is None:
                for k in range(i, j):
                    self._run(runs[k], sink)
            else:
                link = md.Link(target, None)
                link_sink = _InlineSink(link)
                for k in range(i, j):
                    self._run(runs[k], link_sink)
                link_sink.flush()
                sink.node(link)
            i = j

    def _run(self, run: TextRun, sink: _InlineSink) -> None:
        kind = run.kind
        if kind == "TEXT":
            self._text(run, sink)
        elif kind in ("FIELD_CODE", "LIST_LABEL"):
            sink.text(run.text, False, False, False)
        elif kind == "NOTE_REFERENCE":
            if self.text_options.notes and self.text_options.note_references:
                sink.node(md.FootnoteReference(_note_label(run.note_type, run.note_number, run.note_id)))
        elif kind == "COMMENT_REFERENCE":
            comment = self.result.comments.get(run.comment_id) if run.comment_id is not None else None
            # a reply is written after its parent, never at its own reference (REQ-504)
            if comment is None or (comment.parent_id is not None and comment.parent_id in self.result.comments):
                return
            self._comment(comment, sink)
            for reply in comment.replies:
                self._comment(reply, sink)
        elif kind == "IMAGE":
            self._image(run, sink)
        elif kind == "PAGE_MARKER" and self.text_options.page_markers:
            sink.text(f"[page {run.page_number}]", False, False, False)
        # else a note's own mark (its definition's label carries it), an object (no text)

    def _text(self, run: TextRun, sink: _InlineSink) -> None:
        value = run.text
        if not value:
            return
        if run.math:
            latex = _latex(run)
            if latex:
                sink.node(md.InlineMath(latex, False))
            return
        bold = run.bold and self.table_header_depth == 0  # a header row's bold is convention, not markup
        lines = value.split("\n")
        for n, line in enumerate(lines):
            if n > 0:
                sink.node(md.HardLineBreak())
            if not line:
                continue
            if run.monospace:
                sink.node(md.Code(line))
            else:
                sink.text(line, bold, run.italic, run.strike)

    def _comment(self, comment: TextComment, sink: _InlineSink) -> None:
        sink.node(md.HtmlInline("{>>" + self._detail(comment.author, comment.date)))
        sink.text(comment.text, False, False, False)
        sink.node(md.HtmlInline("<<}"))

    def _same_marker(self, a: TextRevision | None, b: TextRevision | None) -> bool:
        """Runs under one marker: the same type, and the same author and date where the detail shows them."""
        if a is None or b is None:
            return a is b
        if a.type != b.type:
            return False
        detail = self.text_options.markup_detail
        if detail == "author":
            return a.author == b.author
        if detail == "author_date":
            return a.author == b.author and a.date == b.date
        return True

    def _open(self, revision: TextRevision) -> str:
        return ("{++" if revision.type == "INSERTION" else "{--") + self._detail(revision.author, revision.date)

    def _detail(self, author: str | None, date: str | None) -> str:
        detail = self.text_options.markup_detail
        if detail == "author":
            return f"{author or ''}: "
        if detail == "author_date":
            return f"{author or ''}, {date or ''}: "
        return ""

    def _image(self, run: TextRun, sink: _InlineSink) -> None:
        """A picture with the describer's text (REQ-703, REQ-514); its destination is the exporter's to fill in."""
        image = md.Image("", None)
        image.append_child(md.Text(run.description.text if run.description is not None else ""))
        self.images.append((image, run))
        sink.node(image)


# ---------------------------------------------------------------- helpers


def _tagged(node: Any, block: TextBlock) -> Any:
    """A node made for a block: it begins the block's Markdown (:func:`markdown_blocks`)."""
    node.tag = block
    return node


class _InlineSink:
    """Accumulates same-formatted text, wrapping it on flush."""

    __slots__ = ("bold", "buf", "italic", "parent", "strike")

    def __init__(self, parent: md.Node) -> None:
        self.parent = parent
        self.buf: list[str] = []
        self.bold = False
        self.italic = False
        self.strike = False

    def text(self, s: str, b: bool, i: bool, st: bool) -> None:
        if (b != self.bold or i != self.italic or st != self.strike) and self.buf:
            self.flush()
        self.bold = b
        self.italic = i
        self.strike = st
        if s:
            self.buf.append(s)

    def node(self, n: md.Node) -> None:
        self.flush()
        self.parent.append_child(n)

    def flush(self) -> None:
        if not self.buf:
            return
        n: md.Node = md.Text("".join(self.buf))
        # strong inside emphasis: the renderer then writes ***both***, rather than switching the inner marker to _
        if self.bold:
            strong = md.StrongEmphasis()
            strong.append_child(n)
            n = strong
        if self.italic:
            em = md.Emphasis()
            em.append_child(n)
            n = em
        if self.strike:
            st = md.Strikethrough("~~")
            st.append_child(n)
            n = st
        self.parent.append_child(n)
        self.buf = []


def _value(i: int | None) -> int:
    return 0 if i is None else i


def _is_empty(block: TextBlock) -> bool:
    for r in block.runs:
        if r.kind == "TEXT" and r.text:
            return False
        if r.kind in ("IMAGE", "NOTE_REFERENCE"):
            return False
    return True


def _is_empty_cell(cell: TextCell) -> bool:
    for b in cell.blocks:
        if b.kind == "TABLE":
            for r in b.rows:
                for c in r.cells:
                    if not _is_empty_cell(c):
                        return False
        elif b.kind == "PARAGRAPH" and not _is_empty(b):
            return False
    return True


def _sole_display_math(block: TextBlock) -> TextRun | None:
    """The paragraph's one run, a display equation with no revision; else None."""
    found: TextRun | None = None
    for r in block.runs:
        if r.kind == "TEXT" and not r.text:
            continue
        if found is not None or not r.display_math or r.revision is not None:
            return None
        found = r
    return found


def _latex(run: TextRun) -> str:
    """An equation's LaTeX: its text without the delimiters."""
    s = run.text
    d = 2 if run.display_math else 1
    return s[d : len(s) - d] if len(s) >= 2 * d else s


def _is_code_paragraph(block: TextBlock) -> bool:
    """A code paragraph: styled as one, or in a named paragraph style whose own font is mono.

    Every run monospaced, none of them beyond the style. Inline code is what
    a run adds to its paragraph style, so a paragraph whose style is already
    mono would get no code marks at all; this is the paragraph-level
    counterpart. A mono document default does not count (a document set in
    Courier New throughout is prose, not one fence), which is why a named
    style is required.
    """
    style = block.style
    if style is None:
        return False
    if style in _CODE_STYLES:
        return True
    if not block.monospace:
        return False
    return not any(r.kind == "TEXT" and r.monospace for r in block.runs)


def _same_story(a: TextBlock, b: TextBlock) -> bool:
    return a.story == b.story and a.note_id == b.note_id


def _create_list_block(numbering: TextNumbering, own: bool, tight: bool, count: int | None) -> md.ListBlock:
    """A list block: an ordered one starts at the count given, else (its own level at the top) at the level's start."""
    # a level the numbering definitions don't give is a bullet list
    list_block: md.ListBlock
    if numbering.bullet or numbering.format is None:
        list_block = md.BulletList()
    else:
        ordered = md.OrderedList()
        if count is not None:
            if count != 1:
                ordered.marker_start_number = count
        elif own and numbering.level == 0 and numbering.start > 1:
            ordered.marker_start_number = numbering.start
        list_block = ordered
    list_block.tight = tight
    return list_block


def _append_list(parent: md.Node, list_block: md.ListBlock) -> None:
    """Appends a list; one that directly follows another of its kind takes the other delimiter or bullet.

    Markdown runs two adjacent lists of one kind together, numbering the
    second on from the first; ``1.`` against ``1)`` and ``-`` against ``*``
    are read as a new list.
    """
    previous = parent.last_child
    if isinstance(list_block, md.OrderedList) and isinstance(previous, md.OrderedList):
        list_block.marker_delimiter = "." if previous.marker_delimiter == ")" else ")"
    elif isinstance(list_block, md.BulletList) and isinstance(previous, md.BulletList):
        list_block.marker = "-" if previous.marker == "*" else "*"
    parent.append_child(list_block)


def _paragraphs_of(cell: TextCell) -> list[TextBlock]:
    out: list[TextBlock] = []
    for b in cell.blocks:
        if b.kind == "TABLE":
            for r in b.rows:
                for c in r.cells:
                    out.extend(_paragraphs_of(c))
        elif b.kind == "PARAGRAPH":
            out.append(b)
    return out


def _cell_alignment(block: TextBlock) -> str | None:
    if block.alignment == "center":
        return "CENTER"
    if block.alignment in ("right", "end"):
        return "RIGHT"
    return None  # LEFT is GFM's default


def _chart_row(cells: list[str], header: bool, headings: list[str]) -> md.TableRow:
    row = md.TableRow()
    labels = 0
    while labels < len(headings) and headings[labels] == "":
        labels += 1
    for i, text in enumerate(cells):
        cell = md.TableCell(header)
        if i >= max(labels, 1):
            cell.alignment = "RIGHT"
        if text:
            cell.append_child(md.Text(text))
        row.append_child(cell)
    return row


def _is_note(block: TextBlock) -> bool:
    return block.note_type is not None and block.story in ("FOOTNOTE", "ENDNOTE", "TEXT_BOX")


def _same_note(a: TextBlock, b: TextBlock) -> bool:
    return _is_note(b) and a.note_type == b.note_type and a.note_id == b.note_id


def _note_label(note_type: str | None, number: int, note_id: str | None) -> str:
    """A note's label: a footnote's number (``[^1]``), an endnote's prefixed ``e`` (``[^e1]``).

    The two are numbered apart (WD REQ-503); a note whose reference carries a
    custom mark has no number, so its ``w:id`` labels it (``[^f-5]``,
    ``[^e-5]``).
    """
    foot = note_type == "footnote"
    if number > 0:
        return str(number) if foot else f"e{number}"
    return ("f-" if foot else "e-") + str(note_id)


def _trim_leading_marker_space(definition: md.FootnoteDefinition) -> None:
    """Drops the leading space left behind by the note's own mark."""
    paragraph = definition.first_child
    if isinstance(paragraph, md.Paragraph) and isinstance(paragraph.first_child, md.Text):
        text = paragraph.first_child
        text.literal = text.literal.removeprefix(" ")


def _close(revision: TextRevision) -> str:
    return "++}" if revision.type == "INSERTION" else "--}"
