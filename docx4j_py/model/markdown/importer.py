"""Markdown *in*: CommonMark plus GFM tables and strikethrough, as WordprocessingML.

CR-003 section 3.5, Phase K. The behavioural oracle is Java docx4j's
``docx4j-markdown`` (``MarkdownToWmlVisitor``, ``ImportStyles``,
``ImportNumbering``): the same style ids, the same numbering shape, the same
fallbacks. What is different is the parser --- ``markdown-it-py`` rather than
commonmark-java, so this walks a **token stream** rather than an AST --- and
the posture on anything that would add a dependency (section 3.5's limits are
in CR-003 section 13).

``markdown_it`` is imported **inside** :func:`parser`, on the first
``insert_markdown`` of the process, so that ``import docx4j_py`` does not pay
its 30 ms and ``to_markdown`` never pays it at all (CR-003 decided question 6).

What this may touch beyond the body's own part:

* ``/word/styles.xml`` --- a style the markdown needs and the document does not
  define is added from docx4j's ``KnownStyles.xml``, or, for the two code
  styles Word has no built-in equivalent of, from the minimal definitions Java
  ``ImportStyles`` writes. An existing definition always wins. The machinery is
  :func:`~docx4j_py.model.content.styles.ensure_style`, which lives in
  :mod:`docx4j_py.model.content.styles` with the rest of the style rule (CR-003
  section 14.9) and is re-exported here, where Phase K built it.
* ``/word/numbering.xml`` --- a list needs a numbering definition, and the part
  is **created** when the document has none, through
  :func:`~docx4j_py.model.content.lists.numbering_part_of`, which is the one
  place any of this package creates one (CR-003 section 13.6). Since Phase H it
  goes through the trial's undo log, so a
  :func:`~docx4j_py.model.content.trial.dry_run` of a markdown fragment with a
  list leaves **nothing** behind --- the limitation CR-003 sections 12.5 and
  14.4 recorded is closed.
* the body's part's relationships --- one external relationship per link.

Every id is allocated from the document's own state (the next free
``w:numId``, the next free ``w:abstractNumId``, the relationships part's own
``rIdN``), so the same markdown inserted into the same document twice gives
the same bytes, which is CR-003 section 3.4's determinism.
"""

from __future__ import annotations

import dataclasses
from pathlib import Path
from typing import Any

from docx4j_py.child import ChildList, link_parents
from docx4j_py.model.content.errors import ContentError, StyleError
from docx4j_py.model.content.lists import (
    next_abstract_num_id,
    next_num_id,
    numbering_part_of,
)
from docx4j_py.model.content.styles import CUSTOM_STYLE_XML, ensure_style, style_ids_of
from docx4j_py.model.content.table import writable_width
from docx4j_py.model.listnumbering import invalidate
from docx4j_py.wml import CtLvlStart, P, R, RPr, el, t, tbl, tc, tr

__all__ = [
    "CUSTOM_STYLE_XML",
    "STYLE_IDS",
    "blocks_for",
    "ensure_style",
    "parser",
    "style_ids_of",
]

#: The style ids the mapping uses, by role. docx4j ``ImportStyles``.
STYLE_IDS: dict[str, str] = {
    "heading": "Heading",  # plus the level: Heading1 ... Heading6
    "quote": "Quote",
    "hyperlink": "Hyperlink",
    "list": "ListParagraph",
    "code_char": "CodeChar",
    "source_code": "SourceCode",
    "table": "TableGrid",
}

#: Twips of indent per list or quote level, and the hanging indent of a marker.
TWIPS_PER_LEVEL = 720
HANGING = 360

#: Word's default bullet glyph and font cycle (Java ``ImportNumbering``).
BULLET_CHARS = ("", "o", "")
BULLET_FONTS = ("Symbol", "Courier New", "Wingdings")

#: How many levels an abstract numbering definition carries.
LEVELS = 9

_HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


# ---------------------------------------------------------------------------
# numbering
# ---------------------------------------------------------------------------


class _Numbering:
    """Real numbering definitions for imported lists. docx4j ``ImportNumbering``.

    Each top-level list is scanned for the format at each depth, giving a
    nine-character signature from which one ``w:abstractNum`` is built, so that
    ordered-inside-bullet and the other way round both get the right markers.
    A bullet-only signature is reused; anything with an ordered level gets its
    own definition, so that two ordered lists both restart.
    """

    def __init__(self, package: Any, touched: set[str]) -> None:
        """Find or create the numbering part, and prepare the id allocators."""
        self.package = package
        self.touched = touched
        # the part, its creation and the two id allocators are Phase H's
        # (``model/content/lists.py``), so that a definition this makes and one
        # ``start_new_list`` makes are made the same way and numbered from the
        # same state --- CR-003 section 13.6 asked for exactly that, and it is
        # what lets a dry run un-add a numbering part it created (section 14.4)
        self.part = numbering_part_of(package, create=True, touched=touched)
        self.by_signature: dict[str, int] = {}
        numbering = self.part.contents
        self.numbering = numbering
        self._next_abstract = next_abstract_num_id(numbering)
        self._next_num = next_num_id(numbering)

    def num_id_for(self, signature: list[str], start: int) -> int:
        """The ``w:numId`` for a top-level list of that signature."""
        key = "".join("b" if fmt == "bullet" else "d" for fmt in signature)
        if key == "b" * LEVELS:
            found = self.by_signature.get(key)
            if found is not None:
                return found
            found = self._create(signature, 1)
            self.by_signature[key] = found
            return found
        return self._create(signature, start)

    def _create(self, signature: list[str], start: int) -> int:
        abstract_id = self._next_abstract
        self._next_abstract += 1
        num_id = self._next_num
        self._next_num += 1

        abstract = el.abstractNum(
            abstract_num_id=abstract_id,
            multi_level_type=el.multiLevelType(val="hybridMultilevel"),
        )
        bullets = 0
        for index in range(LEVELS):
            fmt = signature[index]
            level = el.lvl(
                ilvl=index,
                # ``el.start`` is ``w:tblCellMar/w:start``; the ``w:lvl`` one is
                # ``CtLvlStart``, which has no element name of its own
                start=CtLvlStart(val=start if index == 0 else 1),
                num_fmt=el.numFmt(val=fmt),
                lvl_jc=el.lvlJc(val="left"),
                p_pr=el.pPr(ind=el.ind(left=TWIPS_PER_LEVEL * (index + 1), hanging=HANGING)),
            )
            if fmt == "bullet":
                cycle = bullets % len(BULLET_CHARS)
                bullets += 1
                level.lvl_text = el.lvlText(val=BULLET_CHARS[cycle])
                level.r_pr = el.rPr(
                    r_fonts=el.rFonts(ascii=BULLET_FONTS[cycle], h_ansi=BULLET_FONTS[cycle])
                )
            else:
                level.lvl_text = el.lvlText(val=f"%{index + 1}.")
            abstract.lvl.append(level)

        number = el.num(num_id=num_id, abstract_num_id=el.abstractNumId(val=abstract_id))
        self.numbering.abstract_num.append(abstract)
        self.numbering.num.append(number)
        link_parents(self.numbering)
        # the definitions have changed: the emulator reads them again before the
        # next label (CR-003 Phase H)
        invalidate(self.package)
        return num_id


# ---------------------------------------------------------------------------
# the token walk
# ---------------------------------------------------------------------------


@dataclasses.dataclass(slots=True)
class _ListContext:
    """One open markdown list level: which numbering, which ``w:ilvl``."""

    num_id: int
    ilvl: int
    tight: bool
    item_pending: bool = False


class _Importer:
    """Walks markdown-it's token stream and builds block-level elements."""

    def __init__(self, body: Any, recorder: Any) -> None:
        """Prepare the importer over a body's part and package."""
        self.body = body
        self.part = getattr(body, "part", None)
        self.package = getattr(body, "package", None)
        self.recorder = recorder
        self.results: list[Any] = []
        self.touched: set[str] = set()
        self._defined_cache: set[str] | None = None
        self._styles_cache: list[Any] | None = None
        self.numbering: _Numbering | None = None
        self.current_p: P | None = None
        self.current_hyperlink: Any = None
        self.bold = 0
        self.italic = 0
        self.strike = 0
        self.quote_depth = 0
        self.lists: list[_ListContext] = []

    # -- styles ------------------------------------------------------------

    def style(self, style_id: str) -> str:
        """Ensure a style and report the id, warning rather than failing."""
        try:
            return ensure_style(
                self.package, style_id, touched=self.touched, defined=self._defined()
            )
        except StyleError as error:
            self.recorder.warn(f"{error.message}; the paragraph is unstyled")
            return ""

    def chosen(self, role: str, level: int | None = None) -> str:
        """The style id for a role, from what the document defines.

        ``describe()``'s style list is what section 3.5 says to choose from,
        and the filter is **every style the document defines**, not
        ``in_use=True``: a styles template defines ``Heading 1`` without having
        used it, and refusing to use it would be exactly wrong (CR-003 section
        13).
        """
        wanted = STYLE_IDS[role] + (str(level) if level is not None else "")
        defined = self._defined()
        if wanted in defined:
            return wanted
        # a document whose styles part spells the id differently: match on the
        # display name, which is the vocabulary ``describe()`` reports in
        from docx4j_py.model.content.styles import display_name_of

        target = display_name_of(wanted).lower()
        for info in self._description_styles():
            if info.name.lower() == target:
                return info.id
        return self.style(wanted)

    def _defined(self) -> set[str]:
        """The style ids the document defines, read once for the whole fragment."""
        if self._defined_cache is None:
            part = getattr(self.package, "style_definitions_part", None)
            self._defined_cache = style_ids_of(part) if part is not None else set()
        return self._defined_cache

    def _description_styles(self) -> list[Any]:
        """``describe()``'s style list, which section 3.5 says to choose from."""
        if self._styles_cache is None:
            describe = getattr(self.package, "describe", None)
            self._styles_cache = list(describe().styles) if describe is not None else []
        return self._styles_cache

    # -- the walk ----------------------------------------------------------

    def run(self, tokens: list) -> list[Any]:
        """Every block the tokens describe, in document order."""
        self.blocks(tokens, 0, len(tokens))
        return self.results

    def blocks(self, tokens: list, start: int, stop: int) -> None:
        """The block-level tokens in ``[start, stop)``."""
        index = start
        while index < stop:
            token = tokens[index]
            kind = token.type
            if kind == "heading_open":
                level = min(int(token.tag[1:]), 6)
                close = _closing(tokens, index, "heading_close")
                self.new_paragraph(_p_pr(self.chosen("heading", level)))
                self.inlines(tokens, index + 1, close)
                self.current_p = None
                index = close
            elif kind == "paragraph_open":
                close = _closing(tokens, index, "paragraph_close")
                self.new_paragraph(self.context_p_pr())
                self.inlines(tokens, index + 1, close)
                self.current_p = None
                index = close
            elif kind in ("bullet_list_open", "ordered_list_open"):
                close = _closing(tokens, index, kind.replace("_open", "_close"))
                self.list(tokens, index, close)
                index = close
            elif kind == "blockquote_open":
                close = _closing(tokens, index, "blockquote_close")
                self.quote_depth += 1
                self.blocks(tokens, index + 1, close)
                self.quote_depth -= 1
                index = close
            elif kind in ("fence", "code_block"):
                self.code_block(token.content)
            elif kind == "hr":
                self.thematic_break()
            elif kind == "table_open":
                close = _closing(tokens, index, "table_close")
                self.table(tokens, index + 1, close)
                index = close
            elif kind == "html_block":
                first = token.content.strip().splitlines()[0] if token.content.strip() else ""
                self.recorder.warn(f"an HTML block was skipped: {first[:60]}")
            elif kind == "inline":
                # a stray inline (a list item with no paragraph token)
                self.new_paragraph(self.context_p_pr())
                self.inline_children(token.children or [])
                self.current_p = None
            index += 1

    # -- lists -------------------------------------------------------------

    def list(self, tokens: list, start: int, stop: int) -> None:
        """One markdown list, and the items and lists nested in it."""
        token = tokens[start]
        ordered = token.type == "ordered_list_open"
        if self.lists:
            parent = self.lists[-1]
            num_id = parent.num_id
            ilvl = min(parent.ilvl + 1, LEVELS - 1)
        else:
            signature = _signature(tokens, start, stop)
            start_number = 1
            if ordered:
                try:
                    start_number = max(1, int(token.attrs.get("start", 1)))
                except (TypeError, ValueError):  # pragma: no cover - malformed start
                    start_number = 1
            num_id = self.numbering_for().num_id_for(signature, start_number)
            ilvl = 0
        tight = _is_tight(tokens, start, stop)
        self.lists.append(_ListContext(num_id, ilvl, tight))
        index = start + 1
        while index < stop:
            if tokens[index].type == "list_item_open":
                close = _closing(tokens, index, "list_item_close")
                self.lists[-1].item_pending = True
                self.blocks(tokens, index + 1, close)
                index = close
            index += 1
        self.lists.pop()

    def numbering_for(self) -> _Numbering:
        """The numbering helper, made (and the part created) on first use."""
        if self.numbering is None:
            if self.package is None:
                raise ContentError(
                    "a markdown list needs a numbering part, and this body has no package",
                    code="markdown.no_package",
                    hint="insert into a body that belongs to a loaded or created package",
                )
            self.numbering = _Numbering(self.package, self.touched)
        return self.numbering

    # -- blocks ------------------------------------------------------------

    def code_block(self, literal: str) -> None:
        """A fenced or indented code block: one ``SourceCode`` paragraph."""
        style_id = self.chosen("source_code")
        lines = literal.replace("\r", "").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        self.new_paragraph(_p_pr(style_id))
        for position, line in enumerate(lines):
            if position:
                self.current_p.content.append(R(content=ChildList([el.br()])))
            run = R(content=ChildList([t(line)]))
            self.current_p.content.append(run)
        self.current_p = None

    def thematic_break(self) -> None:
        """``---``: an empty paragraph with a bottom border, as Java writes it."""
        border = el.bottom(val="single", sz=6, space=1, color="auto")
        self.new_paragraph(el.pPr(p_bdr=el.pBdr(bottom=border)))
        self.current_p = None

    def table(self, tokens: list, start: int, stop: int) -> None:
        """A GFM pipe table as a ``w:tbl`` with a grid sized from ``w:sectPr``."""
        rows: list[list[Any]] = []
        alignments: list[str] = []
        header = False
        index = start
        while index < stop:
            token = tokens[index]
            if token.type == "thead_open":
                header = True
            elif token.type == "thead_close":
                header = False
            elif token.type == "tr_open":
                close = _closing(tokens, index, "tr_close")
                cells: list[Any] = []
                inner = index + 1
                while inner < close:
                    cell_token = tokens[inner]
                    if cell_token.type in ("th_open", "td_open"):
                        cell_close = _closing(
                            tokens, inner, cell_token.type.replace("_open", "_close")
                        )
                        if not rows:
                            alignments.append(_alignment(cell_token))
                        cells.append(self.table_cell(tokens, inner + 1, cell_close, header))
                        inner = cell_close
                    inner += 1
                rows.append(cells)
                index = close
            index += 1
        if not rows:
            self.recorder.warn("an empty markdown table was skipped")
            return
        columns = max(len(row) for row in rows)
        total = self.writable_width()
        each = total // columns
        widths = [
            (total - each * (columns - 1)) if i == columns - 1 else each for i in range(columns)
        ]
        table = tbl([[] for _ in rows], widths=widths, style=self.chosen("table") or None)
        table.content.clear()
        for position, cells in enumerate(rows):
            padded = list(cells) + [tc("") for _ in range(columns - len(cells))]
            for cell_index, cell in enumerate(padded):
                cell.tc_pr = el.tcPr(tc_w=el.tcW(w=widths[cell_index], type_value="dxa"))
                alignment = alignments[cell_index] if cell_index < len(alignments) else "left"
                if alignment != "left":
                    block = cell.content[0]
                    block.p_pr = block.p_pr or el.pPr()
                    block.p_pr.jc = el.jc(val="center" if alignment == "center" else "right")
            row = tr(padded, header=position == 0)
            table.content.append(row)
        link_parents(table)
        self.results.append(table)

    def table_cell(self, tokens: list, start: int, stop: int, header: bool) -> Any:
        """One cell: GFM cells are inline only, so it is one paragraph."""
        cell = tc("")
        paragraph = cell.content[0]
        saved_p, saved_link = self.current_p, self.current_hyperlink
        self.current_p, self.current_hyperlink = paragraph, None
        if header:
            self.bold += 1
        self.inlines(tokens, start, stop)
        if header:
            self.bold -= 1
        self.current_p, self.current_hyperlink = saved_p, saved_link
        return cell

    def writable_width(self) -> int:
        """The section's page width minus its margins, in twips; A4 if unstated.

        CR-003 section 13.5: the one function to move. It is now
        :func:`docx4j_py.model.content.table.writable_width`, which is what
        ``Body.insert_table`` sizes its grid from, so a markdown table and an
        ``insert_table`` table are the same width by construction.
        """
        return writable_width(self.body)

    # -- inlines -----------------------------------------------------------

    def inlines(self, tokens: list, start: int, stop: int) -> None:
        """The ``inline`` tokens in ``[start, stop)`` and their children."""
        for index in range(start, stop):
            if tokens[index].type == "inline":
                self.inline_children(tokens[index].children or [])

    def inline_children(self, children: list) -> None:
        """One inline token's children: text, emphasis, code, links, images."""
        for token in children:
            kind = token.type
            if kind == "text":
                self.add_text(token.content)
            elif kind == "strong_open":
                self.bold += 1
            elif kind == "strong_close":
                self.bold -= 1
            elif kind == "em_open":
                self.italic += 1
            elif kind == "em_close":
                self.italic -= 1
            elif kind == "s_open":
                self.strike += 1
            elif kind == "s_close":
                self.strike -= 1
            elif kind == "code_inline":
                self.add_code(token.content)
            elif kind == "softbreak":
                self.add_text(" ")
            elif kind == "hardbreak":
                self.run_target().append(R(content=ChildList([el.br()])))
            elif kind == "link_open":
                self.open_hyperlink(str(token.attrs.get("href") or ""))
            elif kind == "link_close":
                self.current_hyperlink = None
            elif kind == "image":
                self.image(token)
            elif kind == "html_inline":
                self.recorder.warn(f"inline HTML was skipped: {token.content[:40]}")

    def image(self, token: Any) -> None:
        """An image: a **local file** is embedded, a remote one stays a link.

        docx4j-mcp's posture (its section 6, and Java's
        ``DefaultMarkdownImageHandler``, which declines a remote URL): an agent
        must not be able to make the library open a socket, so a ``http:`` or
        ``https:`` destination becomes a link whose text is the alt text and
        nothing is fetched. A destination that names a **file on disk** is read
        and embedded as a real
        :class:`~docx4j_py.model.content.picture.InlinePicture`, through the
        same ``add_image`` that ``insert_inline_picture`` uses (CR-003 section
        13.5); a file this cannot read --- missing, or not a PNG, JPEG, GIF or
        BMP --- falls back to the link, with the reason in ``warnings``.
        """
        destination = str(token.attrs.get("src") or "")
        alt = token.content or str(token.attrs.get("alt") or "")
        if destination and self.embed_image(destination, alt):
            return
        if destination:
            self.recorder.warn(f"the image {destination[:60]!r} was kept as a link, not fetched")
            self.open_hyperlink(destination)
            self.add_text(alt or destination)
            self.current_hyperlink = None
        elif alt:
            self.add_text(alt)

    def embed_image(self, destination: str, alt: str) -> bool:
        """Embed a local image file; False when it is remote or cannot be read."""
        from urllib.parse import urlparse

        scheme = urlparse(destination).scheme
        if scheme and scheme not in ("file",) and len(scheme) > 1:
            return False
        path = Path(destination[7:] if scheme == "file" else destination)
        try:
            data = path.read_bytes()
        except OSError as error:
            self.recorder.warn(f"the image {destination[:60]!r} could not be read: {error.strerror}")
            return False
        from docx4j_py.model.content.picture import add_image

        try:
            made = add_image(self.body, data, alt_text_description=alt, name=path.name)
        except ContentError as error:
            self.recorder.warn(f"the image {destination[:60]!r} was not embedded: {error.message}")
            return False
        target = self.run_target()
        target.append(made.run)
        link_parents(made.run)
        made.run.parent = self.current_hyperlink or self.current_p
        self.touched.add(str(made.image_part.part_name))
        return True

    def open_hyperlink(self, destination: str) -> None:
        """Start a ``w:hyperlink``, with an external relationship on this part."""
        if not destination or self.current_hyperlink is not None:
            return
        rels = getattr(self.part, "relationships_part", None)
        if rels is None:
            self.recorder.warn(f"a link to {destination[:60]!r} has no part to hang a rel on")
            return
        rel = rels.add_external_relationship(_HYPERLINK_REL, destination)
        hyperlink = el.hyperlink(id=rel.id)
        self.style(STYLE_IDS["hyperlink"])
        self.ensure_paragraph()
        self.current_p.content.append(hyperlink)
        hyperlink.parent = self.current_p
        self.current_hyperlink = hyperlink

    def add_text(self, value: str) -> None:
        """Text in the current inline formatting."""
        if not value:
            return
        run = R(content=ChildList([t(value)]))
        r_pr = self.run_r_pr()
        if r_pr is not None:
            run.r_pr = r_pr
            r_pr.parent = run
        self.run_target().append(run)

    def add_code(self, value: str) -> None:
        """Inline code: the ``CodeChar`` character style, as Java writes it."""
        style_id = self.chosen("code_char")
        run = R(content=ChildList([t(value)]))
        r_pr = self.run_r_pr() or RPr()
        if style_id:
            r_pr.r_style = el.rStyle(val=style_id)
        run.r_pr = r_pr
        r_pr.parent = run
        self.run_target().append(run)

    def run_r_pr(self) -> Any:
        """The ``w:rPr`` for the current formatting, or None when there is none."""
        if not (self.bold or self.italic or self.strike or self.current_hyperlink is not None):
            return None
        r_pr = RPr()
        if self.current_hyperlink is not None:
            r_pr.r_style = el.rStyle(val=STYLE_IDS["hyperlink"])
        if self.bold:
            r_pr.b = el.b()
        if self.italic:
            r_pr.i = el.i()
        if self.strike:
            r_pr.strike = el.strike()
        return r_pr

    def run_target(self) -> list:
        """Where a run goes: the open hyperlink's content, else the paragraph's."""
        if self.current_hyperlink is not None:
            return self.current_hyperlink.content
        self.ensure_paragraph()
        return self.current_p.content

    # -- paragraphs --------------------------------------------------------

    def new_paragraph(self, p_pr: Any) -> None:
        """Start a paragraph and put it in the results."""
        paragraph = P(content=ChildList([]))
        if p_pr is not None:
            paragraph.p_pr = p_pr
            p_pr.parent = paragraph
        self.results.append(paragraph)
        self.current_p = paragraph

    def ensure_paragraph(self) -> None:
        """Open a paragraph if none is open."""
        if self.current_p is None:
            self.new_paragraph(self.context_p_pr())

    def context_p_pr(self) -> Any:
        """The ``w:pPr`` of a plain paragraph in the current block context.

        A quote wins over a list, as Java's ``contextPPr`` decides: a block
        quote inside a list item is a quote and carries no numbering.
        """
        if self.quote_depth:
            p_pr = _p_pr(self.chosen("quote"))
            if self.quote_depth > 1 and p_pr is not None:
                p_pr.ind = el.ind(left=TWIPS_PER_LEVEL * (self.quote_depth - 1))
            return p_pr
        if self.lists:
            context = self.lists[-1]
            p_pr = _p_pr(self.chosen("list")) or el.pPr()
            if context.item_pending:
                context.item_pending = False
                p_pr.num_pr = el.numPr(
                    ilvl=el.ilvl(val=context.ilvl), num_id=el.numId(val=context.num_id)
                )
            else:
                p_pr.ind = el.ind(left=TWIPS_PER_LEVEL * (context.ilvl + 1))
            p_pr.contextual_spacing = el.contextualSpacing(val=context.tight)
            return p_pr
        return None


# ---------------------------------------------------------------------------
# token helpers
# ---------------------------------------------------------------------------


def _closing(tokens: list, index: int, closing: str) -> int:
    """The index of the token that closes the one at `index`."""
    depth = 0
    opening = tokens[index].type
    for position in range(index, len(tokens)):
        kind = tokens[position].type
        if kind == opening:
            depth += 1
        elif kind == closing:
            depth -= 1
            if depth == 0:
                return position
    return len(tokens) - 1


def _is_tight(tokens: list, start: int, stop: int) -> bool:
    """True when markdown-it made the list tight (``hidden`` paragraph tokens)."""
    for index in range(start, stop):
        token = tokens[index]
        if token.type == "paragraph_open" and token.level == tokens[start].level + 2:
            return bool(token.hidden)
    return True


def _signature(tokens: list, start: int, stop: int) -> list[str]:
    """The numbering format at each of the nine levels of one top-level list."""
    signature: list[str | None] = [None] * LEVELS
    depth_of: dict[int, int] = {}
    depth = -1
    for index in range(start, stop + 1):
        token = tokens[index]
        if token.type in ("bullet_list_open", "ordered_list_open"):
            depth += 1
            depth_of[index] = depth
            if depth < LEVELS and signature[depth] is None:
                signature[depth] = "decimal" if token.type == "ordered_list_open" else "bullet"
        elif token.type in ("bullet_list_close", "ordered_list_close"):
            depth -= 1
    for index in range(LEVELS):
        if signature[index] is None:
            signature[index] = "bullet" if index == 0 else signature[index - 1]
    return [value or "bullet" for value in signature]


def _alignment(token: Any) -> str:
    """A GFM column's alignment, from markdown-it's inline style attribute."""
    style = str(token.attrs.get("style") or "")
    if "center" in style:
        return "center"
    if "right" in style:
        return "right"
    return "left"


def _p_pr(style_id: str) -> Any:
    """A ``w:pPr`` carrying a paragraph style, or None when there is no style."""
    return el.pPr(p_style=el.pStyle(val=style_id)) if style_id else None


# ---------------------------------------------------------------------------
# the entry point
# ---------------------------------------------------------------------------


_parser: Any = None


def parser() -> Any:
    """The shared ``MarkdownIt``, built on first use.

    ``markdown-it-py`` is imported **here** rather than at module level, so
    that ``import docx4j_py`` does not pay its 30 ms and ``to_markdown`` never
    pays it at all (CR-003 decided question 6). One parser serves the process,
    as :func:`docx4j_py.runtime.context` does for the model: a ``MarkdownIt``
    holds no per-parse state --- the block and inline states are made inside
    ``parse`` --- so it is shared where a ``ParserConfig`` could not be
    (CR-001 section 14.7).

    The preset is **commonmark** with GFM's tables and strikethrough enabled.
    ``gfm-like`` is deliberately not used: it turns on ``linkify``, which needs
    ``linkify-it-py``, a second dependency this CR does not take.
    """
    global _parser  # one parser for the process
    if _parser is None:
        try:
            from markdown_it import MarkdownIt
        except ImportError as error:  # pragma: no cover - the dependency is declared
            raise ContentError(
                "insert_markdown needs markdown-it-py, which is not installed",
                code="markdown.no_parser",
                hint="pip install 'markdown-it-py>=3'",
            ) from error

        _parser = MarkdownIt("commonmark").enable(["table", "strikethrough"])
    return _parser


def blocks_for(body: Any, markdown: str, recorder: Any) -> tuple[list[Any], set[str]]:
    """Parse markdown and build the block-level elements it describes.

    The parser is :func:`parser`, which imports ``markdown-it-py`` on first use.

    Returns:
        The elements, and the names of the parts the call changed besides the
        body's own.
    """
    importer = _Importer(body, recorder)
    elements = importer.run(parser().parse(markdown))
    if not elements:
        recorder.warn("this markdown produced no block-level content")
    return elements, importer.touched
