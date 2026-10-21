"""The text model's walk: ``scope(view(resolve(package)))`` in one pass. CR-011 section 3.2.

docx4j's ``org.docx4j.model.text.Extractor`` (its CR-046 section 2.3), as
core-ts CR-012 ports it over typed objects, here over this package's model.
One pass over the unresolved tree makes the block model:

* **resolve** (item 1): each ``mc:AlternateContent`` is read from the branch
  a reader of text takes (:func:`docx4j_py.child.mce_branch`: by namespace
  URI over docx4j's published list, REQ-302 to 306), and
  ``mce.branches_differ`` is warned where the branches' texts differ
  (REQ-307). A kept branch is a wildcard tree in this model (CR-002's lossless
  MCE): its run content is read by element name, and a text box's
  ``w:txbxContent`` in it is read as a typed island (CR-010 section 3.5),
  shared with the views' plan;
* **view** (item 2): ``accepted`` and ``original`` ask a
  :class:`~docx4j_py.model.text.view.ResolvedView` what Accept All or Reject
  All would remove or join; ``markup`` reads every run with its revision;
* **scope** (item 3): the main story, each text box once after its anchoring
  paragraph (after the table for one in a cell), the notes in reference order
  (REQ-503), headers and footers by section under ``include`` (REQ-502),
  comments where included (REQ-504), field codes under ``field_codes``
  (REQ-505), hidden text through the styles (REQ-507), placeholders
  (REQ-508), list labels with the level's owner (REQ-510), the table of
  contents under ``toc`` (REQ-520).

Nothing in the package is changed (REQ-204): every tree here was read with
``read_contents``.

Not yet (CR-011 Phase C): fields computed under ``fields: update``, bound
values, math as LaTeX (an equation is its ``m:t`` text inside ``$...$``,
REQ-604's form for one that is not converted), charts and SmartArt, WordArt,
embedded packages, ``w:altChunk``, images and their alt text, page markers.
"""

from __future__ import annotations

import re
from typing import Any

from docx4j_py.child import (
    MC_CHOICE,
    is_alternate_content,
    is_any_element,
    iter_children,
    mce_branch,
    mce_branch_content,
    mce_choices,
    mce_fallback,
    requires_of,
)
from docx4j_py.model.revisions import trees as rtrees
from docx4j_py.model.text.options import TextOptions
from docx4j_py.model.text.result import (
    TextBlock,
    TextCell,
    TextComment,
    TextField,
    TextNumbering,
    TextResult,
    TextRevision,
    TextRow,
    TextRun,
    TextRunSource,
    TextWarning,
)
from docx4j_py.model.text.trees import Context, PackageTrees

__all__ = ["Extractor", "normalize"]

W_NS = rtrees.W_NS
_W = f"{{{W_NS}}}"
_MATH_NS = rtrees.MATH_NS
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_TXBX = f"{_W}txbxContent"

#: Fonts whose glyphs are all one width (WD 1.3.2's list, docx4j's).
MONOSPACE = frozenset(
    {
        "courier new", "courier", "consolas", "lucida console", "lucida sans typewriter", "menlo", "monaco",
        "cascadia mono", "cascadia code", "source code pro", "dejavu sans mono", "liberation mono",
        "andale mono", "fira code", "fira mono", "inconsolata", "roboto mono", "ubuntu mono", "jetbrains mono",
        "noto sans mono", "ocr a extended", "letter gothic", "pt mono", "sf mono", "ibm plex mono",
    }
)  # fmt: skip

_CONSTANT = {"tab": "\t", "ptab": "\t", "br": "\n", "cr": "\n", "noBreakHyphen": "‑", "softHyphen": "­"}

#: REQ-1201's normalisation: CriticMarkup's brackets, then whitespace, \p{Z}, *, _, \, {, } and $.
_MARKS = re.compile(r"\{\+\+|\+\+\}|\{--|--\}|\{>>|<<\}")


def normalize(value: str | None) -> str:
    """WD REQ-1201's normalisation, under which texts are compared."""
    import unicodedata

    if not value:
        return ""
    removed = set(" \t\n\x0b\f\r*_\\{}$")
    return "".join(
        c for c in _MARKS.sub("", value) if c not in removed and unicodedata.category(c) not in ("Zs", "Zl", "Zp")
    )


def _v(value: Any) -> Any:
    """An enum's value, else the value."""
    return getattr(value, "value", value)


def _on(flag: Any) -> bool:
    """An on/off property: present, and not ``w:val="0"``."""
    return flag is not None and getattr(flag, "val", None) is not False


def _local(item: Any) -> str:
    """A list member's WordprocessingML (or Office Math, markup-compatibility) local name; '' for another namespace's wildcard."""
    if is_any_element(item):
        qname = item.qname
        if not isinstance(qname, str):
            return ""
        uri, _, local = qname.rpartition("}")
        return local if uri[1:] in (W_NS, _MATH_NS, "http://schemas.openxmlformats.org/markup-compatibility/2006") else ""
    return rtrees.name(item)


def _attr(item: Any, local: str) -> Any:
    """A ``w:`` attribute of a typed or wildcard node."""
    return rtrees.attribute(item, local)


def _date_text(value: Any) -> str | None:
    """A ``w:date`` as the package wrote it."""
    if value is None:
        return None
    to_string = getattr(value, "to_string", None)
    if callable(to_string):
        try:
            return to_string()
        except Exception:  # noqa: BLE001 - fall back to str
            return str(value)
    return str(value)


def _revision(holder: Any, kind: str) -> TextRevision:
    return TextRevision(kind, _attr(holder, "author"), _date_text(_attr(holder, "date")))  # type: ignore[arg-type]


def _is_monospace(font: str | None) -> bool:
    return font is not None and font.strip().lower() in MONOSPACE


# ---------------------------------------------------------------------------
# Office JS's paragraph text, for each run's office_offset (CR-011 Q4)
# ---------------------------------------------------------------------------


def _office_spans(p: Any, view: str) -> dict[int, tuple[int, int]]:
    """Each run item's place in the content API's ``Paragraph.get_text(view=)``: ``id`` -> (offset, length)."""
    from docx4j_py.model.content.text_model import segments_of

    try:
        segments = segments_of(p, view="original" if view == "original" else "accepted")
    except Exception:  # noqa: BLE001 - no offsets, not no text
        return {}
    return {id(s.item): (s.start, s.end - s.start) for s in segments}


# ---------------------------------------------------------------------------
# the walk's state
# ---------------------------------------------------------------------------


class _Field:
    """An open complex field."""

    __slots__ = (
        "code_block",
        "code_start",
        "floor_paragraph",
        "instr",
        "link",
        "previous_floor",
        "result",
        "separated",
        "toc",
    )

    def __init__(self) -> None:
        self.instr = ""
        self.result = ""
        self.separated = False
        #: A HYPERLINK field's target, once its code is read.
        self.link: str | None = None
        #: A table of contents under ``toc: false`` (docx4j D12): nothing of it is read.
        self.toc = False
        self.code_block: TextBlock | None = None
        self.code_start = 0
        self.floor_paragraph: _Paragraph | None = None
        self.previous_floor = 0


class _RunState:
    """The revision and the control a run-level walk is in."""

    __slots__ = ("kept", "placeholder", "revision")

    def __init__(self, revision: TextRevision | None, kept: bool, placeholder: bool) -> None:
        #: markup: the innermost revision; else None.
        self.revision = revision
        #: accepted, original: inside a revision holder the view keeps.
        self.kept = kept
        self.placeholder = placeholder

    def with_revision(self, revision: TextRevision | None, kept: bool) -> _RunState:
        return _RunState(revision, kept, self.placeholder)

    def with_placeholder(self) -> _RunState:
        return _RunState(self.revision, self.kept, True)


_PLAIN = _RunState(None, False, False)


class _Story:
    """The walk of one story."""

    __slots__ = (
        "fields",
        "kind",
        "note_id",
        "note_number",
        "note_type",
        "part",
        "pending",
        "prefix",
        "sections",
        "state",
        "table_depth",
    )

    def __init__(self, kind: str, prefix: str, part: Any, state: Any) -> None:
        self.kind = kind
        self.prefix = prefix
        self.part = part
        #: Open complex fields; [0] is the innermost.
        self.fields: list[_Field] = []
        #: The text boxes anchored since the last flush, to follow their anchor (or its table).
        self.pending: list[tuple[Any, str]] = []
        self.table_depth = 0
        #: The sections ended so far (the main story only).
        self.sections: list[Any] | None = None
        self.note_type: str | None = None
        self.note_number = 0
        self.note_id: str | None = None
        #: The numbering state the story counts its lists in.
        self.state = state


class _Paragraph:
    """A paragraph being read: the block its runs go to, its address, what it anchored so far."""

    __slots__ = (
        "_office",
        "address",
        "block",
        "boxes",
        "display_math",
        "format",
        "link",
        "math",
        "merge_floor",
        "monospace",
        "p",
        "p_style",
        "story",
        "text_chars",
        "toc",
        "view",
    )

    def __init__(self, block: TextBlock, p: Any, address: str, story: _Story, view: str) -> None:
        self.block = block
        self.p = p
        self.address = address
        self.story = story
        self.view = view
        self.boxes = 0
        p_pr = getattr(p, "p_pr", None)
        style = getattr(p_pr, "p_style", None) if p_pr is not None else None
        #: The paragraph's own style id, for its runs' styles.
        self.p_style = _v(style.val) if style is not None else None
        #: How many characters of text its runs have added.
        self.text_chars = 0
        #: Whether every run with text so far is monospaced; None before the first.
        self.monospace: bool | None = None
        #: The formatting of the run being read.
        self.format: TextRun | None = None
        #: The target of the ``w:hyperlink`` being read.
        self.link: str | None = None
        self.math = False
        self.display_math = False
        #: A table of contents left out is in the paragraph.
        self.toc = False
        #: The first of the block's runs a run may be merged into.
        self.merge_floor = 0
        self._office: dict[int, tuple[int, int]] | None = None

    def office_span(self, item: Any) -> tuple[int, int] | None:
        if self._office is None:
            self._office = _office_spans(self.p, self.view)
        return self._office.get(id(item))


# ---------------------------------------------------------------------------
# the extractor
# ---------------------------------------------------------------------------


class Extractor:
    """One reading of a package in one view: the block model. docx4j ``Extractor``."""

    def __init__(self, package_trees: PackageTrees, options: TextOptions, *, resolved: bool = False) -> None:
        """A walk over the trees.

        Args:
            package_trees: what the reading walks.
            options: the options, checked.
            resolved: the package was resolved already by Accept All or Reject
                All (the copy route of REQ-406): no view logic at all.
        """
        self.trees = package_trees
        self.options = options
        self.view = options.view
        self.resolved = resolved
        self.context = Context(package_trees)
        self.plan = None
        #: Content with no package to plan over: run-level revisions only, no joins or moves (core-ts's planless mode).
        self.planless = package_trees.document is None and not resolved
        if not resolved and self.view != "markup" and package_trees.document is not None:
            self.plan = self.context.plan(self.view == "accepted")
        self.resolver = package_trees.resolver
        self.warnings: list[TextWarning] = []
        self._warned: set[tuple[Any, ...]] = set()
        self._style_rpr: dict[tuple[Any, Any], Any] = {}
        self._typed: dict[int, Any] = {}
        #: What every address this reading makes starts with: empty, or an
        #: embedded package's or a chunk's address and a slash.
        self.root = ""
        #: A reading one level down (an embedded package's, an ``altChunk``'s):
        #: none of its own are expanded, and it has no metadata (REQ-512, REQ-513).
        self.nested = False

    # -- the package ----------------------------------------------------------

    def extract(self) -> TextResult:
        """The package's text: REQ-501's order."""
        out: list[TextBlock] = []
        comments: dict[str, TextComment] = {}
        metadata: dict[str, Any] = {TextResult.FIELDS_UPDATE_ON_OPEN: self._fields_update_on_open()}
        if not self.nested:
            from docx4j_py.model.text import metadata as properties

            properties.read(self.trees.package, metadata)
        body = self.trees.body
        if body is None:
            return TextResult(out, comments, self.warnings, metadata, self.options)
        story = _Story("MAIN", "body", self.trees.main, self.context.numbering_states.main)
        story.sections = []
        main_blocks: list[TextBlock] = []
        self.blocks(body.content, "body", story, main_blocks)
        story.sections.append(body.sect_pr)
        headers = "headers" in self.options.includes
        footers = "footers" in self.options.includes
        if headers or footers:
            self._sections_with_headers(story.sections, main_blocks, headers, footers, out)
        else:
            out.extend(main_blocks)
        if self.options.notes:
            self._notes("footnote", out)
            self._notes("endnote", out)
        if self.comments_included():
            self._comments(comments)
        return TextResult(out, comments, self.warnings, metadata, self.options)

    def extract_content(self, content: Any, part: Any) -> TextResult:
        """Content in the package (a paragraph, a table, a body) as the main story, addressed ``content``."""
        out: list[TextBlock] = []
        comments: dict[str, TextComment] = {}
        state = self.context.numbering_states.for_part(part) if part is not None else self.context.numbering_states.main
        story = _Story("MAIN", "content", part, state)
        items = rtrees.block_items(content)
        if items is not None and not isinstance(content, rtrees.cls("P")):
            self.blocks(items, "content", story, out)
        else:
            self.blocks([content], "content", story, out)
        if self.comments_included():
            self._comments(comments)
        metadata = {TextResult.FIELDS_UPDATE_ON_OPEN: self._fields_update_on_open()}
        return TextResult(out, comments, self.warnings, metadata, self.options)

    def _fields_update_on_open(self) -> bool:
        settings = self.trees.settings
        return _on(getattr(settings, "update_fields", None)) if settings is not None else False

    def comments_included(self) -> bool:
        return self.view == "markup" or self.options.comments

    # -- block level ------------------------------------------------------------

    def blocks(
        self, items: list | None, prefix: str, story: _Story, sink: list[TextBlock], state: _RunState = _PLAIN
    ) -> None:
        """Block-level content into `sink`: paragraphs (joined in a view), tables, and what is transparent."""
        if not items:
            return
        p_class, tbl_class = rtrees.cls("P"), rtrees.cls("Tbl")
        i = 0
        n = len(items)
        while i < n:
            item = items[i]
            if isinstance(item, str) or self.removes(item):
                i += 1
                continue
            if is_any_element(item):
                typed = self._typed_block(item)
                if typed is None:
                    i += 1
                    continue
                item = typed
            address = f"{prefix}/{i}"
            if isinstance(item, p_class):
                chain = [item]
                addresses = [address]
                j = i
                while self._joins_next(chain[-1]):
                    k = self._next_paragraph(items, j)
                    if k < 0:
                        break  # no paragraph next: only the mark goes
                    chain.append(items[k])
                    addresses.append(f"{prefix}/{k}")
                    j = k
                self._paragraph(chain, addresses, story, sink, state)
                i = j + 1
                continue
            if isinstance(item, tbl_class):
                self._table(item, address, story, sink, state)
            elif is_alternate_content(item):
                self._compare_branches(item, address)
                self.blocks(mce_branch_content(mce_branch(item)), address, story, sink, state)
            elif rtrees.revision_of(item) is not None:
                self.blocks(rtrees.run_items(item), address, story, sink, self._holder_state(item, state))
            elif rtrees.is_holder(item) and self.unwraps(item):
                # an inserted control rejected: its content in its place
                self.blocks(rtrees.children(item), address, story, sink, state)
            elif _local(item) == "sdt":
                placeholder = self._showing_placeholder(item)
                if not (placeholder and not self.options.placeholders):
                    inner = state.with_placeholder() if placeholder else state
                    self.blocks(rtrees.block_items(item), address, story, sink, inner)
            else:
                inner_items = rtrees.block_items(item)
                if inner_items is not None:
                    self.blocks(inner_items, address, story, sink, state)
            i += 1

    def _next_paragraph(self, items: list, start: int) -> int:
        """The index of the paragraph a paragraph joins: the next past markers and what the view removes, or -1."""
        for k in range(start + 1, len(items)):
            item = items[k]
            if _is_marker(item) or self.removes(item):
                continue
            return k if isinstance(item, rtrees.cls("P")) else -1
        return -1

    def _paragraph(
        self, chain: list[Any], addresses: list[str], story: _Story, sink: list[TextBlock], state: _RunState
    ) -> None:
        first, last = chain[0], chain[-1]
        block = TextBlock("PARAGRAPH", story.kind, addresses[0])  # type: ignore[arg-type]
        block.para_id = getattr(first, "para_id", None)
        block.source = first
        # a joined paragraph keeps the last one's properties (CR-043's join)
        p_pr = getattr(last, "p_pr", None)
        block.style = _v(getattr(getattr(p_pr, "p_style", None), "val", None))
        block.note_type = story.note_type
        block.note_number = story.note_number
        block.note_id = story.note_id
        monospace: bool | None = None
        toc = False
        for p, address in zip(chain, addresses, strict=True):
            block.revisions += _mark_revisions(p)
            para = _Paragraph(block, p, address, story, self.view)
            starts_in_toc = _in_toc(story)
            self.runs(p.content, para, state)
            if starts_in_toc or para.toc or _in_toc(story):
                toc = True
            if para.monospace is not None:
                monospace = para.monospace if monospace is None else (monospace and para.monospace)
        block.monospace = monospace is True
        self._properties(block, last)
        mark_hidden = self._is_hidden_mark(p_pr)
        any_visible = any(not r.hidden for r in block.runs)
        hidden_paragraph = mark_hidden and not any_visible
        if hidden_paragraph and not self.options.hidden:
            # runs and mark hidden: no line at all (REQ-507)
            block.runs.clear()
            self._flush_outside_table(story, sink)
            self._end_section(last, story)
            return
        block.hidden = hidden_paragraph
        if not self.options.hidden:
            block.runs[:] = [r for r in block.runs if not r.hidden]
        if toc and not block.runs:
            # a table of contents' paragraph, under toc: false (D12): no line at all
            self._flush_outside_table(story, sink)
            self._end_section(last, story)
            return
        # one count per paragraph: the label (list_labels) and the numbering's value come from it
        if self.trees.emulator is not None and (self.options.list_labels or block.numbering is not None):
            counted = self._count(last, story)
            if self.options.list_labels:
                block.list_label = _list_label(counted)
            if block.numbering is not None and counted is not None:
                block.numbering.value = counted.count
        block.section = len(story.sections) if story.sections is not None else 0
        sink.append(block)
        self._flush_outside_table(story, sink)
        self._end_section(last, story)

    def _end_section(self, last: Any, story: _Story) -> None:
        p_pr = getattr(last, "p_pr", None)
        if story.sections is not None and story.table_depth == 0 and p_pr is not None and p_pr.sect_pr is not None:
            story.sections.append(p_pr.sect_pr)

    def _flush_outside_table(self, story: _Story, sink: list[TextBlock]) -> None:
        """What a paragraph anchored, after it; in a table, after the table (REQ-518)."""
        if story.table_depth > 0 or not story.pending:
            return
        pending = list(story.pending)
        story.pending.clear()
        section = len(story.sections) if story.sections is not None else 0
        for content, prefix in pending:
            inner = _Story("TEXT_BOX", prefix, story.part, story.state)
            inner.note_type = story.note_type
            inner.note_number = story.note_number
            inner.note_id = story.note_id
            blocks: list[TextBlock] = []
            self.blocks(content.content, prefix, inner, blocks)
            for b in blocks:
                b.section = section
            sink.extend(blocks)

    def _properties(self, block: TextBlock, p: Any) -> None:
        own = getattr(p, "p_pr", None)
        block.bottom_border = own is not None and own.p_bdr is not None and own.p_bdr.bottom is not None
        effective = own
        if self.resolver is not None:
            try:
                effective = self.resolver.get_effective_ppr(own)
            except Exception:  # noqa: BLE001 - not resolved: the paragraph's own properties
                effective = own
        if effective is None:
            return
        ind = effective.ind
        if ind is not None:
            block.indent_left = _twips(ind.left if ind.left is not None else ind.start)
            block.indent_hanging = _twips(ind.hanging)
            block.indent_first_line = _twips(ind.first_line)
        if effective.jc is not None and effective.jc.val is not None:
            block.alignment = str(_v(effective.jc.val))
        level = _v(getattr(effective.outline_lvl, "val", None)) if effective.outline_lvl is not None else None
        if level is not None:
            try:
                number = int(level)
            except (TypeError, ValueError):
                number = -1
            if 0 <= number <= 8:
                block.outline_level = number  # 9 is body text
        if effective.contextual_spacing is not None:
            block.contextual_spacing = _on(effective.contextual_spacing)
        direct = own is not None and own.num_pr is not None and own.num_pr.num_id is not None
        block.numbering = self._numbering(effective.num_pr, direct, block.style)

    def _numbering(self, num_pr: Any, direct: bool, p_style: str | None) -> TextNumbering | None:
        """A paragraph's numbering; None for none, for ``w:numId`` 0, and where Word numbers it not (REQ-510)."""
        if num_pr is None or num_pr.num_id is None:
            return None
        num_id_value = _v(num_pr.num_id.val)
        if num_id_value is None:
            return None
        num_id = str(num_id_value)
        if num_id == "0":
            return None
        ilvl_value = _v(num_pr.ilvl.val) if num_pr.ilvl is not None else None
        try:
            ilvl = int(ilvl_value) if ilvl_value is not None else 0
        except (TypeError, ValueError):
            ilvl = 0
        emulator = self.trees.emulator
        fmt: str | None = None
        text: str | None = None
        abstract: str | None = None
        start = 1
        if emulator is not None:
            try:
                if not direct and emulator.style_linked_elsewhere(num_id, str(ilvl), p_style):
                    return None
                definition = emulator.instance_list_definitions.get(num_id)
                if definition is not None:
                    if definition.abstract_definition is not None:
                        abstract = definition.abstract_definition.id
                    level = definition.level(str(ilvl))
                    if level is not None:
                        fmt = level.num_fmt if level.num_fmt is not None else ("bullet" if level.is_bullet else None)
                        text = level.lvl_text
                        start = level.start_value + 1
            except Exception:  # noqa: BLE001 - the definitions could not be read
                pass
        return TextNumbering(num_id, ilvl, None if fmt is None else str(fmt), text, start, None, abstract)

    def _count(self, p: Any, story: _Story) -> Any:
        try:
            return self.trees.emulator.number(getattr(p, "p_pr", None), story.state)
        except Exception:  # noqa: BLE001 - not counted
            return None

    # -- tables ---------------------------------------------------------------

    def _table(self, tbl: Any, address: str, story: _Story, sink: list[TextBlock], state: _RunState) -> None:
        block = TextBlock("TABLE", story.kind, address)  # type: ignore[arg-type]
        block.section = len(story.sections) if story.sections is not None else 0
        block.source = tbl
        story.table_depth += 1
        try:
            self._rows(tbl.content, address, story, block, state)
        finally:
            story.table_depth -= 1
        if block.rows:
            sink.append(block)  # a table left with no rows goes (CR-043)
        self._flush_outside_table(story, sink)

    def _rows(self, items: list | None, prefix: str, story: _Story, table: TextBlock, state: _RunState) -> None:
        if not items:
            return
        tr_class = rtrees.cls("Tr")
        for i, item in enumerate(items):
            if isinstance(item, str) or self.removes(item):
                continue
            address = f"{prefix}/{i}"
            if isinstance(item, tr_class):
                row = TextRow(address)
                tr_pr = item.tr_pr
                row_state = state
                if tr_pr is not None:
                    for entry in tr_pr.content or ():
                        if rtrees.name(entry) == "tblHeader":
                            row.header = _on(entry)
                    if self.view == "markup" and not self.resolved:
                        if tr_pr.ins is not None:
                            row_state = state.with_revision(_revision(tr_pr.ins, "INSERTION"), False)
                        if tr_pr.del_value is not None:
                            row_state = state.with_revision(_revision(tr_pr.del_value, "DELETION"), False)
                self._cells(item.content, address, story, row, row_state)
                table.rows.append(row)
            else:
                inner = rtrees.block_items(item)
                if inner is not None and not isinstance(item, rtrees.cls("Tc")):
                    self._rows(inner, address, story, table, state)

    def _cells(self, items: list | None, prefix: str, story: _Story, row: TextRow, state: _RunState) -> None:
        if not items:
            return
        tc_class = rtrees.cls("Tc")
        for i, item in enumerate(items):
            if isinstance(item, str) or self.removes(item):
                continue
            address = f"{prefix}/{i}"
            if isinstance(item, tc_class):
                cell = TextCell(address)
                tc_pr = item.tc_pr
                cell_state = state
                if tc_pr is not None:
                    if tc_pr.grid_span is not None and tc_pr.grid_span.val is not None:
                        cell.grid_span = _twips(_v(tc_pr.grid_span.val)) or 1
                    if tc_pr.v_merge is not None:
                        if _v(tc_pr.v_merge.val) == "restart":
                            cell.merge_restart = True
                        else:
                            cell.merge_continue = True
                    marker = tc_pr.cell_ins_or_cell_del_or_cell_merge
                    if self.view == "markup" and not self.resolved and marker is not None:
                        local = rtrees.name(marker)
                        if local == "cellIns":
                            cell_state = state.with_revision(_revision(marker, "INSERTION"), False)
                        elif local == "cellDel":
                            cell_state = state.with_revision(_revision(marker, "DELETION"), False)
                self.blocks(item.content, address, story, cell.blocks, cell_state)
                row.cells.append(cell)
            else:
                inner = rtrees.block_items(item)
                if inner is not None:
                    self._cells(inner, address, story, row, state)

    # -- run level --------------------------------------------------------------

    def runs(self, items: list | None, para: _Paragraph, state: _RunState) -> None:
        """Run-level content: runs, revision holders, controls, fields, links, equations, branches."""
        if not items:
            return
        r_class = rtrees.cls("R")
        for item in items:
            if isinstance(item, str) or self.removes(item):
                continue
            if is_any_element(item):
                typed = self._typed_run_level(item)
                if typed is None:
                    if is_alternate_content(item):
                        self._compare_branches(item, para.block.address)
                        self.runs(mce_branch_content(mce_branch(item)), para, state)
                    continue
                item = typed
            if rtrees.revision_of(item) is not None:
                para.block.revisions += 1
                self.runs(rtrees.run_items(item), para, self._holder_state(item, state))
            elif isinstance(item, r_class):
                self._run(item, para, state)
            elif is_alternate_content(item):
                self._compare_branches(item, para.block.address)
                self.runs(mce_branch_content(mce_branch(item)), para, state)
            elif rtrees.is_holder(item) and self.unwraps(item):
                self.runs(rtrees.children(item), para, state)
            elif _local(item) == "sdt":
                placeholder = self._showing_placeholder(item)
                if placeholder and not self.options.placeholders:
                    continue
                self.runs(rtrees.run_items(item), para, state.with_placeholder() if placeholder else state)
            elif _local(item) == "fldSimple":
                self._simple_field(item, para, state)
            elif rtrees.is_math(item):
                self._math(item, para, state)
            elif _local(item) == "hyperlink":
                saved = para.link
                target = self._hyperlink_target(item, para.story)
                if target is not None:
                    para.link = target
                try:
                    self.runs(item.content, para, state)
                finally:
                    para.link = saved
            else:
                inner = rtrees.run_items(item)
                if inner is not None and not isinstance(item, rtrees.cls("P")):
                    self.runs(inner, para, state)

    def _holder_state(self, item: Any, state: _RunState) -> _RunState:
        if self.view == "markup" and not self.resolved:
            revision = rtrees.revision_of(item)
            kind = "INSERTION" if revision in (rtrees.Revision.INS, rtrees.Revision.MOVE_TO) else "DELETION"
            return state.with_revision(_revision(item, kind), False)
        return state.with_revision(None, True)

    def _run(self, r: Any, para: _Paragraph, state: _RunState) -> None:
        r_pr = r.r_pr
        r_style = _r_style(r)
        paragraph = self._style_rpr_of(None, para.p_style)
        styles = paragraph if r_style is None else self._style_rpr_of(r_style, para.p_style)
        # hidden text (REQ-507): the run's own w:vanish, else its styles'
        own = r_pr.vanish if r_pr is not None else None
        hidden = _on(own) if own is not None else (styles is not None and _on(styles.vanish))
        font = _ascii_font(r_pr) or _ascii_font(styles)
        # its formatting beyond its paragraph style's (docx4j CR-046 phase 3)
        f = TextRun("TEXT")
        f.bold = _on_of(r_pr, styles, "b") and not (paragraph is not None and _on(paragraph.b))
        f.italic = _on_of(r_pr, styles, "i") and not (paragraph is not None and _on(paragraph.i))
        f.strike = _on_of(r_pr, styles, "strike") and not (paragraph is not None and _on(paragraph.strike))
        mono = _is_monospace(font)
        f.monospace = mono and not _is_monospace(_ascii_font(paragraph))
        before = para.text_chars
        saved = para.format
        para.format = f
        try:
            for item in r.content:
                if not isinstance(item, str):
                    self._run_item(item, r, hidden, para, state)
        finally:
            para.format = saved
        if para.text_chars > before:
            para.monospace = mono if para.monospace is None else (para.monospace and mono)

    def _run_item(self, item: Any, r: Any, hidden: bool, para: _Paragraph, state: _RunState) -> None:
        name = _local(item)
        story = para.story
        if name == "fldChar":
            fields = story.fields
            kind = _v(_attr(item, "fldCharType")) if is_any_element(item) else _v(getattr(item, "fld_char_type", None))
            if kind == "begin":
                f = _Field()
                f.code_block = para.block
                f.code_start = len(para.block.runs)
                if not self.options.toc and self.options.field_codes:
                    f.floor_paragraph = para
                    f.previous_floor = para.merge_floor
                    para.merge_floor = len(para.block.runs)
                fields.insert(0, f)
            elif kind == "separate":
                if fields:
                    f = fields[0]
                    f.separated = True
                    f.link = field_hyperlink_target(f.instr)
                    if not self.options.toc and not _suppressed(story) and is_toc_field(f.instr):
                        f.toc = True
                        para.toc = True
                        _remove_field_code(f, para)
            elif kind == "end" and fields:
                f = fields.pop(0)
                if _suppressed(story):
                    return
                if f.toc or (not self.options.toc and is_toc_field(f.instr)):
                    if not f.toc:
                        _remove_field_code(f, para)
                    para.toc = True
                    return
                self._end_field(f.instr, f.result if f.separated else None, para)
            return
        if name in ("instrText", "delInstrText"):
            if name == "delInstrText" and not state.kept and state.revision is None and self.view == "accepted":
                return
            value = item.text if is_any_element(item) else getattr(item, "value", None)
            if not value:
                return
            if story.fields:
                story.fields[0].instr += value
            if self.options.field_codes and not _suppressed(story):
                self._add_text(para, value, "FIELD_CODE", hidden, state)
            return
        if name == "t":
            value = item.text if is_any_element(item) else getattr(item, "value", None)
            self._text(para, value, hidden, state, (r, item))
            return
        if name == "delText":
            # the holder decides; outside one, the element: no text after Accept All (REQ-404)
            if not state.kept and state.revision is None:
                if self.view == "accepted":
                    return
                if self.view == "markup" and not self.resolved:
                    state = state.with_revision(TextRevision("DELETION"), False)
            value = item.text if is_any_element(item) else getattr(item, "value", None)
            self._text(para, value, hidden, state, (r, item))
            return
        constant = _CONSTANT.get(name)
        if constant is not None:
            self._text(para, constant, hidden, state, (r, item))
            return
        if name == "sym":
            code = _attr(item, "char")
            if isinstance(code, str) and re.fullmatch(r"[0-9A-Fa-f]{1,6}", code):
                point = int(code, 16)
                if point <= 0x10FFFF and not 0xD800 <= point <= 0xDFFF:
                    self._text(para, chr(point), hidden, state, (r, item))
            return
        if name in ("footnoteReference", "endnoteReference"):
            # in a field's code: neither written nor numbered
            if _in_code(story) or _suppressed(story):
                return
            foot = name == "footnoteReference"
            note_id = _attr(item, "id")
            run = TextRun("NOTE_REFERENCE", "", state.revision, hidden=hidden)
            run.note_type = "footnote" if foot else "endnote"
            run.note_id = "" if note_id is None else str(note_id)
            run.source = TextRunSource(r, para.p)
            if not (hidden and not self.options.hidden):
                custom = _attr(item, "customMarkFollows")
                if not is_any_element(item):
                    custom = getattr(item, "custom_mark_follows", None)
                run.note_number = self.context.note_number(run.note_type, run.note_id, custom in (True, "1", "true", "on"))
            para.block.runs.append(run)
            return
        if name in ("footnoteRef", "endnoteRef"):
            if story.note_type is None:
                return
            run = TextRun("NOTE_MARK", "", state.revision)
            run.note_type = story.note_type
            run.note_number = story.note_number
            run.note_id = story.note_id
            para.block.runs.append(run)
            return
        if name == "commentReference":
            if not self.comments_included() or _in_code(story) or _suppressed(story):
                return
            comment_id = _attr(item, "id")
            if self.plan is not None and self.plan.removes_comment(comment_id):
                return  # gone with the revision it was on
            run = TextRun("COMMENT_REFERENCE", "", state.revision, hidden=hidden)
            run.comment_id = "" if comment_id is None else str(comment_id)
            run.source = TextRunSource(r, para.p)
            para.block.runs.append(run)
            return
        if name in ("drawing", "pict", "object"):
            if not _in_code(story) and not _suppressed(story) and not (hidden and not self.options.hidden):
                self._anchored(item, para)
            return
        if name == "AlternateContent":
            self._compare_branches(item, para.block.address)
            for entry in mce_branch_content(mce_branch(item)):
                if not isinstance(entry, str):
                    self._run_item(entry, r, hidden, para, state)
            return
        if name == "ruby":
            base = getattr(item, "ruby_base", None)
            if base is not None:
                self.runs(getattr(base, "content", None), para, state)
            return
        if isinstance(item, rtrees.cls("R")):
            # a run within a run (w:r in m:r)
            self._run(item, para, state)

    def _text(
        self, para: _Paragraph, value: str | None, hidden: bool, state: _RunState, source: tuple[Any, Any] | None = None
    ) -> None:
        """Document text: a field's code where a field is in its code (shown only under ``field_codes``), else text."""
        if not value:
            return
        story = para.story
        if _in_code(story):
            if self.options.field_codes:
                self._add_text(para, value, "FIELD_CODE", hidden, state)
            return
        if _suppressed(story):
            return
        if not (hidden and not self.options.hidden):
            for f in story.fields:
                f.result += value
        self._add_text(para, value, "TEXT", hidden, state, source)

    def _add_text(
        self,
        para: _Paragraph,
        value: str,
        kind: str,
        hidden: bool,
        state: _RunState,
        source: tuple[Any, Any] | None = None,
    ) -> None:
        run = TextRun(kind, value, state.revision, hidden=hidden, placeholder=state.placeholder)  # type: ignore[arg-type]
        if kind == "TEXT":
            para.text_chars += len(value)
            if para.format is not None:
                run.copy_format(para.format)
            run.link = _link_of(para)
            run.math = para.math
            run.display_math = para.display_math
            if source is not None:
                run.source = TextRunSource(source[0], para.p)
                span = para.office_span(source[1])
                if span is not None and span[1] == len(value):
                    run.office_offset = span[0]
        runs = para.block.runs
        if len(runs) > para.merge_floor and not run.math:
            last = runs[-1]
            # one run as long as its text and Paragraph.text's agree character for character (core-ts CR-012 4.4)
            if last.office_offset is None:
                contiguous = run.office_offset is None
            else:
                contiguous = (
                    run.office_offset == last.office_offset + len(last.text)
                    and last.source is not None
                    and run.source is not None
                    and last.source.paragraph is run.source.paragraph
                )
            if (
                last.kind == kind
                and last.hidden == hidden
                and last.placeholder == state.placeholder
                and last.revision == state.revision
                and last.same_format(run)
                and contiguous
            ):
                last.text += value
                return
        runs.append(run)

    def _end_field(self, instr: str, result: str | None, para: _Paragraph) -> None:
        para.block.fields.append(TextField(instr, result))
        if result is None:
            self.warn(TextWarning(TextWarning.FIELD_NO_RESULT, para.block.address))

    def _simple_field(self, field_: Any, para: _Paragraph, state: _RunState) -> None:
        story = para.story
        if _suppressed(story):
            return
        instr = getattr(field_, "instr", None) or ""
        if not self.options.toc and is_toc_field(instr):
            para.toc = True  # a table of contents left out (D12)
            return
        if self.options.field_codes and not _in_code(story):
            self._add_text(para, instr, "FIELD_CODE", False, state)
        before = len(para.block.runs)
        capture = _Field()
        capture.separated = True
        capture.link = field_hyperlink_target(instr)
        story.fields.append(capture)
        try:
            self.runs(getattr(field_, "content", None), para, state)
        finally:
            for k in range(len(story.fields) - 1, -1, -1):
                if story.fields[k] is capture:
                    del story.fields[k]
                    break
        has_result = capture.result != "" or len(para.block.runs) > before
        para.block.fields.append(TextField(instr, capture.result if has_result else None))
        if not has_result:
            self.warn(TextWarning(TextWarning.FIELD_NO_RESULT, para.block.address))

    def _hyperlink_target(self, link: Any, story: _Story) -> str | None:
        """A ``w:hyperlink``'s target: its relationship's target, then ``#`` and its anchor."""
        target: str | None = None
        rel_id = getattr(link, "id", None)
        if rel_id and story.part is not None:
            rels = getattr(story.part, "relationships_part", None)
            try:
                rel = rels.get_relationship_by_id(rel_id) if rels is not None else None
            except Exception:  # noqa: BLE001 - a broken relationship names nothing
                rel = None
            found = getattr(rel, "target", None) if rel is not None else None
            if found:
                target = found
        anchor = getattr(link, "anchor", None) or None
        if anchor is not None:
            return f"#{anchor}" if target is None else f"{target}#{anchor}"
        return target

    # -- math (REQ-604) --------------------------------------------------------

    def _math(self, value: Any, para: _Paragraph, state: _RunState) -> None:
        """An equation: its ``m:t`` text inside ``$...$`` (``$$...$$`` for an ``m:oMathPara``) until Phase C's LaTeX."""
        display = type(value).__name__ in ("CTOMathPara", "OMathPara")
        saved = (para.math, para.display_math)
        para.math, para.display_math = True, display
        try:
            if self.view == "markup" and not self.resolved:
                deletion, insertion = _math_revisions(value)
                if deletion is not None or insertion is not None:
                    original = self.context.plan(False)
                    accepted = self.context.plan(True)
                    before = _math_text(value, lambda o: original is None or not original.removes(o))
                    after = _math_text(value, lambda o: accepted is None or not accepted.removes(o))
                    if before != after:
                        if before:
                            self._text(para, _delimit(before, display), False, state.with_revision(deletion or TextRevision("DELETION"), False))
                        if after:
                            self._text(para, _delimit(after, display), False, state.with_revision(insertion or TextRevision("INSERTION"), False))
                        return
                    if after:
                        self._text(para, _delimit(after, display), False, state)
                    return
            text = _math_text(value, lambda o: not self.removes(o))
            if text:
                self._text(para, _delimit(text, display), False, state)
        finally:
            para.math, para.display_math = saved

    # -- what a run anchors: text boxes (REQ-518) --------------------------------

    def _anchored(self, value: Any, para: _Paragraph) -> None:
        """The text boxes under a run's drawing, picture or object, branches selected, to follow the paragraph."""
        if not self.options.text_boxes:
            return
        txbx_class = rtrees.cls("CTTxbxContent")

        def visit(node: Any, holder: list | None) -> None:
            if isinstance(node, str) or node is None:
                return
            if is_alternate_content(node):
                self._compare_branches(node, para.block.address)
                branch = mce_branch(node)
                if branch is None:
                    return
                kids = branch.children if is_any_element(branch) else branch.any_element
                for child in list(kids or ()):
                    visit(child, kids if is_any_element(child) else None)
                return
            if isinstance(node, txbx_class):
                self._box(node, para)
                return
            if is_any_element(node):
                if node.qname == _TXBX:
                    typed = self._island(node, holder, para)
                    if typed is not None:
                        self._box(typed, para)
                    return
                kids = node.children or []
                for child in list(kids):
                    visit(child, kids)
                return
            for _qname, child in iter_children(node, mce="all", wildcards=True):
                visit(child, _holder_of(node, child) if is_any_element(child) else None)

        visit(value, None)

    def _box(self, content: Any, para: _Paragraph) -> None:
        para.story.pending.append((content, f"{para.address}/textbox:{para.boxes}"))
        para.boxes += 1

    def _island(self, node: Any, holder: list | None, para: _Paragraph) -> Any:
        """A wildcard ``w:txbxContent`` as a typed island, shared with the plan's listing (CR-010 section 3.5)."""
        islands = self.context.islands
        island = islands.get(id(node))
        if island is None:
            from docx4j_py.model.revisions.islands import Island

            try:
                island = Island(node, holder if holder is not None else [])
            except ValueError:
                self.warn(TextWarning(TextWarning.CONTENT_NOT_READ, para.block.address, part_name=_part_name(para.story.part)))
                return None
            islands[id(node)] = island
        return island.typed

    # -- typing a wildcard ---------------------------------------------------------

    def _typed_block(self, node: Any) -> Any:
        """A wildcard paragraph or table, in a branch at block level, parsed typed (no revision is planned in it)."""
        return self._typed_node(node, ("p", "tbl", "sdt"))

    def _typed_run_level(self, node: Any) -> Any:
        """A wildcard run, revision or control at paragraph level, parsed typed."""
        return self._typed_node(node, ("r", "ins", "del", "moveFrom", "moveTo", "hyperlink", "sdt", "fldSimple", "smartTag"))

    def _typed_node(self, node: Any, locals_: tuple[str, ...]) -> Any:
        qname = node.qname
        if not isinstance(qname, str) or not qname.startswith(_W):
            return None
        local = qname[len(_W) :]
        if local not in locals_:
            return None
        found = self._typed.get(id(node))
        if found is not None:
            return found
        from lxml import etree

        from docx4j_py.model.revisions.islands import _to_lxml
        from docx4j_py.runtime import parser
        from docx4j_py.wml import el

        factory = getattr(el, local, None)
        if factory is None:
            return None
        try:
            cls = type(factory())
            data = etree.tostring(_to_lxml(node))
            typed = parser(lenient=True, skipped_report=True).from_bytes(data, cls)
        except Exception:  # noqa: BLE001 - read nothing rather than fail the reading
            return None
        self._typed[id(node)] = typed
        return typed

    # -- markup compatibility: REQ-307 ----------------------------------------------

    def _compare_branches(self, ac: Any, address: str) -> None:
        """``mce.branches_differ`` (REQ-307): the selected branch and another both hold text, and they differ.

        A branch's text is every ``w:t``, ``m:t`` and ``a:t`` in it, in document
        order, in every view, hidden or not (WD 0.29, the user's ruling for
        core-ts CR-012), normalised as REQ-1201 normalises.
        """
        if not self.options.warnings:
            return
        selected = mce_branch(ac)
        if selected is None:
            return
        mine = _branch_text(mce_branch_content(selected))
        if not mine:
            return
        branches = [*mce_choices(ac)]
        fallback = mce_fallback(ac)
        if fallback is not None:
            branches.append(fallback)
        for other in branches:
            if other is selected:
                continue
            theirs = _branch_text(mce_branch_content(other))
            if theirs and theirs != mine:
                if is_any_element(selected):
                    is_choice = selected.qname == MC_CHOICE
                else:
                    is_choice = type(selected).__name__ == "AlternateContentChoice"
                label = f'mc:Choice Requires="{requires_of(selected)}"' if is_choice else "mc:Fallback"
                self.warn(TextWarning(TextWarning.MCE_BRANCHES_DIFFER, address, branch=label))
                return

    def warn(self, warning: TextWarning) -> None:
        """A warning, kept once (REQ-801)."""
        key = (warning.kind, warning.address, warning.branch, warning.content_type, warning.part_name)
        if key in self._warned:
            return
        self._warned.add(key)
        self.warnings.append(warning)

    # -- headers and footers by section (REQ-502) ----------------------------------

    def _sections_with_headers(
        self, sections: list[Any], main_blocks: list[TextBlock], headers: bool, footers: bool, out: list[TextBlock]
    ) -> None:
        settings = self.trees.settings
        even_and_odd = _on(getattr(settings, "even_and_odd_headers", None)) if settings is not None else False
        emitted: set[str] = set()
        last_header: dict[str, str] = {}
        last_footer: dict[str, str] = {}
        b = 0
        for s, sect_pr in enumerate(sections):
            title_pg = False
            if sect_pr is not None:
                for ref in sect_pr.header_reference_or_footer_reference or ():
                    kind = _v(getattr(ref, "type_value", None))
                    rel_id = getattr(ref, "id", None)
                    if kind is None or rel_id is None:
                        continue
                    (last_header if rtrees.name(ref) == "headerReference" else last_footer)[str(kind)] = rel_id
                title_pg = _on(sect_pr.title_pg)
            shown = (["first"] if title_pg else []) + ["default"] + (["even"] if even_and_odd else [])
            if headers:
                for kind in shown:
                    self._header_or_footer(last_header.get(kind), "HEADER", emitted, out)
            while b < len(main_blocks) and main_blocks[b].section <= s:
                out.append(main_blocks[b])
                b += 1
            if s == len(sections) - 1:
                out.extend(main_blocks[b:])
                b = len(main_blocks)
            if footers:
                for kind in shown:
                    self._header_or_footer(last_footer.get(kind), "FOOTER", emitted, out)

    def _header_or_footer(self, rel_id: str | None, kind: str, emitted: set[str], out: list[TextBlock]) -> None:
        if rel_id is None or rel_id in emitted:
            return
        emitted.add(rel_id)
        found = self.trees.header_or_footer(rel_id)
        if found is None:
            return
        part = found[0]
        prefix = f"{'header' if kind == 'HEADER' else 'footer'}:{rel_id}"
        story = _Story(kind, prefix, part, self.context.numbering_states.for_part(part))
        self.blocks(rtrees.block_items(found[1]), prefix, story, out)

    # -- notes (REQ-503) ---------------------------------------------------------------

    def _notes(self, note_type: str, out: list[TextBlock]) -> None:
        holder = self.trees.footnotes if note_type == "footnote" else self.trees.endnotes
        numbers = self.context.footnotes if note_type == "footnote" else self.context.endnotes
        if holder is None or not numbers:
            return
        part, tree = holder
        by_id = {str(note.id): note for note in getattr(tree, note_type, None) or () if note.id is not None}
        state = self.context.numbering_states.for_part(part)
        for note_id, number in numbers.items():
            note = by_id.get(note_id)
            if note is None:
                continue
            kind = _v(getattr(note, "type_value", None))
            if kind is not None and kind != "normal":
                continue
            prefix = f"{note_type}:{note_id}"
            story = _Story("FOOTNOTE" if note_type == "footnote" else "ENDNOTE", prefix, part, state)
            story.note_type = note_type
            story.note_number = number
            story.note_id = note_id
            blocks: list[TextBlock] = []
            self.blocks(note.content, prefix, story, blocks)
            if not blocks:
                continue
            # the note is prefixed [n]: its own mark, or one added at its start
            if not any(r.kind == "NOTE_MARK" for b in blocks for r in b.runs):
                first = next((b for b in blocks if b.kind == "PARAGRAPH"), None)
                if first is not None:
                    mark = TextRun("NOTE_MARK")
                    mark.note_type = note_type
                    mark.note_number = number
                    mark.note_id = note_id
                    first.runs.insert(0, mark)
            for b in blocks:
                if b.note_type is None:
                    b.note_type = note_type
                    b.note_number = number
                    b.note_id = note_id
            out.extend(blocks)

    # -- comments (REQ-504) ---------------------------------------------------------------

    def _comments(self, out: dict[str, TextComment]) -> None:
        holder = self.trees.comments
        if holder is None:
            return
        part, tree = holder
        extended: dict[str, Any] = {}
        ex_tree = self.trees.comments_extended
        for ex in getattr(ex_tree, "comment_ex", None) or ():
            if ex.para_id:
                extended[str(ex.para_id).upper()] = ex
        by_para_id: dict[str, str] = {}
        anchored = self.trees.anchored_comments()
        state = self.context.numbering_states.for_part(part)
        for c in getattr(tree, "comment", None) or ():
            if self.plan is not None and self.plan.removes_comment(c.id):
                continue
            comment_id = "" if c.id is None else str(c.id)
            # a comment with no anchor in the document is none of its text: Word drops it on save (REQ-504)
            if comment_id not in anchored:
                continue
            comment = TextComment(comment_id, author=c.author, initials=c.initials, date=_date_text(c.date))
            prefix = f"comment:{comment_id}"
            self.blocks(c.content, prefix, _Story("COMMENT", prefix, part, state), comment.blocks)
            para_id = _last_para_id(c.content)
            if para_id is not None:
                by_para_id[para_id.upper()] = comment_id
                ex = extended.get(para_id.upper())
                if ex is not None:
                    comment.resolved = ex.done in (True, 1, "1", "true")
                    comment.parent_id = ex.para_id_parent
            out[comment_id] = comment
        for comment in out.values():
            if comment.parent_id is None:
                continue
            parent = by_para_id.get(str(comment.parent_id).upper())
            comment.parent_id = parent
            if parent is not None and parent in out:
                out[parent].replies.append(comment)

    # -- the view ---------------------------------------------------------------

    def removes(self, item: Any) -> bool:
        """Whether the view removes a list member (never in markup, nor on a resolved copy).

        With no package, a run-level revision of the other view's kind.
        """
        if self.plan is not None:
            return self.plan.removes(item)
        if self.planless and self.view != "markup":
            revision = rtrees.revision_of(item)
            if revision is not None:
                inserted = revision in (rtrees.Revision.INS, rtrees.Revision.MOVE_TO)
                return inserted != (self.view == "accepted")
        return False

    def unwraps(self, item: Any) -> bool:
        """Whether the view takes a content control away and leaves its content (an inserted one, rejected)."""
        return self.plan is not None and self.plan.unwraps(item)

    def _joins_next(self, p: Any) -> bool:
        return self.plan is not None and self.plan.removes_mark(p)

    # -- hidden text (REQ-507) and formatting ----------------------------------------

    def _style_rpr_of(self, r_style: str | None, p_style: str | None) -> Any:
        """The effective run properties of a paragraph style and a character style, cached; None without styles."""
        if self.resolver is None:
            return None
        key = (p_style, r_style)
        if key in self._style_rpr:
            return self._style_rpr[key]
        effective = None
        try:
            from docx4j_py import wml

            r_pr = wml.RPr(r_style=wml.RStyle(val=r_style)) if r_style is not None else None
            pp = wml.PPr(p_style=wml.CtPprBasePStyle(val=p_style)) if p_style is not None else None
            effective = self.resolver.get_effective_rpr(r_pr, pp)
        except Exception:  # noqa: BLE001 - not resolved
            effective = None
        self._style_rpr[key] = effective
        return effective

    def _is_hidden_mark(self, p_pr: Any) -> bool:
        """Whether the paragraph mark is hidden: its own ``w:vanish``, else its styles', cached by the two style ids."""
        mark = getattr(p_pr, "r_pr", None) if p_pr is not None else None
        own = getattr(mark, "vanish", None) if mark is not None else None
        if own is not None:
            return _on(own)
        if self.resolver is None:
            return False
        p_style = _v(getattr(getattr(p_pr, "p_style", None), "val", None)) if p_pr is not None else None
        r_style = _v(getattr(getattr(mark, "r_style", None), "val", None)) if mark is not None else None
        key = ("mark", p_style, r_style)
        found = self._style_rpr.get(key)
        if found is None:
            try:
                from docx4j_py import wml

                pp = wml.PPr(
                    p_style=wml.CtPprBasePStyle(val=p_style) if p_style is not None else None,
                    r_pr=wml.ParaRPr(r_style=wml.RStyle(val=r_style)) if r_style is not None else None,
                )
                found = _on(getattr(self.resolver.get_effective_paragraph_mark_rpr(pp), "vanish", None))
            except Exception:  # noqa: BLE001 - not resolved
                found = False
            self._style_rpr[key] = found
        return found

    def _showing_placeholder(self, sdt: Any) -> bool:
        from docx4j_py.wml.sdt import sdt_property

        try:
            flag = sdt_property(getattr(sdt, "sdt_pr", None), "showingPlcHdr", W_NS)
        except Exception:  # noqa: BLE001 - a control read as not showing its placeholder
            return False
        return flag is not None and _on(flag)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _part_name(part: Any) -> str | None:
    name = getattr(part, "part_name", None)
    return None if name is None else str(name)


def _holder_of(typed: Any, child: Any) -> list | None:
    """The list of a typed node that holds a wildcard child (its ``any_element`` field, usually)."""
    import dataclasses

    for f in dataclasses.fields(typed):
        found = getattr(typed, f.name, None)
        if isinstance(found, list) and any(item is child for item in found):
            return found
    return None


def _is_marker(item: Any) -> bool:
    """A range marker, bookmark, permission or proofing mark: a join passes over it."""
    local = rtrees.name(item) if not isinstance(item, str) else ""
    return local.endswith(("RangeStart", "RangeEnd")) or local in (
        "bookmarkStart",
        "bookmarkEnd",
        "permStart",
        "permEnd",
        "proofErr",
    )


def _r_style(r: Any) -> str | None:
    r_pr = r.r_pr
    if r_pr is None or r_pr.r_style is None:
        return None
    return _v(r_pr.r_style.val)


def _on_of(own: Any, styles: Any, name: str) -> bool:
    """An on/off property: the run's own where it has one, else its styles'."""
    flag = getattr(own, name, None) if own is not None else None
    if flag is not None:
        return _on(flag)
    return _on(getattr(styles, name, None)) if styles is not None else False


def _ascii_font(r_pr: Any) -> str | None:
    fonts = getattr(r_pr, "r_fonts", None) if r_pr is not None else None
    return getattr(fonts, "ascii", None) if fonts is not None else None


def _twips(value: Any) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _mark_revisions(p: Any) -> int:
    p_pr = getattr(p, "p_pr", None)
    r_pr = getattr(p_pr, "r_pr", None) if p_pr is not None else None
    if r_pr is None:
        return 0
    return sum(1 for name in ("ins", "del_value", "move_from", "move_to") if getattr(r_pr, name, None) is not None)


def _list_label(result: Any) -> str | None:
    if result is None:
        return None
    label = result.num_string
    if not label:
        label = result.bullet
    if label is None:
        return None
    suff = getattr(getattr(result.lvl, "suff", None), "val", None) if result.lvl is not None else None
    suff = _v(suff)
    suffix = " " if suff == "space" else "" if suff == "nothing" else "\t"
    return label + suffix


def _in_code(story: _Story) -> bool:
    return any(not f.separated for f in story.fields)


def _suppressed(story: _Story) -> bool:
    """Inside a table of contents left out."""
    return any(f.toc for f in story.fields)


def _in_toc(story: _Story) -> bool:
    return any(f.toc for f in story.fields)


def is_toc_field(instr: str) -> bool:
    """A TOC field's instruction: a table of contents, of figures (``\\c``), of anything it collects."""
    tokens = field_tokens(instr)
    return bool(tokens) and tokens[0].upper() == "TOC"


def _remove_field_code(f: _Field, para: _Paragraph) -> None:
    """A table of contents' code, shown under ``field_codes`` until it was known to be one, taken out again."""
    if f.floor_paragraph is not None:
        f.floor_paragraph.merge_floor = f.previous_floor
    if f.code_block is not None:
        runs = f.code_block.runs
        start = min(f.code_start, len(runs))
        runs[start:] = [r for r in runs[start:] if r.kind != "FIELD_CODE"]
    if para.block is not f.code_block:
        para.block.runs[:] = [r for r in para.block.runs if r.kind != "FIELD_CODE"]


def _link_of(para: _Paragraph) -> str | None:
    if para.link is not None:
        return para.link
    for f in para.story.fields:
        if f.separated and f.link is not None:
            return f.link
    return None


_JAVA_WHITESPACE = re.compile(r"[\t\n\x0b\f\r\x1c-\x1f   -  -    　]")


def field_tokens(instr: str | None) -> list[str]:
    """A field instruction's words: a quoted argument is one, its quotes removed and ``\\\\`` and ``\\"`` unescaped."""
    tokens: list[str] = []
    if not instr:
        return tokens
    i, n = 0, len(instr)
    while i < n:
        c = instr[i]
        if _JAVA_WHITESPACE.match(c):
            i += 1
        elif c == '"':
            out = []
            i += 1
            while i < n and instr[i] != '"':
                d = instr[i]
                if d == "\\" and i + 1 < n and instr[i + 1] in ('\\', '"'):
                    i += 1
                    d = instr[i]
                out.append(d)
                i += 1
            i += 1
            tokens.append("".join(out))
        else:
            j = i
            while j < n and not _JAVA_WHITESPACE.match(instr[j]) and instr[j] != '"':
                j += 1
            tokens.append(instr[i:j])
            i = j
    return tokens


def field_hyperlink_target(instr: str) -> str | None:
    """A HYPERLINK field's target (REQ-703): its address, and ``#`` and its ``\\l`` bookmark; None for any other."""
    tokens = field_tokens(instr)
    if not tokens or tokens[0].upper() != "HYPERLINK":
        return None
    address: str | None = None
    bookmark: str | None = None
    i = 1
    while i < len(tokens):
        t = tokens[i]
        if len(t) == 2 and t[0] == "\\":
            switch = t[1].lower()
            if switch in ("l", "o", "t") and i + 1 < len(tokens):
                i += 1
                if switch == "l":
                    bookmark = tokens[i]
        elif address is None:
            address = t
        i += 1
    address = address or None
    bookmark = bookmark or None
    if bookmark is not None:
        return f"#{bookmark}" if address is None else f"{address}#{bookmark}"
    return address


def _branch_text(branch: list) -> str:
    """A branch's text (``w:t``, ``m:t``, ``a:t``), normalised as REQ-1201 normalises. docx4j ``normalized``."""
    out: list[str] = []
    wanted = {f"{_W}t", f"{{{_MATH_NS}}}t", f"{{{_A_NS}}}t"}

    def enter(node: Any) -> bool:
        if is_any_element(node):
            if node.qname in wanted and node.text:
                out.append(node.text)
            return True
        local = rtrees.name(node)
        cls = type(node).__name__
        if local == "t" or cls == "CTText":
            value = getattr(node, "value", None)
            if isinstance(value, str):
                out.append(value)
        elif cls == "CTRegularTextRun":
            value = getattr(node, "t", None)
            if isinstance(value, str):
                out.append(value)
        return True

    for item in branch:
        if not isinstance(item, str):
            rtrees.walk(item, enter)
    return normalize("".join(out))


def _last_para_id(content: list | None) -> str | None:
    """The last paragraph's ``w14:paraId`` in a comment's content."""
    found: str | None = None
    for item in content or ():
        if isinstance(item, rtrees.cls("P")) and getattr(item, "para_id", None):
            found = item.para_id
        elif is_alternate_content(item):
            inner = _last_para_id(mce_branch_content(mce_branch(item)))
            if inner is not None:
                found = inner
    return found


def _delimit(text: str, display: bool) -> str:
    return f"$${text}$$" if display else f"${text}$"


def _math_revision_type(item: Any) -> bool | None:
    """For a ``w:ins`` or ``w:del`` in an equation's run, whether it is an insertion; else None."""
    if isinstance(item, rtrees.cls("CTMathRunTrackChange")):
        return rtrees.name(item) == "ins"
    return None


def _math_revisions(value: Any) -> tuple[TextRevision | None, TextRevision | None]:
    """The first deletion's and the first insertion's revision in an equation."""
    found: list[TextRevision | None] = [None, None]

    def enter(node: Any) -> bool:
        ins = _math_revision_type(node)
        if ins is not None:
            index = 1 if ins else 0
            if found[index] is None:
                found[index] = _revision(node, "INSERTION" if ins else "DELETION")
        return True

    rtrees.walk(value, enter, wildcards=False)
    return found[0], found[1]


def _math_text(value: Any, keep: Any) -> str:
    """An equation's runs' text (``m:t``, ``w:t``), what `keep` refuses left out: REQ-604's text for one not converted."""
    out: list[str] = []
    ctr = rtrees.cls("CTR")
    ctrl_pr = rtrees.cls("CTCtrlPr")

    def item(entry: Any) -> None:
        if type(entry).__name__ == "CTText" or rtrees.name(entry) == "t":
            v = getattr(entry, "value", None)
            if isinstance(v, str):
                out.append(v)

    def visit(node: Any, depth: int) -> None:
        if node is None or depth > 64 or isinstance(node, str):
            return
        if isinstance(node, ctr):
            for entry in node.content:
                if isinstance(entry, str) or not keep(entry):
                    continue
                if _math_revision_type(entry) is not None:
                    for inner in entry.content:
                        item(inner)
                    continue
                item(entry)
            return
        if isinstance(node, ctrl_pr):
            return
        for _qname, child in iter_children(node, mce="resolve"):
            if not keep(child):
                continue
            visit(child, depth + 1)

    visit(value, 0)
    return "".join(out)

