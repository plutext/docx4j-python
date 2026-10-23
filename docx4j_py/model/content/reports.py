"""What the agent surface returns: frozen dataclasses with ``to_dict()``.

CR-003 section 3.1 --- "results are frozen dataclasses with ``to_dict()``, never
tuples or ad-hoc dicts" --- and section 3.4, which says what each one is for:

:class:`Outline`
    what an agent reads first. A 200-page document does not fit in a context
    window; its outline does. ``entries`` carry the address to edit by, the
    style and level to choose from, and the text cut to a budget that says so.
:class:`SearchHit`
    what ``find()`` returns: a match with its address and the characters either
    side, so a server shows an agent *where* the matches are without a second
    call per hit. :meth:`SearchHit.range` converts one back into a
    :class:`~docx4j_py.model.content.range.Range` to edit.
:class:`ChangeReport`
    what every mutating call records, so an agent verifies what it did without
    re-reading the document and a human can be shown a summary. Always recorded
    (decided question 4): ``pkg.last_change`` is the last one and ``pkg.changes``
    accumulates until the caller clears it. It costs one object, one ``datetime``
    and a list append per call.
:class:`TextExcerpt`
    ``get_text(max_chars=)`` returns a ``str``, because that is what a caller
    wants; this is for the caller who needs the flag as well.

``to_dict()`` leaves out what is None or empty, because a tool result is paid
for by the token; ``dataclasses.asdict`` gives the whole shape when that is
wanted (CR-003 section 3.1).
"""

from __future__ import annotations

import contextlib
import dataclasses
import datetime
import json
import re
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from docx4j_py.model.content.addresses import address_of, ordinal_of, para_id_address
from docx4j_py.model.content.text_model import (
    block_children_of,
    cells_of,
    rows_of,
)
from docx4j_py.traversal import element_name
from docx4j_py.wml import P

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "DEFAULT_ENTRY_LIMIT",
    "ChangeRecorder",
    "ChangeReport",
    "Outline",
    "OutlineEntry",
    "OutlineSection",
    "OutlineStats",
    "SearchHit",
    "TextExcerpt",
    "composed_moves",
    "container_prefix",
    "find_in",
    "heading_level_of",
    "outline_of",
    "recording",
    "recording_on",
    "refuse_read_only",
]

#: How many entries an ``outline()`` reports before it says ``truncated``.
#: CR-003 section 3.4 asks for budgets everywhere and section 7 for a stated
#: size; three hundred entries is about 53 KB of JSON at the default
#: ``max_chars``, which is section 7's 64 KB with room to spare. ``limit=None``
#: asks for all of them, and ``stats`` counts the whole document either way.
DEFAULT_ENTRY_LIMIT = 300

#: What a truncated text ends with.
ELLIPSIS = "…"

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_TBL = f"{{{_W}}}tbl"
_W_TR = f"{{{_W}}}tr"
_W_TC = f"{{{_W}}}tc"
_W_SDT = f"{{{_W}}}sdt"
_HEADING = re.compile(r"^heading\s*([1-9])$", re.IGNORECASE)

def _compact(text: str, max_chars: int | None) -> tuple[str, int, bool]:
    """`text` cut to a budget: the cut text, its real length, and whether it was."""
    length = len(text)
    if max_chars is None or length <= max_chars:
        return text, length, False
    return text[:max_chars] + ELLIPSIS, length, True


# ---------------------------------------------------------------------------
# text with a budget
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class TextExcerpt:
    """Text read under a budget, and whether the budget bit.

    ``get_text(max_chars=)`` returns a plain ``str``; this is what
    ``text_budget(max_chars)`` returns for a caller who needs the flag.
    """

    text: str
    """The text, cut to the budget."""
    chars: int
    """How many characters there really are."""
    truncated: bool
    """True when :attr:`text` is shorter than :attr:`chars`."""

    def __str__(self) -> str:
        """The text."""
        return self.text

    def __len__(self) -> int:
        """The length of the text that was returned, not of the whole."""
        return len(self.text)

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return {"text": self.text, "chars": self.chars, "truncated": self.truncated}


# ---------------------------------------------------------------------------
# outline
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class OutlineEntry:
    """One block of a document, as an agent needs to see it."""

    address: str
    """The address to edit by: the paraId when there is one, else the ordinal."""
    para_id: str | None
    """``w14:paraId``, without the ``w14:`` prefix, or None."""
    ordinal: str
    """The ordinal address (``"body/4/0/1/0"``), always."""
    kind: str
    """``"paragraph"``, ``"table"``, ``"control"``, or the element's name."""
    style_id: str | None
    """``w:pStyle`` (docx4j's view), or None for a block that has none."""
    level: int | None
    """1 to 9 for a heading, from ``w:outlineLvl`` or a ``HeadingN`` style."""
    text: str
    """The text model's text (CR-011), cut to ``max_chars``; a table's is its first row."""
    chars: int
    """How many characters the block really has."""
    truncated: bool
    """True when :attr:`text` was cut."""
    rows: int | None = None
    """A table's row count."""
    cols: int | None = None
    """A table's column count, from its first row."""
    children: tuple[OutlineEntry, ...] = ()
    """The blocks inside a table's cells or a content control."""
    textbox: bool = False
    """A text box's paragraph (CR-011 Phase D): listed after its anchor, at the text
    model's ``textbox:`` address (``body/3/textbox:0/1``), which is for reading;
    ``element_at`` does not take it (CR-011 Q9)."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view; what is None or empty is left out."""
        out: dict[str, Any] = {"address": self.address, "ordinal": self.ordinal}
        if self.para_id:
            out["para_id"] = self.para_id
        out["kind"] = self.kind
        if self.style_id:
            out["style_id"] = self.style_id
        if self.level is not None:
            out["level"] = self.level
        out["text"] = self.text
        out["chars"] = self.chars
        if self.truncated:
            out["truncated"] = True
        if self.rows is not None:
            out["rows"] = self.rows
        if self.cols is not None:
            out["cols"] = self.cols
        if self.children:
            out["children"] = [child.to_dict() for child in self.children]
        if self.textbox:
            out["textbox"] = True
        return out


@dataclasses.dataclass(frozen=True, slots=True)
class OutlineStats:
    """What the document holds, counted structurally.

    Counted over everything, not over the entries that fitted the budget, so an
    agent reading a truncated outline still knows how big the document is.
    """

    paragraphs: int = 0
    """The text model's paragraphs (CR-011): tables' and text boxes' included, a joined
    pair one, a hidden or a table-of-contents paragraph none; the notes not."""
    tables: int = 0
    """The text model's tables, nested ones included."""
    words: int = 0
    """Whitespace-separated words of the text model's accepted view."""
    chars: int = 0
    """Characters of the text model's accepted view, a paragraph's text at a time."""
    comments: int = 0
    """``w:comment`` elements in the comments part, if there is one."""
    tracked_changes: int = 0
    """The tracked changes the engine lists, grouped as Word's Review pane groups them.

    A body's ``get_tracked_changes()`` count (CR-010); the package outline's adds
    its headers' and footers', the stories it shows. Notes and comments are
    ``pkg.get_tracked_changes()``'s and ``revisions.report(pkg)``'s.
    """
    skipped: int = 0
    """What lenient parsing dropped (CR-002 section 8). Zero is the norm."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view; every count, including the zeroes."""
        return dataclasses.asdict(self)

    def _add(self, other: OutlineStats) -> OutlineStats:
        return OutlineStats(
            paragraphs=self.paragraphs + other.paragraphs,
            tables=self.tables + other.tables,
            words=self.words + other.words,
            chars=self.chars + other.chars,
            comments=self.comments + other.comments,
            tracked_changes=self.tracked_changes + other.tracked_changes,
            skipped=self.skipped + other.skipped,
        )


@dataclasses.dataclass(frozen=True, slots=True)
class OutlineSection:
    """One header or footer of an outline: its prefix, its part and its blocks."""

    prefix: str
    """The address prefix its blocks are under (``"header:rId8"``)."""
    part_name: str
    """``/word/header1.xml``."""
    entries: tuple[OutlineEntry, ...] = ()
    """Its blocks, under the same budget as the body's."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return {
            "prefix": self.prefix,
            "part_name": self.part_name,
            "entries": [entry.to_dict() for entry in self.entries],
        }


@dataclasses.dataclass(frozen=True, slots=True)
class Outline:
    """What an agent reads first: the document's structure, within a budget."""

    entries: tuple[OutlineEntry, ...] = ()
    """The body's blocks, in document order."""
    headers: tuple[OutlineSection, ...] = ()
    """One per header part; empty for a body's own outline."""
    footers: tuple[OutlineSection, ...] = ()
    """One per footer part."""
    stats: OutlineStats = dataclasses.field(default_factory=OutlineStats)
    """The counts, over the whole document rather than over what fitted."""
    truncated: bool = False
    """True when a text was cut, an entry was dropped, or ``depth`` bit."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a tool returns."""
        out: dict[str, Any] = {"entries": [entry.to_dict() for entry in self.entries]}
        if self.headers:
            out["headers"] = [section.to_dict() for section in self.headers]
        if self.footers:
            out["footers"] = [section.to_dict() for section in self.footers]
        out["stats"] = self.stats.to_dict()
        out["truncated"] = self.truncated
        return out

    def to_json(self, *, indent: int | None = None) -> str:
        """:meth:`to_dict` as JSON, compact by default because it is paid for."""
        separators = (",", ":") if indent is None else None
        return json.dumps(self.to_dict(), indent=indent, separators=separators, ensure_ascii=False)

    def to_markdown(self) -> str:
        """A nested list, headings as headings: the cheapest thing to show a model."""
        lines: list[str] = []

        def render(entries: tuple[OutlineEntry, ...], depth: int) -> None:
            for entry in entries:
                if entry.level is not None and depth == 0:
                    lines.append(f"{'#' * min(entry.level, 6)} {entry.text}".rstrip())
                else:
                    label = entry.text
                    if entry.kind == "table":
                        size = f"table {entry.rows}x{entry.cols}"
                        label = f"{size}: {label}" if label else size
                    elif entry.kind == "control":
                        label = f"control: {label}" if label else "control"
                    lines.append(f"{'  ' * depth}- {label}".rstrip())
                render(entry.children, depth + 1)

        render(self.entries, 0)
        for section in (*self.headers, *self.footers):
            if not section.entries:
                continue
            lines.append(f"<!-- {section.prefix} -->")
            render(section.entries, 0)
        if self.truncated:
            lines.append("- …")
        return "\n".join(lines)

    def addresses(self) -> list[str]:
        """Every address in the outline, in document order. Extension."""
        out: list[str] = []

        def collect(entries: tuple[OutlineEntry, ...]) -> None:
            for entry in entries:
                out.append(entry.address)
                collect(entry.children)

        collect(self.entries)
        for section in (*self.headers, *self.footers):
            collect(section.entries)
        return out


# ---------------------------------------------------------------------------
# building an outline
# ---------------------------------------------------------------------------


def heading_level_of(element: P) -> int | None:
    """A paragraph's heading level: ``w:outlineLvl + 1``, or its ``HeadingN`` style."""
    p_pr = element.p_pr
    if p_pr is not None:
        outline = getattr(p_pr, "outline_lvl", None)
        value = getattr(outline, "val", None) if outline is not None else None
        value = getattr(value, "value", value)
        if value is not None:
            return int(value) + 1
        style = getattr(p_pr, "p_style", None)
        style_id = getattr(style, "val", None) if style is not None else None
        style_id = getattr(style_id, "value", style_id)
        if style_id:
            match = _HEADING.match(str(style_id))
            if match:
                return int(match.group(1))
    return None


def _style_id_of(element: Any) -> str | None:
    p_pr = getattr(element, "p_pr", None)
    style = getattr(p_pr, "p_style", None) if p_pr is not None else None
    if style is None:
        tbl_pr = getattr(element, "tbl_pr", None)
        style = getattr(tbl_pr, "tbl_style", None) if tbl_pr is not None else None
    value = getattr(style, "val", None) if style is not None else None
    value = getattr(value, "value", value)
    return str(value) if value else None


def _kind_of(element: Any) -> str:
    name = element_name(element)
    if isinstance(element, P):
        return "paragraph"
    if name == _W_TBL:
        return "table"
    if name == _W_SDT:
        return "control"
    if not name:
        return type(element).__name__
    return name.rpartition("}")[2]


def _table_shape(element: Any) -> tuple[int, int]:
    """A table's rows and columns (its first row's), in the tree.

    CR-003 section 12.8: the ``Table`` view is slotted in here rather than
    duplicated beside it. ``rows_of`` and ``cells_of`` unwrap a row- or
    cell-level ``w:sdt`` --- the OpenDoPE repeat of section 4 --- so a repeat
    reports the rows it holds rather than counting itself as one. The preview
    is the text model's since CR-011 Phase D (:func:`_preview_of`).
    """
    rows = rows_of(element)
    if not rows:
        return 0, 0
    return len(rows), len(cells_of(rows[0][0]))


def _cells_with_paths(tbl: Any, path: list[int]) -> list[tuple[list[int], Any]]:
    """Every cell of a table with its ordinal path, through what wraps a row or a cell.

    A row- or cell-level content control or custom XML element, and since
    CR-012 Phase A2 a row- or cell-level ``mc:AlternateContent`` kept whole (by
    the branch Word reads), is one index step of the path and no entry of its
    own: each cell's blocks are entries under the table, as ``rows_of`` and
    ``cells_of`` read the rows and cells (``body/1/1/0/1/0`` for a paragraph of
    the second cell of the row in a row-level element's branch).
    """
    out: list[tuple[list[int], Any]] = []

    def rows(items: list, at: list[int]) -> None:
        for index, item in enumerate(items):
            here = [*at, index]
            name = element_name(item)
            if name == _W_TR:
                cells(block_children_of(item) or [], here)
            elif name != _W_TC:
                inner = block_children_of(item)
                if inner is not None:
                    rows(inner, here)

    def cells(items: list, at: list[int]) -> None:
        for index, item in enumerate(items):
            here = [*at, index]
            if element_name(item) == _W_TC:
                out.append((here, item))
            else:
                inner = block_children_of(item)
                if inner is not None:
                    cells(inner, here)

    rows(block_children_of(tbl) or [], path)
    return out


class _Budget:
    """How many entries are left, and whether anything was left out."""

    __slots__ = ("left", "truncated")

    def __init__(self, limit: int | None) -> None:
        self.left = limit
        self.truncated = False

    def take(self) -> bool:
        if self.left is None:
            return True
        if self.left <= 0:
            self.truncated = True
            return False
        self.left -= 1
        return True


class _Model:
    """The text model's reading of a body, indexed for the outline (CR-011 Phase D).

    Each block by the ``id`` of the element it was read from (a joined
    paragraph's first, a table's ``w:tbl``), the text boxes by the model
    address of the paragraph that anchors them, and the counts.
    """

    __slots__ = ("boxes", "by_source", "chars", "paragraphs", "prefix", "tables", "words")

    def __init__(self, body: Body) -> None:
        from docx4j_py.model.content.text_model import reading_of

        reading = reading_of(body)
        self.prefix = reading.prefix
        self.by_source: dict[int, Any] = {}
        self.boxes: dict[str, list[Any]] = {}
        self.paragraphs = self.tables = self.words = self.chars = 0
        for block in reading.result.blocks:
            if block.note_type is not None:
                continue  # a note is not the body's (the main body's reading holds its notes)
            if block.story == "TEXT_BOX" and "/textbox:" in block.address:
                self.boxes.setdefault(block.address.split("/textbox:", 1)[0], []).append(block)
            self._index(block)

    def _index(self, block: Any) -> None:
        if block.source is not None and block.story != "TEXT_BOX":
            self.by_source.setdefault(id(block.source), block)
        if block.kind == "PARAGRAPH":
            self.paragraphs += 1
            text = block.text
            self.chars += len(text)
            self.words += len(text.split())
        elif block.kind == "TABLE":
            self.tables += 1
            for row in block.rows:
                for cell in row.cells:
                    for inner in cell.blocks:
                        self._index(inner)

    def model_address(self, body: Body, ordinal: str) -> str:
        """A content API ordinal as the reading's address (``body/3``; ``content/3`` for a nested body)."""
        return self.prefix + ordinal[len(body.prefix) :]


def _level_of(block: Any) -> int | None:
    """1 to 9 for a heading: the text model's outline level (through the styles), else none."""
    if block is None:
        return None
    level = block.outline_level
    return None if level is None or level > 8 else level + 1


def _preview_of(block: Any) -> str:
    """A table's preview: its first row's cells in the text model's view, `` | `` apart."""
    if not block.rows:
        return ""
    return " | ".join(
        " ".join(_texts_of(cell.blocks)).replace("\n", " ").replace("\t", " ").strip() for cell in block.rows[0].cells
    )


def _texts_of(blocks: list) -> list[str]:
    out: list[str] = []
    for block in blocks:
        if block.kind == "TABLE":
            for row in block.rows:
                for cell in row.cells:
                    out.extend(_texts_of(cell.blocks))
        elif block.kind == "PARAGRAPH":
            out.append(block.text)
    return out


def outline_of(
    body: Body,
    *,
    depth: int | None = None,
    max_chars: int = 80,
    headings_only: bool = False,
    limit: int | None = DEFAULT_ENTRY_LIMIT,
) -> tuple[tuple[OutlineEntry, ...], OutlineStats, bool]:
    """The entries, the counts and the truncation flag for one body.

    Since CR-011 Phase D the entries' texts, levels and previews and the
    counts are the text model's (WD REQ-706), its accepted view with the
    specification's defaults: hidden text left out, a deleted paragraph mark
    joined (the joined text on the first paragraph's entry, the second's
    empty), fields' results, bound values; each text box's paragraphs after
    the paragraph that anchors it (after the table, for one in a table), at
    their ``textbox:`` addresses and flagged (REQ-518, REQ-519; CR-011 Q9).
    The entries are still one per block of the tree, so that every address
    the body has is listed. The counts are over the whole body; the entries
    are what the budget allowed. ``Body.outline`` and ``pkg.outline`` are
    this plus the headers and footers.
    """
    model = _Model(body)
    budget = _Budget(limit)
    entries = _entries(
        body,
        body.content,
        [],
        0,
        depth=depth,
        max_chars=max_chars,
        headings_only=headings_only,
        model=model,
        budget=budget,
        in_table=False,
    )
    stats = OutlineStats(
        paragraphs=model.paragraphs,
        tables=model.tables,
        words=model.words,
        chars=model.chars,
        tracked_changes=_tracked_changes(body),
    )
    return tuple(entries), stats, budget.truncated


def _box_entries(blocks: list, *, max_chars: int, headings_only: bool, budget: _Budget) -> list[OutlineEntry]:
    """A text box's paragraphs as entries, at their model addresses (read only: CR-011 Q9)."""
    out: list[OutlineEntry] = []
    for block in blocks:
        level = _level_of(block)
        if headings_only and level is None:
            continue
        if not budget.take():
            continue
        text, chars, cut = _compact(block.text, max_chars)
        budget.truncated = budget.truncated or cut
        out.append(
            OutlineEntry(
                address=block.address,
                para_id=block.para_id,
                ordinal=block.address,
                kind="paragraph",
                style_id=block.style,
                level=level,
                text=text,
                chars=chars,
                truncated=cut,
                textbox=True,
            )
        )
    return out


def _entries(
    body: Body,
    items: list,
    path: list[int],
    nesting: int,
    *,
    depth: int | None,
    max_chars: int,
    headings_only: bool,
    model: _Model,
    budget: _Budget,
    in_table: bool,
) -> list[OutlineEntry]:
    """The entries of one block list, each followed by the text boxes it anchors."""
    out: list[OutlineEntry] = []
    for index, element in enumerate(items):
        here = [*path, index]
        ordinal = "/".join([body.prefix, *(str(i) for i in here)])
        kind = _kind_of(element)
        block = model.by_source.get(id(element))
        level: int | None = None
        rows: int | None = None
        cols: int | None = None

        if kind == "paragraph":
            raw = block.text if block is not None else ""
            level = _level_of(block)
        elif kind == "table":
            rows, cols = _table_shape(element)
            raw = _preview_of(block) if block is not None else ""
        else:
            raw = ""

        children: list[OutlineEntry] = []
        descend = depth is None or nesting + 1 < depth
        if kind == "table":
            for cell_path, cell in _cells_with_paths(element, here):
                inner = block_children_of(cell)
                if inner is None:
                    continue
                if not descend:
                    budget.truncated = True
                    continue
                children.extend(
                    _entries(
                        body,
                        inner,
                        cell_path,
                        nesting + 1,
                        depth=depth,
                        max_chars=max_chars,
                        headings_only=headings_only,
                        model=model,
                        budget=budget,
                        in_table=True,
                    )
                )
        elif kind != "paragraph":
            inner = block_children_of(element)
            if inner is not None:
                if not descend:
                    budget.truncated = True
                else:
                    children.extend(
                        _entries(
                            body,
                            inner,
                            here,
                            nesting + 1,
                            depth=depth,
                            max_chars=max_chars,
                            headings_only=headings_only,
                            model=model,
                            budget=budget,
                            in_table=in_table,
                        )
                    )

        # the text boxes this block anchors, as the model reads them (REQ-518): after their
        # paragraph, or after the table a paragraph that anchors one is in
        boxes: list[Any] = []
        if model.boxes and not in_table:
            anchor = model.model_address(body, ordinal)
            if kind == "paragraph":
                boxes = model.boxes.get(anchor, [])
            elif kind == "table":
                boxes = [b for key, found in model.boxes.items() if key.startswith(anchor + "/") for b in found]

        if headings_only:
            # a table of contents is flat: the headings, wherever they are
            if level is not None and budget.take():
                text, chars, cut = _compact(raw, max_chars)
                budget.truncated = budget.truncated or cut
                out.append(
                    OutlineEntry(
                        address=_address_for(element, ordinal),
                        para_id=getattr(element, "para_id", None),
                        ordinal=ordinal,
                        kind=kind,
                        style_id=_style_id_of(element),
                        level=level,
                        text=text,
                        chars=chars,
                        truncated=cut,
                    )
                )
            out.extend(children)
            out.extend(_box_entries(boxes, max_chars=max_chars, headings_only=True, budget=budget))
            continue

        if not budget.take():
            out.extend(children)
            continue
        text, chars, cut = _compact(raw, max_chars)
        budget.truncated = budget.truncated or cut
        out.append(
            OutlineEntry(
                address=_address_for(element, ordinal),
                para_id=getattr(element, "para_id", None),
                ordinal=ordinal,
                kind=kind,
                style_id=_style_id_of(element),
                level=level,
                text=text,
                chars=chars,
                truncated=cut,
                rows=rows,
                cols=cols,
                children=tuple(children),
            )
        )
        out.extend(_box_entries(boxes, max_chars=max_chars, headings_only=False, budget=budget))
    return out


def _address_for(element: Any, ordinal: str) -> str:
    para_id = getattr(element, "para_id", None)
    return para_id_address(para_id) if para_id else ordinal


def _tracked_changes(body: Body) -> int:
    """Every tracked change under this body, counted.

    Counted through the same collection
    :meth:`~docx4j_py.model.content.body.Body.get_tracked_changes` hands out
    (CR-003 Phase F), so the number an agent reads in the stats and the list it
    can act on never disagree. It unmarshals nothing beyond this body's own
    part, which was read to make the outline anyway.
    """
    from docx4j_py.model.content.tracked_change import count_tracked_changes

    return count_tracked_changes(body)


def comments_in(package: Any) -> int:
    """``w:comment`` elements in the package's comments part, if it has one.

    Read with lxml from the bytes the part would be saved as, so the part is
    **not** unmarshalled and stays byte for byte (CR-003 section 3.4: the count
    is structural, and nothing more than needed is read).
    """
    main = getattr(package, "main_document_part", None)
    part = getattr(main, "comments_part", None) if main is not None else None
    if part is None:
        return 0
    if getattr(part, "is_unmarshalled", False):
        return len(getattr(part.contents, "comment", None) or ())
    from lxml import etree

    try:
        root = etree.fromstring(part.xml)
    except Exception:  # noqa: BLE001 - an unreadable comments part has no comments
        return 0
    return len(root.findall(f"{{{_W}}}comment"))


def package_outline(
    package: Any,
    *,
    depth: int | None = None,
    max_chars: int = 80,
    headings_only: bool = False,
    limit: int | None = DEFAULT_ENTRY_LIMIT,
) -> Outline:
    """``pkg.outline()``: the body, then every header and footer."""
    from docx4j_py.model.content.addresses import package_bodies

    body = package.body
    entries, stats, truncated = outline_of(
        body, depth=depth, max_chars=max_chars, headings_only=headings_only, limit=limit
    )
    headers: list[OutlineSection] = []
    footers: list[OutlineSection] = []
    for other in package_bodies(package):
        if other.prefix == body.prefix:
            continue
        kind = other.prefix.split(":", 1)[0]
        if kind not in ("header", "footer"):
            continue
        section_entries, section_stats, section_cut = outline_of(
            other, depth=depth, max_chars=max_chars, headings_only=headings_only, limit=limit
        )
        stats = stats._add(section_stats)
        truncated = truncated or section_cut
        section = OutlineSection(
            prefix=other.prefix,
            part_name=str(other.part.part_name) if other.part is not None else "",
            entries=section_entries,
        )
        (headers if kind == "header" else footers).append(section)

    stats = dataclasses.replace(
        stats,
        comments=comments_in(package),
        skipped=len(getattr(package, "skipped", ()) or ()),
    )
    return Outline(
        entries=entries,
        headers=tuple(headers),
        footers=tuple(footers),
        stats=stats,
        truncated=truncated,
    )


# ---------------------------------------------------------------------------
# find
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class SearchHit:
    """One match of ``find()``: where it is, what it is, and what is around it."""

    address: str
    """The paragraph's address: its paraId when it has one, else the ordinal."""
    para_id: str | None
    """``w14:paraId`` without the prefix, or None."""
    ordinal: str
    """The paragraph's ordinal address, always."""
    start: int
    """The offset of the match in the paragraph's accepted text, in code points."""
    end: int
    """One past the last character of the match."""
    match: str
    """The matched text itself."""
    snippet: str
    """:attr:`before` + :attr:`match` + :attr:`after`."""
    before: str
    """Up to ``context`` characters in front of the match."""
    after: str
    """Up to ``context`` characters behind it."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        out: dict[str, Any] = {"address": self.address, "ordinal": self.ordinal}
        if self.para_id:
            out["para_id"] = self.para_id
        out.update(
            {
                "start": self.start,
                "end": self.end,
                "match": self.match,
                "snippet": self.snippet,
                "before": self.before,
                "after": self.after,
            }
        )
        return out

    def range(self, body: Body) -> Range:
        """This hit as a :class:`~docx4j_py.model.content.range.Range`, to edit.

        The paragraph is found again by address, so a hit survives a round trip
        through JSON and a tool call.
        """
        from docx4j_py.model.content.range import Range

        paragraph = body.paragraph_at(self.address)
        return Range(paragraph, self.start, self.end)


def hit_for(paragraph: Paragraph, start: int, end: int, *, context: int) -> SearchHit:
    """One :class:`SearchHit` for a match in a paragraph."""
    text = paragraph.text
    before = text[max(0, start - context) : start]
    after = text[end : end + context]
    match = text[start:end]
    return SearchHit(
        address=paragraph.address,
        para_id=paragraph.para_id,
        ordinal=paragraph.ordinal,
        start=start,
        end=end,
        match=match,
        snippet=before + match + after,
        before=before,
        after=after,
    )


def find_in(
    body: Body,
    text: str,
    *,
    context: int = 40,
    limit: int = 20,
    **options: Any,
) -> list[SearchHit]:
    """``body.find()``: hits with addresses and context, under a budget.

    Office JS's search rules, as ``body.search`` (CR-013 Phase A,
    :func:`~docx4j_py.model.content.text_model.office_matches`).
    """
    from docx4j_py.model.content.text_model import office_matches

    out: list[SearchHit] = []
    for paragraph in body.iter_paragraphs():
        for start, end in office_matches(paragraph.segments(objects=True), text, **options):
            out.append(hit_for(paragraph, start, end, context=context))
            if limit is not None and len(out) >= limit:
                return out
    return out


# ---------------------------------------------------------------------------
# change reports
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class ChangeReport:
    """What one content-API call did. CR-003 section 3.4, decided question 4."""

    operation: str
    """``"insert_paragraph"``, ``"replace_text"``, ``"format"``, ``"delete"``, ..."""
    addresses: tuple[str, ...] = ()
    """The addresses touched: what was made, edited or removed."""
    moved: tuple[tuple[str, str], ...] = ()
    """``(old ordinal, new ordinal)`` for the blocks that shifted."""
    created_para_ids: tuple[str, ...] = ()
    """The ``w14:paraId``s allocated, so an agent can address what it made."""
    text_before: str | None = None
    """For a text edit, what was there."""
    text_after: str | None = None
    """For a text edit, what is there now; for a format, the value written."""
    parts_touched: tuple[str, ...] = ()
    """The parts that will be re-marshalled on the next save."""
    warnings: tuple[str, ...] = ()
    """What the call could not carry and did not do silently (CR-003 Phase K).

    A markdown construct with no Word form, an HTML block skipped on the way
    in, a nested table flattened on the way out: "nothing may be dropped
    silently" is CR-002's rule for loading and this is its counterpart for the
    content API. Empty for every call that dropped nothing.
    """
    at: datetime.datetime = dataclasses.field(
        default_factory=lambda: datetime.datetime.now(datetime.UTC)
    )
    """When, in UTC."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view; what is empty is left out."""
        out: dict[str, Any] = {"operation": self.operation}
        if self.addresses:
            out["addresses"] = list(self.addresses)
        if self.moved:
            out["moved"] = [list(pair) for pair in self.moved]
        if self.created_para_ids:
            out["created_para_ids"] = list(self.created_para_ids)
        if self.text_before is not None:
            out["text_before"] = self.text_before
        if self.text_after is not None:
            out["text_after"] = self.text_after
        if self.parts_touched:
            out["parts_touched"] = list(self.parts_touched)
        if self.warnings:
            out["warnings"] = list(self.warnings)
        out["at"] = self.at.isoformat()
        return out

    def to_json(self, *, indent: int | None = None) -> str:
        """:meth:`to_dict` as JSON."""
        separators = (",", ":") if indent is None else None
        return json.dumps(self.to_dict(), indent=indent, separators=separators, ensure_ascii=False)


class ChangeRecorder:
    """Collects one :class:`ChangeReport` while a verb runs.

    The outermost verb opens one; the primitives it calls add to it, so
    ``insert_paragraph`` reports one change and not three. A body with no
    package records nothing, which is what a detached fragment wants.
    """

    __slots__ = (
        "addresses",
        "body",
        "created_para_ids",
        "moved",
        "operation",
        "parts",
        "text_after",
        "text_before",
        "warnings",
    )

    #: True: this one is collecting. A verb that can compute an address cheaply
    #: --- an insert knows the index it used --- tests this before doing work
    #: the null recorder would throw away.
    active = True

    def __init__(self, operation: str, body: Body | None) -> None:
        """Start a report for one call.

        `body` is None for a call that has no body of its own --- a custom XML
        node write, ``custom_xml_parts.add()`` --- which :class:`recording_on`
        opens; such a call names its parts with :meth:`part` and its addresses
        as strings.
        """
        self.operation = operation
        self.body = body
        self.addresses: list[str] = []
        self.moved: list[tuple[str, str]] = []
        self.created_para_ids: list[str] = []
        self.text_before: str | None = None
        self.text_after: str | None = None
        self.warnings: list[str] = []
        part = getattr(body, "part", None) if body is not None else None
        self.parts: list[str] = [str(part.part_name)] if part is not None else []

    # -- what the primitives call -----------------------------------------

    def touched(self, *targets: Any) -> None:
        """Record the address of something this call made or changed."""
        for target in targets:
            if target is None:
                continue
            if isinstance(target, str):
                if target not in self.addresses:
                    self.addresses.append(target)
                continue
            if self.body is None:
                continue
            try:
                address = address_of(self.body, target)
            except Exception:  # noqa: BLE001, S112 - a block already gone has no address
                continue
            if address not in self.addresses:
                self.addresses.append(address)

    def created(self, para_id: str | None) -> None:
        """Record a ``w14:paraId`` this call allocated."""
        if para_id:
            self.created_para_ids.append(para_id)

    def text(self, before: str | None = None, after: str | None = None) -> None:
        """Record the text before and after a text edit."""
        if before is not None and self.text_before is None:
            self.text_before = before
        if after is not None:
            self.text_after = after

    def shifted(self, pairs: list[tuple[str, str]]) -> None:
        """Record the ordinals that moved."""
        self.moved.extend(pairs)

    def warn(self, message: str) -> None:
        """Record something this call could not carry (CR-003 Phase K)."""
        if message and message not in self.warnings:
            self.warnings.append(message)

    def part(self, *names: Any) -> None:
        """Record a part this call wrote, beyond the body's own.

        What a call that touches a part the body does not own uses: a custom XML
        data part, the properties part beside it, a relationships part
        (CR-003 section 17.9).
        """
        for name in names:
            if name is None:
                continue
            text = str(getattr(name, "part_name", name))
            if text and text not in self.parts:
                self.parts.append(text)

    def finish(self) -> ChangeReport:
        """The frozen report."""
        return ChangeReport(
            operation=self.operation,
            addresses=tuple(self.addresses),
            moved=tuple(self.moved),
            created_para_ids=tuple(self.created_para_ids),
            text_before=self.text_before,
            text_after=self.text_after,
            parts_touched=tuple(self.parts),
            warnings=tuple(self.warnings),
        )


class _NullRecorder:
    """What a nested call, or a body with no package, gets. Every method is a no-op."""

    __slots__ = ()

    #: False: nothing is collecting, so a caller may skip the work entirely.
    active = False

    def touched(self, *targets: Any) -> None:
        """Nothing."""

    def created(self, para_id: str | None) -> None:
        """Nothing."""

    def text(self, before: str | None = None, after: str | None = None) -> None:
        """Nothing."""

    def shifted(self, pairs: list[tuple[str, str]]) -> None:
        """Nothing."""

    def warn(self, message: str) -> None:
        """Nothing."""

    def part(self, *names: Any) -> None:
        """Nothing."""


NULL_RECORDER = _NullRecorder()


def refuse_read_only(body: Body, operation: str) -> None:
    """Raise ``address.read_only`` when `body` is read-only; what :class:`recording` checks first.

    A view over content in a kept ``mc:AlternateContent`` branch has a
    read-only body, where an edit would change one branch alone (CR-012
    Phases A and A2). A setter that edits without opening a report (a row's
    ``is_header``, a cell's ``width``) asks here, and so does a table's verb
    before it edits any of its rows.

    Raises:
        AddressError: `body` is read-only.
    """
    reason = getattr(body, "read_only", None)
    if reason is not None:
        from docx4j_py.model.content.errors import AddressError

        raise AddressError(
            f"{operation} refused: the content is {reason}",
            code="address.read_only",
            hint="read it with outline(), to_markdown() or the views' readers; to edit the branch Word "
            'shows, load with LoadOptions(mc_load_mode="word"), which takes it in place of the element',
        )


class recording:
    """``with recording(body, "insert_paragraph") as change:`` --- the one idiom.

    A context manager rather than a decorator because a verb records what it
    learns *while* it runs (the addresses it made, the ids it allocated) and
    what it knew before (the text it is about to replace). Re-entrant: the
    outermost one owns the report and the inner ones are no-ops, so a verb
    written over other verbs still reports once.
    """

    __slots__ = ("package", "recorder")

    def __init__(self, body: Body, operation: str) -> None:
        """Open a report for `operation` on `body`, unless one is already open.

        Raises:
            AddressError: `body` is read-only (``address.read_only``): the view
                is over content in a kept ``mc:AlternateContent`` branch, where
                an edit would change one branch alone (CR-012 Phase A).
        """
        refuse_read_only(body, operation)
        package = getattr(body, "package", None)
        if package is None or getattr(package, "_current_change", None) is not None:
            self.package = None
            self.recorder: Any = current_recorder(package)
            return
        self.package = package
        self.recorder = ChangeRecorder(operation, body)
        package._current_change = self.recorder

    def __enter__(self) -> Any:
        """The recorder, real or null."""
        return self.recorder

    def __exit__(self, *exc_info: object) -> None:
        """Close the report and put it on the package.

        What the write's deletions took with them --- a note, a comment, a
        relationship whose reference went --- is settled first (CR-013 Phase C;
        :func:`~.span_edits.settle_removals`), and the parts that touches are
        the report's.
        """
        package = self.package
        if package is None:
            return
        package._current_change = None
        if getattr(package, "_pending_removals", None) is not None:
            from docx4j_py.model.content.span_edits import settle_removals

            settle_removals(package, self.recorder)
        _invalidate_list_labels(package)
        if exc_info[0] is None:
            package.changes.append(self.recorder.finish())


class recording_on:
    """``with recording_on(package, "add", part) as change:`` --- for a call with no body.

    :class:`recording` takes the body, which is what almost every verb has; a
    custom XML node write and ``custom_xml_parts.add()`` have a **package** and a
    part instead (CR-003 section 17.9). Re-entrant in exactly the same way ---
    the outermost owns the report --- so a ``fill()`` that writes twelve nodes
    and then applies the bindings still leaves **one** report behind.

    Args:
        package: the package, or a trial one.
        operation: what to call the report.
        *parts: the parts this call writes, as parts or as names.
    """

    __slots__ = ("package", "recorder")

    def __init__(self, package: Any, operation: str, *parts: Any) -> None:
        """Open a report for `operation` on `package`, unless one is already open."""
        if package is None or getattr(package, "_current_change", None) is not None:
            self.package = None
            self.recorder: Any = current_recorder(package)
            self.recorder.part(*parts)
            return
        self.package = package
        self.recorder = ChangeRecorder(operation, None)
        self.recorder.part(*parts)
        package._current_change = self.recorder

    def __enter__(self) -> Any:
        """The recorder, real or null."""
        return self.recorder

    def __exit__(self, *exc_info: object) -> None:
        """Close the report and put it on the package."""
        package = self.package
        if package is None:
            return
        package._current_change = None
        _invalidate_list_labels(package)
        if exc_info[0] is None:
            package.changes.append(self.recorder.finish())


def _invalidate_list_labels(package: Any) -> None:
    """Forget the list labels: one mutation renumbers everything after it.

    The central place CR-003 Phase H asked for. An emulator is only ever there
    once something has asked for a label, so a document with no lists --- and
    every call of every other phase --- pays one attribute lookup.
    """
    if getattr(package, "_numbering_emulator", None) is None:
        return
    from docx4j_py.model.listnumbering import invalidate_labels

    invalidate_labels(package)


def current_recorder(package: Any) -> Any:
    """The report being collected, or the null one. What a primitive adds to."""
    if package is None:
        return NULL_RECORDER
    return getattr(package, "_current_change", None) or NULL_RECORDER


def moved_by_insert(body: Body, container: list, index: int, count: int) -> list[tuple[str, str]]:
    """The ordinals in one container that an insert at `index` shifted.

    Only the container touched, as CR-003 section 3.4 asks: appending --- the
    default location, and the common call --- moves nothing at all, and
    inserting moves the blocks after it by the number inserted.
    """
    if count <= 0 or index >= len(container) - count:
        return []
    prefix = container_prefix(body, container)
    if prefix is None:
        return []
    return [
        (f"{prefix}/{position - count}", f"{prefix}/{position}")
        for position in range(index + count, len(container))
    ]


def moved_by_delete(body: Body, container: list, index: int) -> list[tuple[str, str]]:
    """The ordinals in one container that a delete at `index` will shift.

    Called **before** the delete, so the container still holds the block.
    """
    if index < 0 or index >= len(container) - 1:
        return []
    prefix = container_prefix(body, container)
    if prefix is None:
        return []
    return [
        (f"{prefix}/{position + 1}", f"{prefix}/{position}")
        for position in range(index, len(container) - 1)
    ]


@contextlib.contextmanager
def composed_moves(body: Body | None, *containers: list | None) -> Iterator[None]:
    """One write's ``moved`` as (ordinal before the call, ordinal after), however many blocks it removed or added.

    A write across paragraphs (CR-013 Phase C) removes several blocks and may
    add some, and each primitive records the shift of its own step
    (:func:`moved_by_delete`, :func:`moved_by_insert`) against the state the
    step met, so their concatenation names intermediate states: two pairs
    ``body/3 -> body/2`` for two blocks. Around such a write this snapshots
    each container it may change, drops the steps' pairs, and records the
    composition instead, by element identity: one pair for each direct child
    that is there before and after and whose index changed, under the
    container's prefix before and after. A block removed, or one that ends
    where it started, has no pair, and nor has a new one (it is among the
    addresses). Within a cell or a control the container's prefix numbers its
    children, as :func:`moved_by_delete` does; a descendant of a block that
    moved moves with it, as before.
    """
    package = getattr(body, "package", None) if body is not None else None
    change = current_recorder(package)
    if body is None or not getattr(change, "active", False):
        yield
        return
    snapshots: list[tuple[list, str, list]] = []
    for container in containers:
        if container is None or any(container is seen for seen, _prefix, _before in snapshots):
            continue
        prefix = container_prefix(body, container)
        if prefix is not None:
            snapshots.append((container, prefix, list(container)))
    mark = len(change.moved)
    yield
    del change.moved[mark:]
    for container, prefix, before in snapshots:
        try:
            now = container_prefix(body, container) or prefix
        except Exception:  # noqa: BLE001, S112 - a container whose owner went has no address
            continue
        after = {id(element): index for index, element in enumerate(container)}
        for index, element in enumerate(before):
            moved_to = after.get(id(element))
            if moved_to is not None and moved_to != index:
                change.moved.append((f"{prefix}/{index}", f"{now}/{moved_to}"))


def container_prefix(body: Body, container: list) -> str | None:
    """The address a container's children are numbered under.

    ``"body"`` for the body's own list --- free --- and the container's own
    ordinal otherwise, which costs its nesting depth. With it, a verb that
    knows the index it inserted at has the address without a scan.
    """
    if container is body.content:
        return body.prefix
    owner = getattr(container, "owner", None)
    if owner is None:
        return None
    return ordinal_of(body, owner)
