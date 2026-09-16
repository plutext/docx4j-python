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

import dataclasses
import datetime
import json
import re
from typing import TYPE_CHECKING, Any

from docx4j_py.model.content.addresses import address_of, ordinal_of, para_id_address
from docx4j_py.model.content.text_model import block_children_of, text_of_view
from docx4j_py.traversal import element_name, text_of
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
    "container_prefix",
    "find_in",
    "heading_level_of",
    "outline_of",
    "recording",
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
_W_SDT = f"{{{_W}}}sdt"
_HEADING = re.compile(r"^heading\s*([1-9])$", re.IGNORECASE)

#: The revision elements :class:`OutlineStats` counts. Structural: the element
#: is there or it is not, and nothing else is unmarshalled to find out.
_TRACKED = frozenset(
    {f"{{{_W}}}ins", f"{{{_W}}}del", f"{{{_W}}}rPrChange"}
)


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
    """The text, cut to ``max_chars``; a table's is its first row."""
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
        return out


@dataclasses.dataclass(frozen=True, slots=True)
class OutlineStats:
    """What the document holds, counted structurally.

    Counted over everything, not over the entries that fitted the budget, so an
    agent reading a truncated outline still knows how big the document is.
    """

    paragraphs: int = 0
    """Every ``w:p``, tables and content controls descended into."""
    tables: int = 0
    """Every ``w:tbl``, nested ones included."""
    words: int = 0
    """Whitespace-separated words of the accepted view."""
    chars: int = 0
    """Characters of the accepted view, a paragraph's text at a time."""
    comments: int = 0
    """``w:comment`` elements in the comments part, if there is one."""
    tracked_changes: int = 0
    """``w:ins``, ``w:del`` and ``w:rPrChange`` elements."""
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


def _table_shape(element: Any) -> tuple[int, int, str]:
    """A table's rows, columns and first-row preview."""
    rows = [row for row in (block_children_of(element) or []) if _is_row(row)]
    if not rows:
        return 0, 0, ""
    cells = block_children_of(rows[0]) or []
    preview = " | ".join(text_of(cell).replace("\n", " ").strip() for cell in cells)
    return len(rows), len(cells), preview


def _is_row(element: Any) -> bool:
    name = element_name(element)
    return bool(name) and name.rpartition("}")[2] in ("tr", "sdt", "customXml")


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


class _Counter:
    """The running stats of one body."""

    __slots__ = ("chars", "paragraphs", "tables", "words")

    def __init__(self) -> None:
        self.paragraphs = 0
        self.tables = 0
        self.words = 0
        self.chars = 0


def outline_of(
    body: Body,
    *,
    depth: int | None = None,
    max_chars: int = 80,
    headings_only: bool = False,
    limit: int | None = DEFAULT_ENTRY_LIMIT,
) -> tuple[tuple[OutlineEntry, ...], OutlineStats, bool]:
    """The entries, the counts and the truncation flag for one body.

    The counts are over the whole body; the entries are what the budget allowed.
    ``Body.outline`` and ``pkg.outline`` are this plus the headers and footers.
    """
    counter = _Counter()
    budget = _Budget(limit)
    entries = _entries(
        body,
        body.content,
        [],
        0,
        depth=depth,
        max_chars=max_chars,
        headings_only=headings_only,
        counter=counter,
        budget=budget,
    )
    stats = OutlineStats(
        paragraphs=counter.paragraphs,
        tables=counter.tables,
        words=counter.words,
        chars=counter.chars,
        tracked_changes=_tracked_changes(body),
    )
    return tuple(entries), stats, budget.truncated


def _entries(
    body: Body,
    items: list,
    path: list[int],
    nesting: int,
    *,
    depth: int | None,
    max_chars: int,
    headings_only: bool,
    counter: _Counter,
    budget: _Budget,
) -> list[OutlineEntry]:
    """The entries of one block list, counting everything it walks over."""
    out: list[OutlineEntry] = []
    for index, element in enumerate(items):
        here = [*path, index]
        ordinal = "/".join([body.prefix, *(str(i) for i in here)])
        kind = _kind_of(element)
        level: int | None = None
        rows: int | None = None
        cols: int | None = None

        if kind == "paragraph":
            counter.paragraphs += 1
            raw = text_of_view(element)
            counter.chars += len(raw)
            counter.words += len(raw.split())
            level = heading_level_of(element)
        elif kind == "table":
            counter.tables += 1
            rows, cols, raw = _table_shape(element)
        else:
            raw = ""

        children: list[OutlineEntry] = []
        descend = depth is None or nesting + 1 < depth
        if kind == "table":
            for row_index, row in enumerate(block_children_of(element) or []):
                for cell_index, cell in enumerate(block_children_of(row) or []):
                    inner = block_children_of(cell)
                    if inner is None:
                        continue
                    if not descend:
                        _count_only(inner, counter)
                        budget.truncated = True
                        continue
                    children.extend(
                        _entries(
                            body,
                            inner,
                            [*here, row_index, cell_index],
                            nesting + 1,
                            depth=depth,
                            max_chars=max_chars,
                            headings_only=headings_only,
                            counter=counter,
                            budget=budget,
                        )
                    )
        elif kind != "paragraph":
            inner = block_children_of(element)
            if inner is not None:
                if not descend:
                    _count_only(inner, counter)
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
                            counter=counter,
                            budget=budget,
                        )
                    )

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
    return out


def _address_for(element: Any, ordinal: str) -> str:
    para_id = getattr(element, "para_id", None)
    return para_id_address(para_id) if para_id else ordinal


def _count_only(items: list, counter: _Counter) -> None:
    """Count a subtree the budget refused to list."""
    for element in items:
        if isinstance(element, P):
            counter.paragraphs += 1
            raw = text_of_view(element)
            counter.chars += len(raw)
            counter.words += len(raw.split())
            continue
        if element_name(element) == _W_TBL:
            counter.tables += 1
        inner = block_children_of(element)
        if inner is not None:
            _count_only(inner, counter)


def _tracked_changes(body: Body) -> int:
    """``w:ins``, ``w:del`` and ``w:rPrChange`` under this body, counted."""
    from docx4j_py.traversal import iter_nodes

    count = 0
    for node in iter_nodes(body.container, mce="all"):
        if element_name(node) in _TRACKED:
            count += 1
    return count


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
    """``body.find()``: hits with addresses and context, under a budget."""
    from docx4j_py.model.content.text_model import find_all, search_pattern

    pattern = search_pattern(text, **options)
    out: list[SearchHit] = []
    for paragraph in body.iter_paragraphs():
        for start, end in find_all(paragraph.text, pattern):
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
    )

    #: True: this one is collecting. A verb that can compute an address cheaply
    #: --- an insert knows the index it used --- tests this before doing work
    #: the null recorder would throw away.
    active = True

    def __init__(self, operation: str, body: Body) -> None:
        """Start a report for one call."""
        self.operation = operation
        self.body = body
        self.addresses: list[str] = []
        self.moved: list[tuple[str, str]] = []
        self.created_para_ids: list[str] = []
        self.text_before: str | None = None
        self.text_after: str | None = None
        part = getattr(body, "part", None)
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


NULL_RECORDER = _NullRecorder()


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
        """Open a report for `operation` on `body`, unless one is already open."""
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
        """Close the report and put it on the package."""
        package = self.package
        if package is None:
            return
        package._current_change = None
        if exc_info[0] is None:
            package.changes.append(self.recorder.finish())


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
