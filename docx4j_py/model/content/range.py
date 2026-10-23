"""``Range``: Office JS ``Word.Range``, a span within one paragraph.

CR-003 section 3.2 and decided question 3: the offsets are **code points** into
the paragraph's accepted-view text, and only a ``Range`` carries them. Positions
are recomputed from the tree on every use, so a range stays valid across the
edits made through it; an edit made elsewhere in the same paragraph may shift
it, as it does in Word.

**Word's branch** (CR-011 section 15, Phase F). A run- or paragraph-level
``mc:AlternateContent`` gives its paragraph the text of the branch Word shows,
so a range may lie over it: it reads, but every edit through a span holding
that text, or at an offset strictly inside it, is refused (``address.read_only``),
before anything changes. An insertion at either end of it goes beside the
element.

Two rules of section 4 land here: a ``Range``'s original view is not offered
(``get_text(view="original")`` raises --- a range's offsets are accepted-view
offsets, and the original view has different ones), and :attr:`Range.font`
splits the runs at the boundaries so that a write touches the span and nothing
else. Phase B splits and does not merge back: two adjacent runs whose ``w:rPr``
ends up identical stay two runs, which Word tolerates and a later phase can
tidy.

:attr:`Range.hyperlink` (CR-003 section 21) reads and writes a link in Office
JS's ``address#location`` shape; :mod:`.hyperlinks` holds the rules.
"""

from __future__ import annotations

import math
import sys
from collections.abc import Callable, Collection
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, deep_copy
from docx4j_py.model.content.enums import (
    Location,
    ParagraphLocation,
    RangeLocation,
    TextView,
)
from docx4j_py.model.content.errors import ContentError, SpanError
from docx4j_py.model.content.font import Font, MixedFont
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.story import (
    StoryLayout,
    first_and_last_paragraph,
    story_layout,
    story_of,
    top_index,
)
from docx4j_py.model.content.styles import built_in_of, style_name_of
from docx4j_py.model.content.text_model import (
    RUN_HOLDER_NAMES,
    W_P,
    W_TBL,
    block_children_of,
    canonical_view,
    drop_unwritable,
    office_matches,
    require_view,
)
from docx4j_py.model.content.tracking import REVISION_NAMES
from docx4j_py.traversal import element_name
from docx4j_py.wml import to_xml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = ["Range"]

#: How much of the text a ``repr`` shows.
_PREVIEW = 30


class Range:
    """A subset of Office JS ``Word.Range``: ``[start, end)`` from one paragraph to the same or a later one of its story."""

    __slots__ = ("_end_paragraph", "end", "end_mark", "mark", "paragraph", "start", "whole")

    def __init__(
        self,
        paragraph: Paragraph,
        start: int,
        end: int,
        *,
        end_paragraph: Paragraph | None = None,
        mark: bool = False,
        end_mark: str | None = None,
        whole: bool = False,
    ) -> None:
        """Build the span from `start` in `paragraph` to `end` in `end_paragraph`, in code points.

        `end_paragraph` is the paragraph holding `end` (CR-013 Phase B, D1;
        core-ts ``endParagraph`` at ``ef57729``): `paragraph` when it is None,
        or the same ``w:p``, for a range within one paragraph, which is
        exactly what a range has been since CR-003. A later paragraph makes a
        range across paragraphs, checked here: both in one story, the end's
        not before the start's.

        `mark` is a ``'Whole'`` range's (core-ts's ``mark`` at ``9307bdb``;
        CR-013 Phase A): its :attr:`text` ends with the last paragraph's mark,
        as Office JS's does, or with `end_mark` in its place (a table's
        ``'Whole'`` range's row end, ``\\r\\n``; a row control's ``\\t\\t``).
        The mark is no offset. `whole` is core-ts's flag for a range holding
        its paragraphs' whole content (a paragraph's, a body's, a control's,
        a table's range), kept as core-ts keeps it; a write within one
        paragraph decides what its edges take as Phase A built it (CR-013
        section 10.11 item 2).

        Raises:
            SpanError: ``range.other_story``, `end_paragraph` in another story
                (a note, a comment, a text box, a branch Word does not read);
                ``range.order``, `end_paragraph` before `paragraph`.
        """
        #: The paragraph holding :attr:`start`.
        self.paragraph = paragraph
        #: The offset of the first character, in :attr:`paragraph`'s text.
        self.start = start
        #: One past the last character, in :attr:`end_paragraph`'s text.
        self.end = end
        #: A ``'Whole'`` range: its text ends with the last paragraph's mark (``\\r``, ``\\t`` for a cell's last).
        self.mark = mark
        #: What the text ends with in place of the last paragraph's mark, where :attr:`mark` is set.
        self.end_mark = end_mark
        #: The range holds its paragraphs' whole content (core-ts's ``whole``). Extension.
        self.whole = whole
        self._end_paragraph = (
            end_paragraph
            if end_paragraph is not None and end_paragraph.element is not paragraph.element
            else None
        )
        if self._end_paragraph is not None:
            _check_order(paragraph, self._end_paragraph)

    # -- the two ends (CR-013 Phase B) ---------------------------------------

    @property
    def end_paragraph(self) -> Paragraph:
        """The paragraph holding :attr:`end`: :attr:`paragraph` for a range within one paragraph."""
        return self._end_paragraph if self._end_paragraph is not None else self.paragraph

    @end_paragraph.setter
    def end_paragraph(self, value: Paragraph | None) -> None:
        self._end_paragraph = value if value is not None and value.element is not self.paragraph.element else None

    @property
    def spans_paragraphs(self) -> bool:
        """True for a range across paragraphs: its two ends are in two. Extension (core-ts's ``spansParagraphs``)."""
        end = self._end_paragraph
        return end is not None and end.element is not self.paragraph.element

    @property
    def holds_end_mark(self) -> bool:
        """True where the range holds its last paragraph's mark, or a table's end in its place. Extension.

        A ``'Whole'`` range ending at its last paragraph's end (core-ts's
        ``holdsEndMark``): what a deletion takes and a replacement keeps
        (core-ts check 52, W1 and W2; CR-013 Phase C).
        """
        return self.mark and self.end >= len(self.end_paragraph.text)

    def require_one_paragraph(self, operation: str = "this write") -> None:
        """Refuse, before anything changes, a write that Phase C brings to a range across paragraphs. Extension.

        CR-013 Phase B (section 4.3): every write across paragraphs is refused
        until Phase C, and so are ``get_xml`` and ``to_api_script``, whose one
        ``w:p`` can't hold several blocks. A range within one paragraph that
        holds a whole table (a one-cell table's ``'Whole'`` range, its
        :attr:`end_mark` the row end) is refused too, since its write would
        take the table.

        Raises:
            SpanError: ``range.spans_paragraphs``.
        """
        if self.spans_paragraphs or (self.end_mark is not None and self.holds_end_mark):
            what = "spans paragraphs" if self.spans_paragraphs else "holds a whole table"
            raise SpanError(
                f"{operation} refused: this range {what}, and a write across paragraphs is "
                "CR-013 Phase C's (ranges across paragraphs, the writes)",
                code="range.spans_paragraphs",
                hint="write each paragraph's range (range.paragraphs[i].get_range()), or the "
                "content control, comment or table the range came from",
            )

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """The same span: both paragraphs, both offsets and :attr:`mark`."""
        return (
            isinstance(other, Range)
            and other.paragraph == self.paragraph
            and other.end_paragraph == self.end_paragraph
            and (other.start, other.end, other.mark) == (self.start, self.end, self.mark)
        )

    def __hash__(self) -> int:
        """Hashes by both paragraphs' elements, the two offsets and :attr:`mark`."""
        return hash(
            (id(self.paragraph.element), id(self.end_paragraph.element), self.start, self.end, self.mark)
        )

    def __str__(self) -> str:
        """The span's text."""
        return self.text

    def __repr__(self) -> str:
        """``<Range 12:21 'quick brown'>``; across paragraphs ``<Range 7:body/5:6 'HEAD3B\\rBODY3A'>``."""
        text = self.text
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        if self.spans_paragraphs:
            return f"<Range {self.start}:{self.end_paragraph.ordinal}:{self.end} {preview!r}>"
        return f"<Range {self.start}:{self.end} {preview!r}>"

    def __len__(self) -> int:
        """The number of code points the span covers: ``len(text)`` across paragraphs, ``end - start`` within one."""
        if self.spans_paragraphs:
            return len(self.text)
        return max(0, self.end - self.start)

    # -- the story (CR-013 Phase B, core-ts Range.mts at ef57729) ------------

    def _story(self) -> tuple[list, Any]:
        """The story's list and its owner (:func:`~.story.story_of`)."""
        return story_of(self.paragraph.element)

    def _view_of(self, element: Any, container: list) -> Paragraph:
        """The view of a paragraph of this range's story: this range's own where it is one of its two."""
        if element is self.paragraph.element:
            return self.paragraph
        end = self._end_paragraph
        if end is not None and element is end.element:
            return end
        from docx4j_py.model.content.paragraph import Paragraph

        return Paragraph(element, container, self._body_of(element))

    def _body_of(self, element: Any) -> Any:
        """The body a view of a paragraph between the ends takes: one of the ends', whichever holds it, else the story's."""
        for body in (self.paragraph.parent_body, self.end_paragraph.parent_body):
            if _holds(body, element):
                return body
        return self._story_body()

    def _story_body(self) -> Any:
        """A body over this range's story, for its tables, controls and tracked changes (core-ts ``storyBody``)."""
        _items, owner = self._story()
        for body in (self.paragraph.parent_body, self.end_paragraph.parent_body):
            if body.container is owner or _owner_of(body) is owner:
                return body
        package = getattr(self.paragraph.parent_body, "package", None)
        main = _main_body_of(package)
        if main is not None and (main.container is owner or _owner_of(main) is owner):
            return main
        return self.paragraph.parent_body.sub(owner)

    def _slice(self, *elements: Any) -> tuple[list, Any, int, int]:
        """The story, and the first and last top-level index holding `elements` (this range's ends by default)."""
        items, owner = self._story()
        targets = elements or (self.paragraph.element, self.end_paragraph.element)
        indices = [top_index(element, items, owner) for element in targets]
        if any(index < 0 for index in indices):
            raise ContentError(
                "this range's paragraphs are not in one story as a reader reads it: an edit through "
                "another view removed or moved them",
                code="range.stale",
                hint="take the range again",
            )
        return items, owner, min(indices), max(indices)

    def _located(self) -> tuple[StoryLayout, int, int]:
        """The layout of the top-level blocks this range touches, and the places of its two paragraphs in it."""
        items, _owner, low, high = self._slice()
        layout = story_layout(
            items, lambda element, container: self._view_of(element, container).text, start=low, stop=high + 1
        )
        first = layout.index_of(self.paragraph.element)
        last = layout.index_of(self.end_paragraph.element)
        if first < 0 or last < 0:
            raise ContentError(
                "this range's paragraphs are not in one story as a reader reads it",
                code="range.stale",
                hint="take the range again",
            )
        return layout, first, last

    def _pieces(self) -> list[Range]:
        """The range as one range per paragraph it touches: its part of the first, each between whole, its part of the last."""
        if not self.spans_paragraphs:
            return [self]
        layout, first, last = self._located()
        out: list[Range] = []
        places = layout.places[first : last + 1]
        for index, place in enumerate(places):
            view = self._view_of(place.element, place.container)
            if index == 0:
                out.append(Range(view, self.start, place.end - place.start, whole=self.whole))
            elif index == len(places) - 1:
                out.append(Range(view, 0, self.end, whole=self.whole))
            else:
                out.append(Range(view, 0, place.end - place.start, whole=True))
        return out

    # -- text --------------------------------------------------------------

    @property
    def text(self) -> str:
        """The span's text in Office JS's characters: within one paragraph a slice of its :attr:`Paragraph.text`.

        A paragraph's ``'Whole'`` range (:attr:`mark`) ends with the
        paragraph's mark, ``\\r``, or ``\\t`` for a cell's last paragraph, as
        Office JS's does (core-ts check 48; CR-013 Phase A), or with
        :attr:`end_mark` in its place. **Across paragraphs** (CR-013 Phase B;
        core-ts check 52) the text is ``Body.text``'s form between the two
        ends: a mark ``\\r`` between paragraphs, a table's cells a ``\\t`` apart
        and each row ended ``\\r\\n``, a note reference U+0002, hidden text left
        out, from the story's layout (:func:`~.story.story_layout`).

        Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text. It reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, section 15).
        """
        if not self.spans_paragraphs:
            text = self.paragraph.text
            span = text[self.start : self.end]
            if self.mark and self.end >= len(text):
                return span + (self.end_mark if self.end_mark is not None else self.paragraph.mark_text)
            return span
        layout, first, last = self._located()
        places = layout.places
        out = layout.text[places[first].start + self.start : places[last].start + self.end]
        if self.mark and self.end >= places[last].end - places[last].start:
            out += self.end_mark if self.end_mark is not None else self.end_paragraph.mark_text
        return out

    @text.setter
    def text(self, value: str) -> None:
        self.insert_text(value, location="Replace")

    @property
    def is_empty(self) -> bool:
        """True when the range holds no text: :attr:`text` is ``""`` (Office JS ``isEmpty``)."""
        return self.text == ""

    def get_text(self, *, view: TextView = "accepted") -> str:
        """The span's text. The original view is not offered on a range.

        Raises:
            SpanError: `view` is ``"original"`` or ``"markup"``; read that view
                on the paragraph, whose offsets it belongs to (CR-003 section
                4). 0.1.0 answered ``"markup"`` with the accepted text.
            ContentError: any other view (``text.view_invalid``).
        """
        if view in ("original", "markup"):
            raise SpanError(
                f"a Range's offsets are accepted-view offsets, so it has no {view} view",
                code=f"range.no_{view}_view",
                hint=f"call get_text(view={view!r}) on the paragraph instead",
            )
        view = canonical_view(view)  # type: ignore[assignment]
        require_view(view, ("accepted",))
        return self.text

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Every paragraph the range touches, in document order, a table's cells' among them (Office JS ``paragraphs``).

        CR-013 Phase B (core-ts check 52): within one paragraph, that
        paragraph, as a list.
        """
        if not self.spans_paragraphs:
            return [self.paragraph]
        layout, first, last = self._located()
        return [self._view_of(place.element, place.container) for place in layout.places[first : last + 1]]

    @property
    def runs(self) -> list[Any]:
        """The runs the span covers; a run only partly inside counts; across paragraphs, every paragraph's."""
        if self.spans_paragraphs:
            return [run for piece in self._pieces() for run in piece.runs]
        out: list[Any] = []
        seen: set[int] = set()
        for segment in self.paragraph.segments():
            # an equation is in no run, and its math runs' formatting is their own (CR-013 Phase A)
            if segment.end <= self.start or segment.start >= self.end or segment.math:
                continue
            if id(segment.run) in seen:
                continue
            seen.add(id(segment.run))
            out.append(segment.run)
        return out

    # -- style and formatting ---------------------------------------------

    @property
    def style(self) -> str:
        """The range's style name, as Office JS reads it (core-ts checks 52 and 53; CR-013 Phase B).

        A range holding its paragraphs whole reads their paragraph style
        where they share one, else ``Normal``. Any other range reads its runs'
        character style where they share one (``Emphasis``), ``""`` where
        their character styles differ (a footnote reference among plain text,
        Office JS's ``Other``), and the paragraph style where none has one.
        Within one paragraph too. A paragraph after a nested table in a cell
        reads its own style, where Office JS reads ``Other`` unless it is
        Normal: core-ts CR-002 section 57's departure, taken here (Q6).

        Setting it sets the paragraph's style; across paragraphs it is
        refused (``range.spans_paragraphs``) until Phase C.
        """
        char, paragraphs = self._style_read()
        if char is None:
            return paragraphs(lambda p: p.style)
        if char == "":
            return ""
        return style_name_of(self.paragraph.parent_body.package, char)

    @style.setter
    def style(self, value: str) -> None:
        self.require_one_paragraph("Range.style")
        self.refuse_branch_text("style")
        self.paragraph.style = value

    @property
    def style_id(self) -> str:
        """The paragraph's style id; across paragraphs the first one's. Extension, as on :class:`Paragraph`."""
        return self.paragraph.style_id

    @style_id.setter
    def style_id(self, value: str) -> None:
        self.require_one_paragraph("Range.style_id")
        self.refuse_branch_text("style_id")
        self.paragraph.style_id = value

    @property
    def style_built_in(self) -> str:
        """The built-in name of :attr:`style` (Office JS ``styleBuiltIn``): ``"Other"`` where it is ``""``, ``"Normal"`` where paragraphs differ.

        core-ts's subset declares it read-only; the setter is kept here for a
        range within one paragraph, and sets the paragraph's style (CR-013
        section 4.3).
        """
        char, paragraphs = self._style_read()
        if char is None:
            return paragraphs(lambda p: p.style_built_in)
        if char == "":
            return "Other"
        return built_in_of(char)

    @style_built_in.setter
    def style_built_in(self, value: str) -> None:
        self.require_one_paragraph("Range.style_built_in")
        self.refuse_branch_text("style_built_in")
        self.paragraph.style_built_in = value

    def _style_read(self) -> tuple[str | None, Callable[[Callable[[Paragraph], str]], str]]:
        """What :attr:`style` reads: the runs' shared character style id, ``""`` where they differ, or None for the paragraph style.

        core-ts's ``styleRead`` (check 53's rule).
        """

        def paragraphs(read: Callable[[Paragraph], str]) -> str:
            if not self.spans_paragraphs:
                return read(self.paragraph)
            values = {read(p) for p in self.paragraphs}
            return values.pop() if len(values) == 1 else "Normal"

        if self.start == 0 and self.end >= len(self.end_paragraph.text):
            return None, paragraphs
        ids = {_character_style_of(run) for run in self.runs}
        if not ids:
            return None, paragraphs
        if len(ids) > 1:
            return "", paragraphs
        only = ids.pop()
        return (only or None), paragraphs

    def refuse_branch_text(self, operation: str = "this edit") -> None:
        """Refuse an edit through this span when it touches text read from a kept branch. Extension.

        CR-011 Phase F: the span holds such text, or, empty, lies strictly
        inside it (:meth:`Paragraph.refuse_branch_text`); across paragraphs,
        each paragraph's piece.

        Raises:
            AddressError: ``address.read_only``.
        """
        for piece in self._pieces():
            piece.paragraph.refuse_branch_text(piece.start, piece.end, operation)

    @property
    def font(self) -> Font:
        """The formatting of exactly this span.

        Asking for it splits nothing; a read reports the first run the span
        touches, **effective** as Office JS reports it (CR-002 section 15.1).
        A **write** splits the runs at the boundaries first --- the start
        backwards and the end forwards to the nearest grapheme boundary, so a
        cluster is never cut in two (CR-003 section 3.12) --- so that only the
        span is touched.

        **Across paragraphs** (CR-013 Phase B, Q4) it is a
        :class:`~.font.MixedFont`: each run read in its own paragraph, a
        property where they differ ``None``, ``""`` for the name, the colour
        and the highlight, ``"Mixed"`` for the underline, as Office JS reads
        it (core-ts checks 52 and 53); its writes are Phase C's and refuse.
        """
        return self.get_font()

    def get_font(self, *, direct: bool = False) -> Any:
        """:attr:`font`, or the **direct** formatting when `direct` is given.

        CR-002 section 15.1, as :meth:`Paragraph.get_font` is.
        """
        if self.spans_paragraphs:

            def readings() -> list[dict[str, Any]]:
                out: list[dict[str, Any]] = []
                for piece in self._pieces():
                    paragraph = piece.paragraph
                    for run in piece.runs:
                        font = Font(lambda run=run: [run], context=paragraph.resolution_context, direct=direct)
                        out.append(font._read())
                return out

            return MixedFont(readings, lambda: self.require_one_paragraph("a font write"))

        def holders() -> list[Any]:
            if self.start == self.end:
                return []
            self.start = self.paragraph.split_at(self.start, prefer="back")
            self.end = self.paragraph.split_at(self.end, prefer="forward")
            return self.runs

        return Font(
            holders,
            scope=f"{self.start}:{self.end}",
            record=self.paragraph.formatting,
            tracking=self.paragraph.font_tracking,
            context=self.paragraph.resolution_context,
            direct=direct,
            guard=lambda: self.refuse_branch_text("a font write"),
        )

    # -- the tables and controls it covers (CR-013 Phase B) ------------------

    @property
    def parent_table_cell(self) -> Any:
        """The table cell holding the range's start, or None (Office JS ``parentTableCell``, which throws instead). Extension."""
        return self.paragraph.parent_table_cell

    @property
    def parent_table(self) -> Any:
        """The table holding the range's start, or None (Office JS ``parentTable``, which throws instead). Extension."""
        cell = self.paragraph.parent_table_cell
        return cell.parent_table if cell is not None else None

    @property
    def tables(self) -> list[Any]:
        """The tables in the range (Office JS ``Range.tables``), in document order.

        core-ts's rule (CR-013 Phase B; core-ts check 52): a table the range
        covers some of, unless the range is inside one of its cells, where
        that cell's tables are looked at instead. A range across a table, or
        into or out of one, lists it; a cell's whole range lists none. Over
        the story's own tables, as ``Body.tables`` lists them (a table inside
        a block content control is the control's).
        """
        covered = {id(p.element) for p in self.paragraphs}
        if not covered:
            return []
        items, owner, low, high = self._slice()

        def pick(tables: list[Any]) -> list[Any]:
            out: list[Any] = []
            for table in tables:
                cells = [cell for row in table.rows for cell in row.cells]
                held = [{id(p.element) for p in cell.paragraphs} for cell in cells]
                if not any(covered & paragraphs for paragraphs in held):
                    continue
                home = next((cells[i] for i, paragraphs in enumerate(held) if covered <= paragraphs), None)
                out.extend(pick(home.body.tables) if home is not None else [table])
            return out

        story = self._story_body()
        return pick([t for t in story.tables if low <= top_index(t.element, items, owner) <= high])

    @property
    def content_controls(self) -> list[Any]:
        """The content controls in the range (Office JS ``Range.contentControls``), in document order.

        core-ts's rule (CR-013 Phase B; core-ts check 53): each control the
        range overlaps, except one holding the whole range. A range from a
        paragraph into a block control lists it, one inside a control's
        content doesn't, one over an inline control lists it. Every kind is
        listed, where Word's collection lists rich-text controls alone
        (CR-013 section 11, G1's recorded departure);
        :meth:`get_content_controls` filters by type.
        """
        if not self.spans_paragraphs:
            return self._run_controls_overlapping()
        from docx4j_py.model.content.controls import ContentControl, _collect

        items, _owner, low, high = self._slice()
        layout = story_layout(
            items, lambda element, container: self._view_of(element, container).text, start=low, stop=high + 1
        )

        def at(paragraph: Paragraph, offset: int) -> int | None:
            index = layout.index_of(paragraph.element)
            return None if index < 0 else layout.places[index].start + offset

        start = at(self.paragraph, self.start)
        end = at(self.end_paragraph, self.end)
        if start is None or end is None:
            return []
        story = self._story_body()
        found: list[ContentControl] = []
        _collect(items, story, found, low, high + 1)
        out = []
        for control in found:
            if control.form == "run":
                span = control._run_span()
                if span is None:
                    continue
                paragraph, a, b = span
                low_at, high_at = at(paragraph, a), at(paragraph, b)
            else:
                paragraphs = control.paragraphs
                if not paragraphs:
                    continue
                last = paragraphs[-1]
                low_at, high_at = at(paragraphs[0], 0), at(last, len(last.text))
            if low_at is None or high_at is None:
                continue
            overlaps = (low_at < end and high_at > start) or (low_at == high_at and start <= low_at <= end)
            holds = low_at <= start and end <= high_at
            if overlaps and not holds:
                out.append(control)
        return out

    def _run_controls_overlapping(self) -> list[Any]:
        """Within one paragraph: each run-level control whose text overlaps the span, not one holding it whole."""
        overlapping = []
        for control in self.paragraph.content_controls:
            span = control._run_span()
            if span is None or span[0].element is not self.paragraph.element:
                continue
            _paragraph, start, end = span
            overlaps = (start < self.end and end > self.start) or (start == end and self.start <= start <= self.end)
            holds = start <= self.start and self.end <= end
            if overlaps and not holds:
                overlapping.append(control)
        return overlapping


    # -- spans that cross a run holder (CR-003 section 3.4) ----------------

    def holder_boundaries(
        self, *, revisions: bool = False, unwrapped: Collection[Any] = ()
    ) -> list[int]:
        """The offsets at which this span enters or leaves a run holder.

        A ``w:hyperlink``, a run-level ``w:sdt``, a ``w:ins`` and the other run
        holders own their runs; an operation that has to sit in exactly one of
        them --- a content control at range level (Phase C), a comment anchor
        (Phase G) --- cannot span two. These are the offsets to split at.

        Args:
            revisions: count the four revision holders (``w:ins``, ``w:del``,
                ``w:moveFrom``, ``w:moveTo``) as holders too, for an operation
                that wraps the runs it covers in a new holder of its own
                (:attr:`hyperlink`, CR-003 section 21).
            unwrapped: holders the operation removes before it wraps; a run in
                one counts as being where the holder is.

        Raises:
            SpanError: ``range.spans_paragraphs``, a range across paragraphs,
                whose holders are each paragraph's (CR-013 Phase B).
        """
        self.require_one_paragraph("holder_boundaries")
        seen_through = {id(holder) for holder in unwrapped}
        out: list[int] = []
        previous: Any = None
        for segment in self.paragraph.segments():
            if segment.end <= self.start or segment.start >= self.end:
                continue
            holder = getattr(segment.run, "parent", None)
            while holder is not None and id(holder) in seen_through:
                holder = getattr(holder, "parent", None)
            name = element_name(holder)
            owns = name in RUN_HOLDER_NAMES or (revisions and name in REVISION_NAMES)
            key = id(holder) if owns else None
            if previous is not None and key != previous[0]:
                out.append(max(self.start, segment.start))
            previous = (key, holder)
        return out

    def require_one_holder(
        self,
        operation: str = "this operation",
        *,
        revisions: bool = False,
        unwrapped: Collection[Any] = (),
    ) -> None:
        """Refuse a span that crosses a run holder, saying where to split.

        `revisions` and `unwrapped` are :meth:`holder_boundaries`'.

        Raises:
            SpanError: the span covers runs in more than one holder; the message
                names the offsets to split at. ``range.spans_paragraphs`` for a
                range across paragraphs (CR-013 Phase B).
        """
        self.require_one_paragraph(operation)
        boundaries = self.holder_boundaries(revisions=revisions, unwrapped=unwrapped)
        if not boundaries:
            return
        where = ", ".join(str(offset) for offset in boundaries)
        raise SpanError(
            f"{operation} needs a span inside one run holder, and "
            f"{self.start}:{self.end} crosses one at {where}",
            code="range.crosses_holder",
            hint=f"split the span at {boundaries[0]}, or use get_range() on each part",
        )

    # -- hyperlinks (CR-003 section 21) -------------------------------------

    @property
    def hyperlink(self) -> str:
        """Office JS ``Range.hyperlink``: the link this span is in, as ``address#location``.

        Read: the first ``w:hyperlink`` the span touches --- its relationship's
        target for an ``r:id``, ``#`` and the ``w:anchor`` for a location in this
        document, both for a link that has both --- or ``""`` where there is
        none. An empty range reads the link it lies in, or ends at the end of.

        Written: the links the span touches are removed and its runs wrapped in
        one new ``w:hyperlink``, split at the span's boundaries as :attr:`font`
        splits them. The address before any ``#`` becomes an external
        relationship of **this span's part** (a header's paragraph relates from
        the header part) and the ``r:id``; what follows the ``#`` becomes the
        ``w:anchor``. The runs take the ``Hyperlink`` character style in place
        of their own and lose their direct colour, as Word styles a link it
        makes, and the style's definition is added to ``styles.xml`` when the
        document lacks it. A span that crosses a link's edge takes that whole
        link in; one inside a link replaces it with a link over the span alone.
        ``""`` (or None) removes the links the span touches, their runs staying
        and the style coming off. Records a ``"hyperlink"``
        :class:`~docx4j_py.model.content.reports.ChangeReport` naming the
        relationships part it wrote; under change tracking the restyle is a
        ``w:rPrChange``. The rules, and where each is Word's, are in
        :mod:`docx4j_py.model.content.hyperlinks`.

        **Across paragraphs** (CR-013 Phase B) the read is the first link among
        its paragraphs' pieces, as core-ts's; the write is Phase C's and
        refuses.

        Raises:
            SpanError: on a write, the span is empty --- Office JS refuses it
                --- or would put the link across a content control, a field or
                a tracked change; nothing has changed. ``range.spans_paragraphs``
                across paragraphs.
        """
        from docx4j_py.model.content.hyperlinks import hyperlink_of

        if self.spans_paragraphs:
            return next((link for piece in self._pieces() if (link := hyperlink_of(piece))), "")
        return hyperlink_of(self)

    @hyperlink.setter
    def hyperlink(self, value: str | None) -> None:
        from docx4j_py.model.content.hyperlinks import set_hyperlink

        self.require_one_paragraph("Range.hyperlink")
        set_hyperlink(self, value)

    # -- editing -----------------------------------------------------------

    def insert_content_control(self, kind: str = "RichText") -> Any:
        """Wrap the runs this span covers in a run-level content control.

        Office JS ``Range.insertContentControl`` (CR-003 section 3.7). The runs
        are split at the span's boundaries first, on grapheme boundaries, so the
        control holds exactly this span's text.

        Raises:
            SpanError: the span crosses a run holder; the message names where
                to split it.
            BindingError: ``"RepeatingSection"`` was asked for, which Word has
                no run-level form of.
            SpanError: ``range.spans_paragraphs``, a range across paragraphs
                (core-ts's phase 3, not ported: CR-013 Q10).
        """
        from docx4j_py.model.customxml.insert import insert_content_control_in_range

        self.require_one_paragraph("insert_content_control")
        return insert_content_control_in_range(self, kind)

    def insert_text(self, text: str, *, location: Location = "Replace") -> Range:
        """Insert text relative to the span.

        Args:
            text: what to insert.
            location: ``"Replace"`` (the default, as Office JS has it for a
                range), ``"Before"`` / ``"Start"`` for in front of the span, or
                ``"After"`` / ``"End"`` for behind it.

        ``\\n``, ``\\r`` and ``\\r\\n`` each make a new paragraph, as
        Word's ``insertText`` does (CR-013 section 11, E1;
        :meth:`Paragraph.splice`); U+0002 and the other characters XML 1.0
        excludes are dropped (E3).

        Returns:
            The range of the inserted text; where it made paragraphs, a range
            across them, as Office JS's is (core-ts check 55: ``x\\ry``; CR-013
            Phase B). This range goes on covering the same text: a replacement
            makes it the new text's range, text before it shifts it, and text
            after it leaves it where it was, in the paragraph that now holds it.

        Raises:
            SpanError: on a range across paragraphs (CR-013 Phase B; the
                writes are Phase C's): ``range.partial_table`` for a
                replacement starting or ending inside a table, as Word refuses
                it (core-ts check 52); ``range.partial_control`` inside a
                content control; ``range.not_trackable`` under change tracking
                for one holding a whole block that is neither a paragraph nor a
                table; ``range.spans_paragraphs`` otherwise.
        """
        if location == "Replace" and (self.spans_paragraphs or self.end_mark is not None):
            self._refuse_span_edit("insert_text")
        self.require_one_paragraph("insert_text")
        if location in ("Replace", "Before", "Start", "After", "End"):
            text = drop_unwritable(text)
            with recording(self.paragraph.parent_body, "insert_text") as change:
                change.touched(self.paragraph)
                change.text(before=self.text)
                paragraph = self.paragraph
                container = paragraph.container
                index = next((i for i, item in enumerate(container) if item is paragraph.element), -1)
                length = len(paragraph.text)
                if location == "Replace":
                    # a paragraph's 'Whole' range holds its hidden runs and pictures at its edges
                    out = paragraph.splice(self.start, self.end, text, whole=self.mark)
                    # the range covers the new text, across the paragraphs it made (Phase B)
                    self.paragraph = out.paragraph
                    self.start, self.end = out.start, out.end
                    self.end_paragraph = out.end_paragraph
                elif location in ("Before", "Start"):
                    # at the paragraph's start, before its leading hidden runs (CR-013 section 11)
                    out = paragraph.splice(self.start, self.start, text, edge="start" if self.start == 0 else None)
                    # the span now follows the text, in the paragraph keeping the mark
                    self.start, self.end = out.end, out.end + (self.end - self.start)
                else:
                    # at the paragraph's end, after its trailing hidden runs (CR-013 section 11)
                    out = paragraph.splice(self.end, self.end, text, edge="end" if self.end == length else None)
                    if 0 <= index < len(container) and container[index] is not paragraph.element:
                        # the span is before the text: in the first paragraph the split made
                        from docx4j_py.model.content.paragraph import Paragraph

                        self.paragraph = Paragraph(container[index], container, paragraph.parent_body)
                change.text(after=text)
                return out
        raise ContentError(
            f"insert_text on a range takes Replace, Before, After, Start or End, "
            f"not {location!r}",
            code="location.invalid",
            hint="Replace is the default; Before and After put text around the span",
        )

    def insert_paragraph(
        self,
        text: str = "",
        *,
        location: ParagraphLocation = "After",
        style: str | None = None,
    ) -> Paragraph:
        """A new paragraph before or after this span's paragraph.

        Raises:
            SpanError: ``range.spans_paragraphs``, a range across paragraphs,
                whose split is CR-013 Phase C's.
        """
        self.require_one_paragraph("insert_paragraph")
        return self.paragraph.insert_paragraph(text, location=location, style=style)

    def insert_ooxml(self, ooxml: str, *, location: str = "Replace") -> list[Any]:
        """Word's ``insertOoxml`` around or in place of this span (CR-003 section 3.2).

        A flat OPC ``pkg:package`` or a bare fragment, as
        :meth:`Body.insert_ooxml`. A fragment of exactly one ``w:p`` has its
        **runs merged into this paragraph** at the span's start or end, as
        Word's paste does; anything with a block in it goes before or after the
        paragraph.

        Args:
            ooxml: the ``pkg:package`` document, or the fragment.
            location: ``"Replace"`` (the default, as a range's other verbs),
                ``"Before"`` or ``"After"``.

        Returns:
            The views of what was inserted --- ``[self.paragraph]`` when the
            runs were merged in.

        Raises:
            SpanError: ``range.spans_paragraphs``, a range across paragraphs
                (CR-013 Phase B).
        """
        from docx4j_py.model.content.ooxml import insert_ooxml_into_range

        self.require_one_paragraph("insert_ooxml")
        return insert_ooxml_into_range(self, ooxml, location=location)

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` within this span, as ranges of the paragraph.

        Matches are found in Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, which reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, section 15).

        Across paragraphs (CR-013 Phase B) a paragraph at a time, from
        :attr:`start` in the first to :attr:`end` in the last, as core-ts's;
        Word's ``^p`` across marks is core-ts's phase 3, not ported (Q10).
        """
        limit = options.pop("limit", None)
        if self.spans_paragraphs:
            out = [hit for piece in self._pieces() for hit in piece.search(text, **options)]
            return out if limit is None else out[:limit]
        hits = office_matches(
            self.paragraph.segments(objects=True), text, start=self.start, end=self.end, **options
        )
        if limit is not None:
            hits = hits[:limit]
        return [Range(self.paragraph, start, end) for start, end in hits]

    def replace_text(self, find: str, replace: str, **options: Any) -> int:
        """Replace every match within the span, last first; returns the count.

        **Tracked** when the package's ``change_tracking_mode`` is on, in which
        case each replacement is a ``w:del`` and a ``w:ins``. The report names
        the pair, as :meth:`Body.replace_text`'s does.

        Raises:
            SpanError: ``range.spans_paragraphs``, a range across paragraphs
                (a paragraph at a time is CR-013 Phase C's).
        """
        from docx4j_py.model.content.paragraph import replace_matches

        self.require_one_paragraph("replace_text")
        replace = drop_unwritable(replace)
        with recording(self.paragraph.parent_body, "replace_text") as change:
            change.text(before=find)
            matches = self.search(find, **options)
            for match in matches:  # every match checked before the first is replaced (CR-011 Phase F)
                match.refuse_branch_text("replace_text")
                match.paragraph.refuse_equation(match.start, match.end)
                match.paragraph.refuse_paragraph_breaks(match.start, match.end, replace, "replace_text")
            # a replacement holding a paragraph break splits the paragraph (CR-013 section 11, E1)
            replace_matches(matches, replace)
            if matches:
                change.touched(self.paragraph)
            change.text(after=replace)
            return len(matches)

    def find(self, text: str, *, context: int = 40, limit: int = 20, **options: Any) -> list[Any]:
        """Every match within this span, as hits with addresses and context.

        Matches are found in Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, which reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, section 15).
        """
        from docx4j_py.model.content.reports import hit_for

        if self.spans_paragraphs:
            found = [
                hit
                for piece in self._pieces()
                for hit in piece.find(text, context=context, limit=limit, **options)
            ]
            return found[:limit]
        hits = office_matches(
            self.paragraph.segments(objects=True), text, start=self.start, end=self.end, **options
        )
        return [hit_for(self.paragraph, start, end, context=context) for start, end in hits[:limit]]

    # -- content controls (CR-013 section 11, G1) ---------------------------

    def get_content_controls(self, *, types: Collection[str] | None = None) -> list[Any]:
        """Office JS ``getContentControls`` (WordApi 1.5): :attr:`content_controls`, those of `types` where given.

        `types` are ``Word.ContentControlType`` values, as a control's ``type``
        reports them. Word's ``contentControls`` lists rich-text controls only,
        where this lists every kind, a recorded departure (CR-013 section 11,
        G1). core-ts's at ``6285dfa``: ``controlsOfTypes(this.contentControls)``.
        """
        from docx4j_py.model.content.controls import controls_of_types

        return controls_of_types(self.content_controls, types)

    # -- change tracking (CR-003 section 3.8, Phase F) ----------------------

    def get_tracked_changes(self) -> list[Any]:
        """The tracked changes this span covers, in document order (Office JS).

        Within one paragraph, the paragraph's listing, filtered to those whose
        range overlaps this span; a deletion, whose range in the accepted text
        is collapsed, counts when it sits inside the span, and a change that
        starts or ends in another paragraph covers this one from where it
        starts in it to where it ends in it. **Across paragraphs** (CR-013
        Phase B; core-ts check 52: a deletion through a mark is one change)
        the story's listing filtered by overlap, since a paragraph's own
        listing cuts a change at its edge.
        """
        if not self.spans_paragraphs:
            out: list[Any] = []
            for change in self.paragraph.get_tracked_changes():
                span = change.get_range()
                if span is None:
                    continue
                start = span.start if span.paragraph.element is self.paragraph.element else 0
                end = span.end if span.end_paragraph.element is self.paragraph.element else sys.maxsize
                if start <= self.end and end >= self.start:
                    out.append(change)
            return out
        items, owner, low, high = self._slice()
        layout = story_layout(
            items, lambda element, container: self._view_of(element, container).text, start=low, stop=high + 1
        )
        tops: dict[int, int] = {}

        def top(element: Any) -> int:
            found = tops.get(id(element))
            if found is None:
                found = tops[id(element)] = top_index(element, items, owner)
            return found

        def at(paragraph: Any, offset: int) -> float:
            index = layout.index_of(paragraph.element)
            if index >= 0:
                return layout.places[index].start + offset
            where = top(paragraph.element)
            if 0 <= where < low:
                return float("-inf")
            return float("inf") if where > high else float("nan")

        first = layout.places[layout.index_of(self.paragraph.element)].start + self.start
        last = layout.places[layout.index_of(self.end_paragraph.element)].start + self.end
        out = []
        for change in self._story_body().get_tracked_changes():
            paras = [piece.para.p for piece in change.change.pieces if piece.para is not None]
            if paras and all(0 <= top(p) < low for p in paras):
                continue
            if paras and all(top(p) > high for p in paras):
                continue
            span = change.get_range()
            if span is None:
                continue
            start = at(span.paragraph, span.start)
            end = at(span.end_paragraph, span.end)
            if math.isnan(start) or math.isnan(end):  # not in this story as a reader reads it
                continue
            if start <= last and end >= first:
                out.append(change)
        return out

    # -- comments (CR-003 section 3.9, Phase G) -----------------------------

    def get_comments(self) -> list[Any]:
        """The comments this span touches (Office JS ``getComments``), replies nested under their parent.

        Each comment's extent over the story, from its first marker to its
        last, a block-level marker at the next or the previous paragraph's
        edge, overlapping the span, both edges inclusive (CR-013 Phase B;
        section 2.5 item 1: a range on a comment's last paragraph lists it,
        though that paragraph holds only its end marker).
        """
        from docx4j_py.model.content.comments import comments_of

        return comments_of(self)

    def insert_comment(self, text: str) -> Any:
        """Comment on this span (Office JS ``Range.insertComment``).

        The runs are split at the span's boundaries, a ``w:commentRangeStart``
        and a ``w:commentRangeEnd`` go around them and a reference run in the
        ``CommentReference`` style goes after the end; the comment itself
        becomes a ``w:comment`` whose paragraphs are in the ``CommentText``
        style, opened by a ``w:annotationRef`` run and carrying a fresh
        ``w14:paraId``. The author is ``pkg.author`` (CR-003 section 3.4's
        audit trail).

        A comment is not a revision (CR-003 section 4), and on text in a
        ``w:ins`` the markers are written in **Word's own form** (CR-010 section
        12.8; docx4j-core-ts CR-002 section 38): the ``w:commentRangeStart``
        inside the ``w:ins``, the ``w:commentRangeEnd`` and the reference run
        after it, the ``w:ins`` split there --- its rest a second ``w:ins`` with a
        new id --- when the span ends inside it. The anchor stays narrow and the
        reference run is no insertion. Accepting the insertion leaves the
        comment on its text; rejecting it removes the comment, as current Word
        does. With tracking on or off alike. A ``w:del`` or a move keeps the
        markers outside it.

        **What it may touch beyond this part**: ``/word/comments.xml``,
        ``/word/commentsExtended.xml``, ``/word/commentsIds.xml`` and
        ``/word/people.xml``, each **created** with its relationship and content
        type when the document has none; ``/word/commentsExtensible.xml`` is
        only kept in step, never created; and ``/word/styles.xml``, when the
        document does not define the comment styles. Every one of them is listed
        in the report's ``parts_touched``. **In a**
        :meth:`~docx4j_py.model.content.trial.dry_run`, a part created here is
        added to the real package and taken away again when the block ends
        (CR-003 section 14.4).

        Raises:
            SpanError: the span crosses a run holder --- a ``w:hyperlink``, a
                run-level ``w:sdt`` --- and the message says where to split.
            ContentError: this body's package has no main document part.
            SpanError: ``range.spans_paragraphs``, a range across paragraphs
                (core-ts's phase 3, not ported: CR-013 Q10).

        Returns:
            The new :class:`~docx4j_py.model.content.comments.Comment`.
        """
        from docx4j_py.model.content.comments import insert_comment_into

        self.require_one_paragraph("insert_comment")
        return insert_comment_into(self, text)

    def delete(self) -> None:
        """Remove the span's text, leaving an empty range where it was.

        The pictures strictly inside the span go with it, and a paragraph's
        whole range (``'Whole'``, or a span of all its text) takes those at its
        edges too, as core-ts's CR-002 section 48 has it; under change tracking
        a picture's run goes into the ``w:del`` (CR-013 section 11).

        Raises:
            SpanError: on a range across paragraphs (CR-013 Phase B; the
                deletion is Phase C's): ``range.partial_table`` for one
                starting or ending inside a table, as Word refuses it (core-ts
                check 52); ``range.partial_control`` inside a content control;
                ``range.not_trackable`` under change tracking for one holding a
                whole block that is neither a paragraph nor a table;
                ``range.spans_paragraphs`` otherwise.
        """
        if self.spans_paragraphs or self.end_mark is not None:
            self._refuse_span_edit("delete")
        self.require_one_paragraph("delete")
        with recording(self.paragraph.parent_body, "delete") as change:
            change.touched(self.paragraph)
            change.text(before=self.text, after="")
            self.paragraph.splice(self.start, self.end, "", whole=self.mark)
            self.end = self.start

    def get_range(self, location: RangeLocation = "Whole") -> Range:
        """A range of this one, as Office JS's ``getRange`` (core-ts check 52; CR-013 Phase B).

        ``"Whole"`` is the range itself, not widened, keeping a last mark
        where this range holds it (:attr:`mark`); ``"Content"`` the same
        without the mark; ``"Start"`` and ``"End"`` collapsed at its two ends;
        ``"After"`` collapsed just after it --- at the start of the next
        paragraph where the range holds its last paragraph's mark, else at
        its end.
        """
        if location == "Start":
            return Range(self.paragraph, self.start, self.start)
        if location == "End":
            return Range(self.end_paragraph, self.end, self.end)
        if location == "After":
            if self.holds_end_mark:
                following = self._paragraph_after(self.end_paragraph)
                if following is not None:
                    return Range(following, 0, 0)
            return Range(self.end_paragraph, self.end, self.end)
        return Range(
            self.paragraph,
            self.start,
            self.end,
            end_paragraph=self._end_paragraph,
            mark=self.mark and location == "Whole",
            end_mark=self.end_mark,
            whole=self.whole,
        )

    def _paragraph_after(self, paragraph: Paragraph) -> Paragraph | None:
        """The paragraph after `paragraph` in its story as a reader reads it, or None."""
        items, owner = story_of(paragraph.element)
        index = top_index(paragraph.element, items, owner)
        if index < 0:
            return None
        for stop in range(index + 1, len(items) + 1):
            layout = story_layout(items, lambda _element, _container: "", start=index, stop=stop)
            at = layout.index_of(paragraph.element)
            if 0 <= at < len(layout.places) - 1:
                place = layout.places[at + 1]
                return self._view_of(place.element, place.container)
        return None

    def expand_to(self, other: Range) -> Range:
        """Office JS ``expandTo``: from whichever start comes first to whichever end comes last.

        CR-013 Phase B (core-ts check 52: which is first doesn't matter). An
        end holding its paragraph's mark counts as after the same offset
        without it. The two must be in one story. Neither range changes.

        Raises:
            SpanError: ``range.other_story``, the two in different stories (a
                note, a comment, a text box, a branch Word does not read), as
                Office JS's ``ItemNotFound`` (core-ts check 53).
        """
        items, owner = story_of(self.paragraph.element)
        theirs, _ = story_of(other.paragraph.element)
        if theirs is not items:
            raise SpanError(
                "expand_to: the two ranges are in different stories, which Office JS refuses too "
                "(ItemNotFound, core-ts check 53)",
                code="range.other_story",
                hint="expand to a range of the same body, note, comment or text box",
            )
        ends = (self.paragraph, self.end_paragraph, other.paragraph, other.end_paragraph)
        indices = [top_index(p.element, items, owner) for p in ends]
        if any(index < 0 for index in indices):
            raise ContentError(
                "expand_to: a range is outside the story as a reader reads it",
                code="range.stale",
                hint="take the ranges again",
            )
        layout = story_layout(
            items,
            lambda element, container: self._view_of(element, container).text,
            start=min(indices),
            stop=max(indices) + 1,
        )

        def at(paragraph: Paragraph, offset: int) -> int:
            index = layout.index_of(paragraph.element)
            if index < 0:
                raise ContentError(
                    "expand_to: a range is outside the story as a reader reads it",
                    code="range.stale",
                    hint="take the ranges again",
                )
            return layout.places[index].start + offset

        first = other if at(other.paragraph, other.start) < at(self.paragraph, self.start) else self
        mine = at(self.end_paragraph, self.end) + (1 if self.mark else 0)
        theirs_end = at(other.end_paragraph, other.end) + (1 if other.mark else 0)
        last = other if theirs_end > mine else self
        return Range(
            first.paragraph,
            first.start,
            last.end,
            end_paragraph=last.end_paragraph,
            mark=last.mark,
            end_mark=last.end_mark,
            whole=first.whole and last.whole,
        )

    def get_xml(self) -> str:
        """The span as a ``w:p``: a copy of the paragraph trimmed to the span.

        The paragraph is deep-copied (CR-001's ``deep_copy``, which never
        follows the parent pointer), the text outside the span is spliced out
        of the copy, and the copy is serialised. The document is not touched.

        Raises:
            SpanError: ``range.spans_paragraphs``, a range across paragraphs,
                whose blocks one ``w:p`` can't hold (CR-013 Phase B).
        """
        from docx4j_py.model.content.body import Body

        self.require_one_paragraph("get_xml")
        from docx4j_py.model.content.paragraph import Paragraph
        from docx4j_py.wml import Body as BodyElement

        copy = deep_copy(self.paragraph.element)
        holder = BodyElement(content=ChildList([copy]))
        copy.parent = holder
        body = Body(
            self.paragraph.parent_body.part,
            holder,
            prefix="fragment",
            package=self.paragraph.parent_body.package,
        )
        view = Paragraph(copy, holder.content, body)
        segments = view.segments()
        if any(segment.branch is not None or segment.math for segment in segments):
            _trim_around_branches(view, self.start, self.end)
            return to_xml(copy)
        length = len(view.text)
        if self.end < length:
            view.splice(self.end, length, "")
        if self.start > 0:
            view.splice(0, self.start, "")
        return to_xml(copy)

    def to_api_script(self, **options: Any) -> str:
        """The content-API calls that would produce this span's runs (CR-003 section 3.11).

        Python source against a name ``body``, executable as it stands; see
        :func:`docx4j_py.model.content.api_script.to_api_script` for the
        options (``variable``, ``location``, ``pictures``, ``limit``,
        ``addresses``).

        Raises:
            SpanError: ``range.spans_paragraphs``, a range across paragraphs
                (CR-013 Phase B).
        """
        from docx4j_py.model.content.api_script import to_api_script

        self.require_one_paragraph("to_api_script")
        return to_api_script(self, **options)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: the address, the text and the two offsets. Extension.

        A range across paragraphs adds ``end_address`` (its end paragraph's)
        and ``mark``, so a range within one paragraph reads as it did.
        """
        out: dict[str, Any] = {
            "address": self.paragraph.address,
            "text": self.text,
            "start": self.start,
            "end": self.end,
        }
        if self.spans_paragraphs:
            out["end_address"] = self.end_paragraph.address
            out["mark"] = self.mark
        return out

    def _refuse_span_edit(self, member: str) -> None:
        """A deletion or replacement across paragraphs: Word's own refusals first, then Phase C's (CR-013 section 4.7).

        core-ts's ``extentOf`` and ``assertTrackable`` (``spanEdits.mts`` at
        ``53db5bc``), read only: a range starting or ending inside a table it
        doesn't hold whole is refused as Word refuses it, and inside a content
        control as unmeasured; under change tracking, one holding a whole
        block that is neither a paragraph nor a table. What would be written
        is Phase C's, refused by :meth:`require_one_paragraph`.

        Raises:
            SpanError: ``range.partial_table``, ``range.partial_control`` or
                ``range.not_trackable``.
        """
        extent = _extent_of(self, member)
        if self.paragraph.change_tracker is not None:
            _assert_trackable(extent, member)


def _trim_around_branches(view: Any, start: int, end: int) -> None:
    """Trim a copy of a paragraph to ``[start, end)`` where its text holds a kept branch's (CR-011 Phase F).

    :meth:`Range.get_xml`'s copy: the paragraph's own text outside the span is
    spliced out, as elsewhere; a run- or paragraph-level ``mc:AlternateContent``
    whose text lies wholly outside the span goes, every branch with it; one
    whose text the span holds, in part or whole, is kept whole, since cutting
    it would change Word's branch alone. An equation is kept whole the same way,
    since an edit never cuts one (CR-013 Phase A).
    """
    from docx4j_py.model.content.text_model import branch_extents

    segments = view.segments()
    length = segments[-1].end if segments else 0
    extents = branch_extents(segments) + [
        (segment.start, segment.end, segment.item)
        for segment in segments
        if segment.math and segment.branch is None
    ]
    extents.sort(key=lambda extent: extent[0])
    cuts: list[tuple[int, int]] = []
    for a, b in ((0, start), (end, length)):
        if b <= a:
            continue
        pieces = [(a, b)]
        for low, high, _element in extents:
            next_pieces = []
            for x, y in pieces:
                if high <= x or low >= y:
                    next_pieces.append((x, y))
                    continue
                if x < low:
                    next_pieces.append((x, low))
                if high < y:
                    next_pieces.append((high, y))
            pieces = next_pieces
        cuts.extend(pieces)
    gone = [
        (low, element)
        for low, high, element in extents
        if high <= start or low >= end
    ]
    by_id = {id(segment.branch): segment for segment in segments if segment.branch is not None}
    by_id.update({id(segment.item): segment for segment in segments if segment.math and segment.branch is None})
    # last first, so that the offsets of what is still to cut hold
    for x, y in sorted(cuts, reverse=True):
        view.splice(x, y, "")
    for _low, element in sorted(gone, key=lambda pair: pair[0], reverse=True):
        _remove_element(by_id[id(element)], element)
    view._remove_empty_runs()


def _remove_element(segment: Any, element: Any) -> None:
    """Take a kept ``mc:AlternateContent`` out of the list holding it (a run's, or a run-level list)."""
    parent = getattr(element, "parent", None)
    for owner in (getattr(parent, "content", None), segment.outer_owner):
        if isinstance(owner, list):
            for index, item in enumerate(owner):
                if item is element:
                    del owner[index]
                    return


# ---------------------------------------------------------------------------
# ranges across paragraphs (CR-013 Phase B)
# ---------------------------------------------------------------------------


def _check_order(first: Paragraph, last: Paragraph) -> None:
    """Refuse a range from `first` to `last` unless both are in one story and `last` is not before `first`.

    Raises:
        SpanError: ``range.other_story`` or ``range.order``.
    """
    items, owner = story_of(first.element)
    other, _ = story_of(last.element)
    if other is not items:
        raise SpanError(
            "a range's two paragraphs must be in one story, and these are not (a note, a comment, "
            "a text box or a branch Word does not read is a story of its own)",
            code="range.other_story",
            hint="take both ends from the same body, note, comment or text box",
        )
    a = top_index(first.element, items, owner)
    b = top_index(last.element, items, owner)
    if a < 0 or b < 0:
        raise SpanError(
            "a range's paragraphs must be in their story as a reader reads it",
            code="range.other_story",
            hint="take both ends from the same body, note, comment or text box",
        )
    if b == a:
        layout = story_layout(items, lambda _element, _container: "", start=a, stop=a + 1)
        i, j = layout.index_of(first.element), layout.index_of(last.element)
        if i < 0 or j < 0:
            raise SpanError(
                "a range's paragraphs must be in their story as a reader reads it",
                code="range.other_story",
                hint="take both ends from the same body, note, comment or text box",
            )
        a, b = i, j
    if b < a:
        raise SpanError(
            "a range's end paragraph comes before its start paragraph",
            code="range.order",
            hint="give the earlier paragraph first, or use expand_to(), which takes either order",
        )


def _main_body_of(package: Any) -> Any:
    """A package's main body, or None where it has none (a detached body)."""
    if package is None:
        return None
    try:
        return package.body
    except Exception:  # noqa: BLE001 - a package without a main document part has no main body
        return None


def _owner_of(body: Any) -> Any:
    """The object whose list a body's content is."""
    try:
        return body._owner()
    except Exception:  # noqa: BLE001 - a body over nothing block-level holds nothing
        return None


def _holds(body: Any, element: Any) -> bool:
    """Whether `element` is under `body`'s container, through the parent pointers."""
    container = body.container
    owner = _owner_of(body)
    node = getattr(element, "parent", None)
    for _ in range(256):
        if node is None:
            return False
        if node is container or node is owner:
            return True
        node = getattr(node, "parent", None)
    return False


def _character_style_of(run: Any) -> str:
    """A run's own ``w:rStyle``, ``""`` where it has none."""
    style = getattr(getattr(run, "r_pr", None), "r_style", None)
    value = getattr(style, "val", None)
    return str(getattr(value, "value", value)) if value else ""


#: Not block-level, though a block container's list holds them: a table's rows, a row's cells and their wrappers.
_ROW_AND_CELL = frozenset({"CTSdtRow", "CTSdtCell", "CTCustomXmlRow", "CTCustomXmlCell", "Tr", "Tc"})


def _block_level(node: Any) -> bool:
    cls = node.__class__
    return not any(base.__name__ in _ROW_AND_CELL for base in cls.__mro__ if base.__module__ == "docx4j_py.wml")


def _blocks_holding(element: Any, story: list) -> list[tuple[list, Any]]:
    """The blocks holding a paragraph, innermost first: the paragraph in its list, each block above in its, up to the story's.

    core-ts's ``blocksHolding`` (``spanEdits.mts`` at ``53db5bc``).
    """
    out: list[tuple[list, Any]] = []
    node = element
    for _ in range(256):
        parent = getattr(node, "parent", None)
        if parent is None:
            break
        items = block_children_of(parent)
        if isinstance(items, list) and any(item is node for item in items) and _block_level(node):
            out.append((items, node))
            if items is story:
                break
        node = parent
    return out


def _extent_of(span: Range, member: str) -> tuple[list, int, int, bool, bool]:
    """Where a range's two ends stand in the innermost block list holding both: ``(list, first, last, head, tail)``.

    core-ts's ``extentOf`` (``53db5bc``), its refusals only: `head` is True
    where the start's paragraph is the first block itself, `tail` where the
    end's is the last; a block the range starts or ends inside without
    holding it whole is refused.

    Raises:
        SpanError: ``range.partial_table`` inside a table, Word's own refusal
            (core-ts check 52, cases 9, 10, 22, 23, 34, 35, 44 and 45);
            ``range.partial_control`` inside a content control or another
            block, unmeasured.
    """
    story, _owner = story_of(span.paragraph.element)
    up = _blocks_holding(span.paragraph.element, story)
    down = _blocks_holding(span.end_paragraph.element, story)
    if not up or not down or up[-1][0] is not story or down[-1][0] is not story:
        raise ContentError(
            "this range's paragraphs are not in the document any more: an edit through another view "
            "removed or joined them",
            code="range.stale",
            hint="take the range again",
        )
    for i, (items, a) in enumerate(up):
        j = next((k for k, (other, _b) in enumerate(down) if other is items), -1)
        if j < 0:
            continue
        b = down[j][1]
        first_index = next(k for k, item in enumerate(items) if item is a)
        last_index = next(k for k, item in enumerate(items) if item is b)
        if i > 0:
            first, _last = first_and_last_paragraph(items, first_index)
            if not (span.start == 0 and first is not None and first.element is span.paragraph.element):
                _refuse_partial(a, member, "starts")
        if j > 0:
            _first, last = first_and_last_paragraph(items, last_index)
            if not (span.holds_end_mark and last is not None and last.element is span.end_paragraph.element):
                _refuse_partial(b, member, "ends")
        return items, first_index, last_index, i == 0, j == 0
    raise SpanError(
        "this range's two ends are not in one story",
        code="range.other_story",
        hint="take both ends from the same body, note, comment or text box",
    )


def _refuse_partial(block: Any, member: str, side: str) -> None:
    if element_name(block) == W_TBL:
        raise SpanError(
            f"Range.{member}: this range {side} inside a table, which Word refuses too (core-ts check 52)",
            code="range.partial_table",
            hint="delete the table, then the paragraphs; or edit inside one cell",
        )
    raise SpanError(
        f"Range.{member}: this range {side} inside a content control (or another block) it doesn't hold "
        "whole, and what Word does with that isn't measured",
        code="range.partial_control",
        hint="take the whole content control, or edit inside it",
    )


#: The block-level markers a tracked deletion may hold whole (core-ts's ``isRangeMarker``).
def _is_range_marker(block: Any) -> bool:
    name = element_name(block) or ""
    local = name.rsplit("}", 1)[-1]
    return local.endswith(("RangeStart", "RangeEnd")) or local in (
        "bookmarkStart",
        "bookmarkEnd",
        "permStart",
        "permEnd",
        "proofErr",
    )


def _assert_trackable(extent: tuple[list, int, int, bool, bool], member: str) -> None:
    """Refuse a tracked deletion holding a whole block that is neither a paragraph nor a table (core-ts's ``assertTrackable``).

    Raises:
        SpanError: ``range.not_trackable``.
    """
    items, first, last, head, tail = extent
    for index in range(first, last + 1):
        if (index == first and head) or (index == last and tail):
            continue
        block = items[index]
        name = element_name(block)
        if name in (W_P, W_TBL) or _is_range_marker(block):
            continue
        local = (name or type(block).__name__).rsplit("}", 1)[-1]
        raise SpanError(
            f"Range.{member}: this range holds a whole {local}, and there is no measured way to delete one "
            "under change tracking (core-ts CR-014 phase 2)",
            code="range.not_trackable",
            hint="turn change tracking off, or write the paragraphs either side of it",
        )
