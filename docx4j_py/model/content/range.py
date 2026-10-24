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

**Across paragraphs** (CR-013 Phases B and C, the port of docx4j-core-ts's
CR-014 phases 1 to 3): a range reads in ``Body.text``'s form between its two
ends, and writes as Word 365 writes through Office JS (core-ts checks 52 and
54). A deletion, a replacement, a ``font`` write, a paragraph split and a
style go through :mod:`.span_edits`; a comment's markers go in the first and
last paragraph; a content control wraps paragraphs held whole; ``^p`` in
``search`` matches a paragraph's own mark. A hyperlink across paragraphs is
refused (``range.spans_paragraphs``), where Word writes a ``HYPERLINK`` field.
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
from docx4j_py.model.content.reports import composed_moves, recording
from docx4j_py.model.content.story import (
    StoryLayout,
    story_layout,
    story_of,
    top_index,
)
from docx4j_py.model.content.styles import built_in_of, style_name_of
from docx4j_py.model.content.text_model import (
    PARAGRAPH_BREAKS,
    RUN_HOLDER_NAMES,
    MarkedSegments,
    canonical_view,
    drop_unwritable,
    office_matches,
    paragraph_mark_pieces,
    refuse_mark_replacement,
    refuse_wildcard_mark_replacement,
    require_view,
    wildcard_matches_across,
)
from docx4j_py.model.content.tracking import REVISION_NAMES
from docx4j_py.traversal import element_name
from docx4j_py.wml import Tc, to_xml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = ["Range"]

#: How much of the text a ``repr`` shows.
_PREVIEW = 30


class Range:
    """A subset of Office JS ``Word.Range``: ``[start, end)`` from one paragraph to the same or a later one of its story."""

    __slots__ = ("_end_paragraph", "end", "end_mark", "mark", "of_control", "paragraph", "start", "whole")

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
        section 10.11 item 2), and since Phase C a `whole` range takes them
        too, so that an empty paragraph's ``'Content'`` range takes the
        comment reference it holds (core-ts's ``WHOLE``).

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
        #: ``(control, "Content" | "Whole")`` for a content control's own range, else None (CR-019
        #: section 4.4; core-ts's internal ``ofControl``): a run-level control showing its placeholder,
        #: or holding nothing, has no width, so its own range is where its writes go in, not its
        #: offset, which the text either side of it shares. Set by :meth:`ContentControl.get_range`.
        self.of_control: tuple[Any, str] | None = None
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
        """Refuse, before anything changes, a write a range across paragraphs doesn't take. Extension.

        CR-013 Phase C (section 13): what stays refused across paragraphs ---
        a ``hyperlink``, where Word writes a ``HYPERLINK`` field from the first
        paragraph into the last (core-ts check 54 case 34), and the members
        that sit in one run holder (:meth:`holder_boundaries`). A range within
        one paragraph that holds a whole table (a one-cell table's ``'Whole'``
        range, its :attr:`end_mark` the row end) is refused too, since its
        write would take the table.

        Raises:
            SpanError: ``range.spans_paragraphs``.
        """
        if self.spans_paragraphs or (self.end_mark is not None and self.holds_end_mark):
            what = "spans paragraphs" if self.spans_paragraphs else "holds a whole table"
            raise SpanError(
                f"{operation} refused: this range {what}, which this write doesn't take "
                "(CR-013 Phase C: a hyperlink is one paragraph's)",
                code="range.spans_paragraphs",
                hint="write each paragraph's range (range.paragraphs[i].get_range()), or the "
                "content control, comment or table the range came from",
            )

    def _cover(self, other: Range, *, fresh: bool = False) -> None:
        """This range moved onto another's span; `fresh`, it no longer holds a mark or a paragraph's whole content."""
        self.paragraph = other.paragraph
        self.start = other.start
        self.end = other.end
        self.end_paragraph = other.end_paragraph
        if fresh:
            self.whole = False
            self.mark = False
            self.end_mark = None

    def _land_at(self, paragraph: Paragraph, start: int, end: int | None = None) -> None:
        """This range collapsed (or over `start`..`end`) in one paragraph, where a write closed it up."""
        self.paragraph = paragraph
        self._end_paragraph = None
        self.start = start
        self.end = start if end is None else end
        self.whole = False
        self.mark = False
        self.end_mark = None

    def _refuse_read_only(self, operation: str, *, ends_only: bool = False) -> None:
        """Refuse a write through a paragraph of the range that is read-only: one in a kept branch (CR-012 Phase A).

        core-ts's ``assertWritableSpan``. A deletion and the writes at the
        range's ends ask of its two ends only: a kept block-level element
        between them goes whole untracked, as ``Body.clear()`` takes one
        (CR-013 section 4.5).
        """
        from docx4j_py.model.content.reports import refuse_read_only

        targets = [self.paragraph, self.end_paragraph] if ends_only else self.paragraphs
        for paragraph in targets:
            refuse_read_only(paragraph.parent_body, operation)

    def _refuse_edge_pieces(self, operation: str) -> None:
        """Refuse, before anything changes, a deletion whose end pieces hold what can't be cut.

        The first paragraph's piece and the last's obey the edit rule over a
        kept branch's text (CR-011 Phase F) and the equation rule (CR-013
        Phase A): an equation cut, or taken under tracking, is refused.
        """
        pieces = self._pieces()
        for piece in (pieces[0], pieces[-1]) if len(pieces) > 1 else pieces:
            piece.paragraph.refuse_branch_text(piece.start, piece.end, operation)
            # a deletion takes an equation whole, under change tracking in Word's form (CR-013 section 13.7, D2)
            piece.paragraph.refuse_equation(piece.start, piece.end, deleting=True)

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

        Setting it sets the paragraph's style; across paragraphs each
        paragraph's (core-ts check 54 cases 32 and 33; CR-013 Phase C), each
        with its ``w:pPrChange`` under tracking. Word also took away the body
        paragraph's centring and italic and the list item's numbering, which
        :attr:`Paragraph.style` doesn't (core-ts CR-014 section 12.1 item 10).
        """
        char, paragraphs = self._style_read()
        if char is None:
            return paragraphs(lambda p: p.style)
        if char == "":
            return ""
        return style_name_of(self.paragraph.parent_body.package, char)

    @style.setter
    def style(self, value: str) -> None:
        self._set_style("style", lambda p: setattr(p, "style", value))

    @property
    def style_id(self) -> str:
        """The paragraph's style id; across paragraphs the first one's. Extension, as on :class:`Paragraph`."""
        return self.paragraph.style_id

    @style_id.setter
    def style_id(self, value: str) -> None:
        self._set_style("style_id", lambda p: setattr(p, "style_id", value))

    @property
    def style_built_in(self) -> str:
        """The built-in name of :attr:`style` (Office JS ``styleBuiltIn``): ``"Other"`` where it is ``""``, ``"Normal"`` where paragraphs differ.

        core-ts's subset declares it read-only; the setter is kept here, and
        sets each paragraph's style as :attr:`style` does (CR-013 section 4.3;
        across paragraphs since Phase C, as core-ts's ``styleBuiltIn`` setter).
        """
        char, paragraphs = self._style_read()
        if char is None:
            return paragraphs(lambda p: p.style_built_in)
        if char == "":
            return "Other"
        return built_in_of(char)

    @style_built_in.setter
    def style_built_in(self, value: str) -> None:
        self._set_style("style_built_in", lambda p: setattr(p, "style_built_in", value))

    def _set_style(self, name: str, write: Callable[[Paragraph], None]) -> None:
        """A style written to the range's paragraph, or to each across paragraphs, as one report (core-ts's ``styledParagraphs``)."""
        if self.spans_paragraphs:
            self._refuse_read_only(f"Range.{name}")
        self.refuse_branch_text(name)
        if not self.spans_paragraphs:
            write(self.paragraph)
            return
        paragraphs = self.paragraphs
        with recording(self.paragraph.parent_body, "format") as change:
            change.text(before=f"{name}={getattr(self, name)}")
            for paragraph in paragraphs:
                write(paragraph)
                change.touched(paragraph)
            change.text(after=f"{name}={getattr(self, name)}")

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

        A **read** splits nothing (CR-016 Q12) and reports **every run the
        span touches** (a run partly inside counts, :attr:`runs`), each
        **effective** as Office JS reports it (CR-002 section 15.1): a property
        where they differ reads ``None``, ``""`` for the name, the colour and
        the highlight, and ``"Mixed"`` for the underline (CR-016 Phase B;
        core-ts at ``222cc20``, its CR-002 section 59). A range holding its
        paragraph's mark and no run --- an empty paragraph's ``"Whole"`` range
        --- reads the mark's run properties, and a collapsed range its
        paragraph's style. An automatic colour reads ``#000000``. A **write**
        splits the runs at the boundaries first --- the start backwards and
        the end forwards to the nearest grapheme boundary, so a cluster is
        never cut in two (CR-003 section 3.12) --- so that only the span is
        touched.

        **Across paragraphs** (CR-013 Phase B, Q4) it is a
        :class:`~.font.MixedFont`: each run read in its own paragraph, by the
        same rule, as Office JS reads it (core-ts checks 52 and 53). A write
        (Phase C) goes to every run in the range, a table's cells' too, and to
        each mark it holds, a property written only where a run or mark doesn't
        read it already, as Word formats them (core-ts checks 52 and 54;
        :func:`~.span_edits.format_span`); tracked, a mark's change is its
        ``w:pPr/w:rPr/w:rPrChange``.
        """
        return self.get_font()

    def get_font(self, *, direct: bool = False) -> Any:
        """:attr:`font`, or the **direct** formatting when `direct` is given.

        CR-002 section 15.1, as :meth:`Paragraph.get_font` is.
        """
        from docx4j_py.model.content.font import mark_run_properties

        if self.spans_paragraphs:

            def readings() -> list[dict[str, Any]]:
                out: list[dict[str, Any]] = []
                for piece in self._pieces():
                    paragraph = piece.paragraph
                    reader = Font(list, context=paragraph.resolution_context, direct=direct)
                    out.extend(
                        reader.read_each([getattr(run, "r_pr", None) for run in piece.runs], direct=direct)
                    )
                return out

            return MixedFont(readings, self._format_across)

        def holders() -> list[Any]:
            if self.start == self.end:
                return []
            # a write formats runs holding the span and nothing else: a tab, a break
            # or text after it in the same run is split off too (core-ts's
            # splitRunsAround, its CR-012 phase 4; CR-013 Phase C)
            self.start, self.end = self.paragraph.split_runs_around(self.start, self.end)
            return self.runs

        def reads() -> list[Any]:
            # a read splits nothing: the runs the span touches (CR-016 Q12, core-ts's holders(false))
            return [] if self.start == self.end else self.runs

        def mark() -> Any:
            # a range holding its paragraph's mark and no run reads the mark (core-ts CR-002 section 59)
            return mark_run_properties(self.paragraph.element) if self.holds_end_mark else None

        return Font(
            holders,
            scope=f"{self.start}:{self.end}",
            record=self.paragraph.formatting,
            tracking=self.paragraph.font_tracking,
            context=self.paragraph.resolution_context,
            direct=direct,
            guard=lambda: self.refuse_branch_text("a font write"),
            reads=reads,
            mark=mark,
        )

    def _format_across(self, options: dict[str, Any]) -> None:
        """A ``font`` write over a range across paragraphs (:func:`~.span_edits.format_span`), one ``"format"`` report."""
        from docx4j_py.model.content.span_edits import FontPiece, format_span

        self._refuse_read_only("a font write")
        self.refuse_branch_text("a font write")
        name, value = next(iter(options.items()))
        pieces = self._pieces()
        with recording(self.paragraph.parent_body, "format") as change:
            change.text(before=f"{name}={self.get_font(direct=True)._read().get(name)}")
            format_span(
                [
                    FontPiece(
                        piece.paragraph,
                        piece.start,
                        piece.end,
                        lambda piece=piece: piece.runs,
                        index < len(pieces) - 1 or self.holds_end_mark,
                    )
                    for index, piece in enumerate(pieces)
                ],
                options,
            )
            change.touched(*(piece.paragraph for piece in pieces))
            change.text(after=f"{name}={value}")

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

        **Across paragraphs** (CR-013 Phase C; docx4j-core-ts check 54 cases 36
        to 43): a range holding its paragraphs whole, from the first's start to
        the last's end, its mark or not, is wrapped in one block-level control,
        as Word wraps it; tracked, the control is an insertion in Word's form
        (``w:customXmlInsRangeStart`` and ``End`` around it, check 30). Any
        other range is refused, as Word refuses it.

        Raises:
            SpanError: the span crosses a run holder; the message names where
                to split it. ``range.partial_paragraph``, a range across
                paragraphs starting or ending inside a paragraph, across a
                table, into one or out of one (Word's ``GeneralException``).
            BindingError: ``"RepeatingSection"`` was asked for, which Word has
                no run-level form of.
        """
        from docx4j_py.model.customxml.insert import (
            insert_content_control_in_range,
            wrap_paragraphs,
        )

        if self.spans_paragraphs:
            return wrap_paragraphs(self, kind)
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

        **A replacement** across paragraphs, or holding a whole table, or under
        change tracking within one paragraph too, is Word's (CR-013 Phase C;
        core-ts checks 52 and 54; :mod:`.span_edits`): untracked, the range
        is deleted keeping a last mark it holds, the paragraphs joining in the
        last one's properties, and the text goes where they joined in the first
        replaced character's look; tracked, the ``w:ins`` goes **first**,
        before the deletion, in the first replaced run's formatting (check 54
        case 4, the user's decision in core-ts: "follow Word (insertion
        first)"). **Across paragraphs** ``"Start"`` and ``"Before"`` put the
        text at the range's first character, ``"End"`` and ``"After"`` after
        its last, before a last mark it holds.

        **A content control's own range** (CR-019 section 4.4, core-ts's D4)
        asks its control first: into a control showing its placeholder the
        text goes in the placeholder's place (:meth:`ContentControl.insert_text`'s
        write), into an empty run-level control inside it, and at ``"Start"``
        or ``"Before"`` of a run-level control holding text before its text,
        inside it (CR-019 Q4).

        Returns:
            The range of the inserted text; where it made paragraphs, a range
            across them, as Office JS's is (core-ts check 55: ``x\\ry``; CR-013
            Phase B). This range goes on covering the same text: a replacement
            makes it the new text's range, text before it shifts it, and text
            after it leaves it where it was, in the paragraph that now holds it.

        Raises:
            SpanError: on a replacement, ``range.partial_table`` for one
                starting or ending inside a table, as Word refuses it (core-ts
                check 52); ``range.partial_control`` inside a content control;
                ``range.replaces_table`` for one starting or ending with a
                whole table or control (Word's refusal, check 54 case 7);
                ``range.not_trackable`` under change tracking for one holding a
                whole ``w:altChunk`` or kept ``mc:AlternateContent``.
        """
        if location not in ("Replace", "Before", "Start", "After", "End"):
            raise ContentError(
                f"insert_text on a range takes Replace, Before, After, Start or End, "
                f"not {location!r}",
                code="location.invalid",
                hint="Replace is the default; Before and After put text around the span",
            )
        if self.of_control is not None:
            # a control's own range: into a control showing its placeholder, or an empty run-level one,
            # its writes go inside it (CR-019 section 4.4, core-ts's D4)
            control, part = self.of_control
            written = control._write_through(part, text, location)
            if written is not None:
                return written
        if location == "Replace" and (
            self.spans_paragraphs or self.end_mark is not None or self.paragraph.change_tracker is not None
        ):
            return self._replace(text)
        text = drop_unwritable(text)
        if self.spans_paragraphs:
            return self._insert_at_ends(text, location)
        with recording(self.paragraph.parent_body, "insert_text") as change:
            change.touched(self.paragraph)
            change.text(before=self.text)
            paragraph = self.paragraph
            container = paragraph.container
            index = next((i for i, item in enumerate(container) if item is paragraph.element), -1)
            length = len(paragraph.text)
            if location == "Replace":
                # a paragraph's 'Whole' or 'Content' range holds its hidden runs, pictures and comment
                # references at its edges, an empty one's too (core-ts's WHOLE; CR-013 Phase C)
                out = paragraph.splice(self.start, self.end, text, whole=self.mark or self.whole)
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

    def _insert_at_ends(self, text: str, location: str) -> Range:
        """``insert_text`` at a range across paragraphs' start (in its first paragraph) or end (in its last, before a last mark)."""
        from docx4j_py.model.content.paragraph import Paragraph

        self._refuse_read_only("insert_text", ends_only=True)
        with recording(self.paragraph.parent_body, "insert_text") as change:
            change.text(before=self.text)
            if location in ("End", "After"):
                paragraph = self.end_paragraph
                container = paragraph.container
                index = next((i for i, item in enumerate(container) if item is paragraph.element), -1)
                at = self.end
                out = paragraph.splice(at, at, text, edge="end" if at == len(paragraph.text) else None)
                if 0 <= index < len(container) and container[index] is not paragraph.element:
                    # the range's end is before the text: in the first paragraph the split made
                    self.end_paragraph = Paragraph(container[index], container, paragraph.parent_body)
                change.touched(paragraph)
            else:
                out = self.paragraph.splice(self.start, self.start, text, edge="start" if self.start == 0 else None)
                # the text may have broken the paragraph (E1): the range follows its text
                self.paragraph = out.end_paragraph
                self.start = out.end
                change.touched(self.paragraph)
            change.text(after=text)
            return out

    def _replaceable_extent(self, member: str) -> Any:
        """The extent of a range a replacement may take (core-ts's ``replaceableExtent``).

        Raises:
            SpanError: ``range.replaces_table``, a range ending with a whole
                table or control, which Word refuses to replace (check 54 case
                7), or starting with one (unmeasured, with no paragraph for the
                text to go in); and :func:`~.span_edits.extent_of`'s.
        """
        from docx4j_py.model.content.span_edits import extent_of

        # an end in a kept branch is read-only before anything else is asked (CR-013 section 4.5)
        self._refuse_read_only(member, ends_only=True)
        extent = extent_of(self, member)
        if extent.tail is None or extent.head is None:
            side = "ends" if extent.tail is None else "starts"
            raise SpanError(
                f"Range.{member}: this range {side} with a whole table or content control, which Word "
                "refuses to replace (core-ts check 54 case 7)",
                code="range.replaces_table",
                hint="delete the range, then insert the text where it closed up; or replace a range "
                "that ends in a paragraph",
            )
        return extent

    def _replace(self, text: str) -> Range:
        """``insert_text(text, "Replace")`` across paragraphs, holding a whole table, and any under tracking (core-ts's ``replaceAcross``)."""
        from docx4j_py.model.content.span_edits import (
            assert_trackable,
            delete_span,
            extent_of,
            first_character_of,
            place_replacement,
            refuse_tracked_content,
        )
        from docx4j_py.model.content.split import refuse_split

        extent = self._replaceable_extent("insert_text")
        text = drop_unwritable(text)
        tracker = self.paragraph.change_tracker
        body = self.paragraph.parent_body
        self._refuse_edge_pieces("insert_text")
        if tracker is not None:
            assert_trackable(extent, "insert_text")
            refuse_tracked_content(extent, "insert_text", body)
            # the insertion goes first, before the range's first character: its breaks checked there
            self.paragraph.refuse_paragraph_breaks(
                self.start,
                self.start,
                text,
                "insert_text",
                prefer="after",
                edge="start" if self.start == 0 and self.whole else None,
            )
        elif PARAGRAPH_BREAKS.search(text):
            # a line break in the text is a paragraph break (E1), refused before anything changes where none can be made
            refuse_split(self.paragraph, self.start, "head", "insert_text")
        with recording(body, "insert_text") as change, composed_moves(body, extent.items):
            change.text(before=self.text)
            change.touched(*self.paragraphs)
            if tracker is not None:
                # tracked: the insertion before the deletion (check 52 cases 39 to 48, check 54 case 4),
                # in the first replaced run's formatting
                length = None if self.spans_paragraphs else self.end - self.start
                # a range holding its paragraph's whole content puts the text before the hidden
                # runs at its start, which the deletion then takes (core-ts's WHOLE)
                edge = "start" if self.start == 0 and self.whole else None
                inserted = self.paragraph.splice(self.start, self.start, text, prefer="after", edge=edge)
                # the range's own text follows what was put in, in the paragraph the insertion ended in (E1)
                self.paragraph = inserted.end_paragraph
                self.start = inserted.end
                if length is not None:
                    self.end_paragraph = inserted.end_paragraph
                    self.end = inserted.end + length
                delete_span(self, extent_of(self, "insert_text"), True)
                self._cover(inserted, fresh=True)
                change.text(after=text)
                return Range(inserted.paragraph, inserted.start, inserted.end, end_paragraph=inserted.end_paragraph)
            first = first_character_of(self)
            landing = delete_span(self, extent, True)
            assert landing is not None  # a replaceable extent has a head and a tail
            first_paragraph, last_paragraph, start, end = place_replacement(landing, first, text)
            out = Range(first_paragraph, start, end, end_paragraph=last_paragraph)
            self._cover(out, fresh=True)
            change.touched(first_paragraph, last_paragraph)
            change.text(after=text)
            return out

    def insert_paragraph(
        self,
        text: str = "",
        *,
        location: ParagraphLocation = "After",
        style: str | None = None,
    ) -> Paragraph:
        """A new paragraph where the range starts (``"Before"``) or ends (``"After"``), splitting the paragraph there.

        As Word does (CR-013 Phase C; docx4j-core-ts's ``insertParagraphAt``):
        across paragraphs (check 52 cases 57 to 70) and within one (check 54
        cases 1 to 3, where 0.1.0 put the new paragraph before or after the
        whole paragraph). Inside a paragraph, the text before the split becomes
        a paragraph of its own in front, and the new paragraph goes between the
        two, its mark and its text in the formatting of the run before the
        split; a link, a revision or a smart tag there is cut in two. At a
        paragraph's start or end, or after a range holding its mark, it goes
        before or after the paragraph, in its properties. Tracked, the new
        paragraph and the first half's mark are inserted. This range keeps
        covering its text.

        Raises:
            SpanError: ``range.splits_control``, ``range.splits_field``,
                ``range.cuts_equation``: the split falls inside a run-level
                content control, a field or an equation; nothing has changed.
            AddressError: ``address.read_only``: inside a kept branch's text.
        """
        from docx4j_py.model.content.span_edits import insert_paragraph_at

        if location not in ("Before", "After"):
            raise ContentError(
                f"insert_paragraph on a range takes Before or After, not {location!r}",
                code="location.invalid",
                hint="Before splits the paragraph where the range starts, After where it ends",
            )
        self._refuse_read_only("insert_paragraph", ends_only=True)
        spans = self.spans_paragraphs
        with recording(self.paragraph.parent_body, "insert_paragraph") as change:
            if location == "Before":
                inserted, first_half = insert_paragraph_at(self.paragraph, self.start, text)
                if first_half is not None:
                    # the paragraph keeps the text from the split on
                    if not spans:
                        self.end -= self.start
                    self.start = 0
            else:
                end_paragraph = self.end_paragraph
                at = len(end_paragraph.text) if self.holds_end_mark else self.end
                inserted, first_half = insert_paragraph_at(end_paragraph, at, text)
                if first_half is not None:
                    # the text before the split moved into the first half
                    if not spans:
                        self.paragraph = first_half
                    self.end_paragraph = first_half
            if style is not None:
                inserted.style = style
            change.touched(inserted)
            change.text(after=drop_unwritable(text))
            return inserted

    def insert_ooxml(self, ooxml: str, *, location: str = "Replace") -> list[Any]:
        """Word's ``insertOoxml`` around or in place of this span (CR-003 section 3.2).

        A flat OPC ``pkg:package`` or a bare fragment, as
        :meth:`Body.insert_ooxml`. A fragment of exactly one ``w:p`` has its
        **runs merged into this paragraph** at the span's start or end, as
        Word's paste does; anything with a block in it goes before or after the
        paragraph.

        **Across paragraphs** (CR-013 Phase C; core-ts check 54 case 35): one
        paragraph's runs replace the range as ``insert_text(..., "Replace")``
        replaces it, where the paragraphs joined (``HEAD35A NEWOOXML BODY35B``);
        ``"Before"`` puts them at the range's start in its first paragraph,
        ``"After"`` at its end in its last; several blocks go before the first
        paragraph or after the last. Within one paragraph a ``"Replace"`` keeps
        a last mark the range holds, as a replacement does.

        Args:
            ooxml: the ``pkg:package`` document, or the fragment.
            location: ``"Replace"`` (the default, as a range's other verbs),
                ``"Before"`` or ``"After"``.

        Returns:
            The views of what was inserted --- the paragraph the runs went into
            when they were merged in.

        Raises:
            SpanError: ``range.spans_paragraphs``, several blocks in place of a
                range across paragraphs, which Word wasn't measured doing.
        """
        from docx4j_py.model.content.ooxml import insert_ooxml_into_range

        return insert_ooxml_into_range(self, ooxml, location=location)

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` within this span, as ranges of the paragraph.

        Matches are found in Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, which reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, section 15).

        Across paragraphs (CR-013 Phase B) a paragraph at a time, from
        :attr:`start` in the first to :attr:`end` in the last, as core-ts's.
        **``^p``**, Word's paragraph mark (CR-013 Phase C; core-ts check 52
        case 2), in a plain query matches across paragraphs: a hit runs from
        its first piece's match ending at a paragraph's end, through each
        middle piece matching a whole paragraph's text, to its last piece's
        match from the next paragraph's start; a query ending with ``^p`` ends
        with that paragraph's mark, which the hit holds (``^p`` alone is a
        range of ``\\r``). A mark matches only a paragraph's own mark the
        range holds: never a cell's end or a row's end, and never across a
        table's edge; ``^p^p`` is an empty paragraph's mark after another
        mark, and ``^^`` stays a caret. Wildcards read ``^p`` as the two
        characters, as they did.

        **Wildcards across marks** (CR-016 Phase C; core-ts ``3ba1e68``, its
        CR-002 section 32's K7, held to Word 365 by check 59): a range across
        paragraphs, or one holding its mark, is searched across its
        paragraphs' marks, where ``^p`` matches one: ``?``, ``*`` and a class
        cross a paragraph's own mark the range holds and ``^13`` is one, never
        a cell's or a row's end, a deleted mark, or anything across a table's
        edge (:meth:`_search_wildcards`). A hit ending with a mark holds it
        (``\\r`` alone is a range of the mark).
        """
        limit = options.pop("limit", None)
        pieces = None if options.get("match_wildcards") else paragraph_mark_pieces(text)
        if pieces is not None:
            out = self._search_marks(pieces, options)
            return out if limit is None else out[:limit]
        if options.get("match_wildcards") and (self.spans_paragraphs or self.holds_end_mark):
            out = self._search_wildcards(text, options)
            return out if limit is None else out[:limit]
        if self.spans_paragraphs:
            out = [hit for piece in self._pieces() for hit in piece.search(text, **options)]
            return out if limit is None else out[:limit]
        hits = office_matches(
            self.paragraph.segments(objects=True), text, start=self.start, end=self.end, **options
        )
        if limit is not None:
            hits = hits[:limit]
        return [Range(self.paragraph, start, end) for start, end in hits]

    def _search_marks(self, pieces: list[str], options: dict[str, Any]) -> list[Range]:
        """A query's hits where it holds ``^p``, as Word finds them (core-ts's ``searchMarks``, check 52 case 2)."""
        paragraphs = self.paragraphs
        last = len(paragraphs) - 1
        marks = len(pieces) - 1
        ends_with_mark = pieces[marks] == ""

        def cell(paragraph: Paragraph) -> int | None:
            found = paragraph.parent_table_cell
            return None if found is None else id(found.element)

        def mark_in(j: int) -> bool:
            # a paragraph's mark the search may match: the range holds it, and it is a paragraph's, not a cell's end
            return (j < last or self.holds_end_mark) and paragraphs[j].mark_text == "\r"

        def matches(paragraph: Paragraph, query: str, start: int, end: int) -> list[tuple[int, int]]:
            return office_matches(paragraph.segments(objects=True), query, start=start, end=end, **options)

        out: list[Range] = []
        for i in range(last + 1):
            end_index = i + marks - (1 if ends_with_mark else 0)  # the paragraph the hit ends in
            if end_index > last:
                break
            scope = cell(paragraphs[i])
            ok = all(mark_in(j) for j in range(i, i + marks))
            ok = ok and all(cell(paragraphs[j]) == scope for j in range(i + 1, end_index + 1))
            # the middle pieces: each a whole paragraph's text (an empty one for an empty piece)
            for j in range(1, marks):
                if not ok:
                    break
                middle = paragraphs[i + j]
                length = len(middle.text)
                if pieces[j] == "":
                    ok = middle.text == ""
                else:
                    ok = any(a == 0 and b == length for a, b in matches(middle, pieces[j], 0, length))
            if not ok:
                continue
            first = paragraphs[i]
            low = self.start if i == 0 else 0
            length = len(first.text)
            if pieces[0] == "":
                starts = [length]
            else:
                starts = [a for a, b in matches(first, pieces[0], low, length) if b == length]
            if not starts or starts[0] < low:
                continue
            end_paragraph = paragraphs[end_index]
            if ends_with_mark:
                to = len(end_paragraph.text)
            else:
                limit = self.end if end_index == last else len(end_paragraph.text)
                hit = next(((a, b) for a, b in matches(end_paragraph, pieces[marks], 0, limit) if a == 0), None)
                if hit is None:
                    continue
                to = hit[1]
            out.append(
                Range(
                    first,
                    starts[0],
                    to,
                    end_paragraph=None if end_paragraph.element is first.element else end_paragraph,
                    mark=ends_with_mark,
                )
            )
        return out

    def _search_wildcards(self, query: str, options: dict[str, Any]) -> list[Range]:
        """A wildcard query's hits across the paragraphs' marks, as Word's wildcards match them (core-ts's ``searchWildcards``).

        CR-016 Phase C (core-ts ``Range.mts`` at ``3ba1e68``; check 59): where
        ``^p`` matches one (:meth:`_search_marks`), a paragraph's own mark the
        range holds (:attr:`Paragraph.mark_text` ``\\r``: never a cell's or a
        row's end, nor a deleted mark, which Word's default view shows as
        nothing), and never across a table's edge, so the paragraphs a hit
        spans are in one cell or in none: the range's paragraphs are taken in
        runs of consecutive ones in one cell (the ``w:tc`` nearest above each,
        the one :attr:`Paragraph.parent_table_cell` views) or in none, and each
        run's text is joined by the marks it may hold
        (:func:`~.text_model.wildcard_matches_across`). The last paragraph's
        mark counts only where the range holds it.
        """
        read = self._paragraphs_read()
        paragraphs = [paragraph for paragraph, _segments in read]
        last = len(paragraphs) - 1
        lengths = [sum(len(segment.text) for segment in segments) for _paragraph, segments in read]
        holds_end = self.mark and self.end >= lengths[last]

        def cell(paragraph: Paragraph) -> int | None:
            node = getattr(paragraph.element, "parent", None)
            for _ in range(256):
                if node is None or isinstance(node, Tc):
                    break
                node = getattr(node, "parent", None)
            return None if node is None else id(node)

        # runs of consecutive paragraphs in one cell, or in none
        groups: list[list[int]] = []
        previous: Any = object()
        for j, paragraph in enumerate(paragraphs):
            scope = cell(paragraph)
            if groups and scope == previous:
                groups[-1].append(j)
            else:
                groups.append([j])
            previous = scope
        out: list[Range] = []
        for group in groups:
            pieces = []
            for j in group:
                paragraph, segments = read[j]
                pieces.append(
                    MarkedSegments(
                        segments=segments,
                        start=self.start if j == 0 else 0,
                        end=self.end if j == last else lengths[j],
                        length=lengths[j],
                        mark=(j < last or holds_end) and paragraph.mark_text == "\r",
                    )
                )
            for hit in wildcard_matches_across(pieces, query, **options):
                first = paragraphs[group[hit.first]]
                end = paragraphs[group[hit.last]]
                out.append(
                    Range(
                        first,
                        hit.start,
                        hit.end,
                        end_paragraph=None if end.element is first.element else end,
                        mark=hit.mark,
                    )
                )
        return out

    def _paragraphs_read(self) -> list[tuple[Paragraph, list[Any]]]:
        """:attr:`paragraphs`, each with its segments (objects and hidden items listed), read once from the story's layout."""
        if not self.spans_paragraphs:
            return [(self.paragraph, self.paragraph.segments(objects=True))]
        items, _owner, low, high = self._slice()
        read: dict[int, tuple[Paragraph, list[Any]]] = {}

        def text_of(element: Any, container: list) -> str:
            view = self._view_of(element, container)
            segments = view.segments(objects=True)
            read[id(element)] = (view, segments)
            # Paragraph.text: an object's or a hidden item's segment holds no text
            return "".join(segment.text for segment in segments)

        layout = story_layout(items, text_of, start=low, stop=high + 1)
        first = layout.index_of(self.paragraph.element)
        last = layout.index_of(self.end_paragraph.element)
        if first < 0 or last < 0:
            raise ContentError(
                "this range's paragraphs are not in one story as a reader reads it",
                code="range.stale",
                hint="take the range again",
            )
        return [read[id(place.element)] for place in layout.places[first : last + 1]]

    def _matches(self, text: str, options: dict[str, Any]) -> list[Range]:
        """The hits a replacement takes: each paragraph's own, a paragraph at a time, none holding a mark.

        core-ts's ``officeMatches`` over each piece (``Range.replaceText`` at
        ``3ba1e68``): what :meth:`search` finds but for ``^p``'s and the
        wildcards' hits across marks, which a replacement refuses first.
        """
        limit = None
        if "limit" in options:
            options = dict(options)
            limit = options.pop("limit")
        if self.spans_paragraphs:
            out = [hit for piece in self._pieces() for hit in piece._matches(text, options)]
        else:
            hits = office_matches(
                self.paragraph.segments(objects=True), text, start=self.start, end=self.end, **options
            )
            out = [Range(self.paragraph, start, end) for start, end in hits]
        return out if limit is None else out[:limit]

    def replace_text(self, find: str, replace: str, **options: Any) -> int:
        """Replace every match within the span, last first; returns the count.

        **Tracked** when the package's ``change_tracking_mode`` is on, in which
        case each replacement is a ``w:ins`` and then a ``w:del``, as Word
        writes one (CR-013 Phase C). The report names the pair, as
        :meth:`Body.replace_text`'s does. Across paragraphs a paragraph at a
        time (CR-013 Phase C), every match checked before the first is
        replaced (CR-011 Phase F's rule; core-ts refuses paragraph by
        paragraph).

        **With wildcards** (CR-016 Phase C, Q9) every hit is found first, as
        :meth:`search` finds it; where any holds a paragraph's mark (Word's
        wildcards match one, core-ts check 59) the call is refused before
        anything changes. The rest replace a paragraph at a time.

        Raises:
            SpanError: ``range.replaces_mark``, a query holding ``^p``, Word's
                paragraph mark, or a wildcard query a hit of which holds a mark:
                the hits are found, but what Word's replacement of a mark leaves
                isn't measured (core-ts CR-014 section 11.3, CR-002 section 32).
        """
        from docx4j_py.model.content.paragraph import replace_matches

        refuse_mark_replacement("Range", find, options)
        if options.get("match_wildcards"):
            refuse_wildcard_mark_replacement("Range", self.search(find, **_unlimited(options)))
        replace = drop_unwritable(replace)
        if self.spans_paragraphs:
            self._refuse_read_only("replace_text")
        with recording(self.paragraph.parent_body, "replace_text") as change:
            change.text(before=find)
            matches = self._matches(find, options)
            for match in matches:  # every match checked before the first is replaced (CR-011 Phase F)
                match.refuse_branch_text("replace_text")
                match.paragraph.refuse_equation(match.start, match.end)
                match.paragraph.refuse_paragraph_breaks(match.start, match.end, replace, "replace_text")
            # a replacement holding a paragraph break splits the paragraph (CR-013 section 11, E1)
            replace_matches(matches, replace)
            if matches:
                change.touched(*{id(m.paragraph.element): m.paragraph for m in matches}.values())
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

        **Across paragraphs** (CR-013 Phase C; core-ts check 52 cases 79 to
        85): the start marker goes before the range's first character, in its
        first paragraph, and the end marker and the reference run after its
        last character, in its last paragraph, before a mark the range holds; a
        table between, or a cell at either end, stays as it is. The comment's
        :meth:`~.comments.Comment.get_range` is the range commented. Within one
        paragraph and across, the runs are split as ``font`` splits them, so a
        marker never goes beside a run holding a tab or a break outside the
        range (core-ts CR-014 section 11.2).

        Raises:
            SpanError: the span crosses a run holder --- a ``w:hyperlink``, a
                run-level ``w:sdt`` --- and the message says where to split
                (within one paragraph).
            ContentError: this body's package has no main document part.

        Returns:
            The new :class:`~docx4j_py.model.content.comments.Comment`.
        """
        from docx4j_py.model.content.comments import insert_comment_into

        return insert_comment_into(self, text)

    def delete(self) -> None:
        """Remove the span's text, leaving an empty range where it was.

        The pictures strictly inside the span go with it, and a paragraph's
        whole range (``'Whole'``, or a span of all its text) takes those at its
        edges too, as core-ts's CR-002 section 48 has it; under change tracking
        a picture's run goes into the ``w:del`` (CR-013 section 11).

        **Across paragraphs, or holding its paragraph's mark** (CR-013 Phase
        C; core-ts checks 52 and 54; :mod:`.span_edits`), as Word deletes it:
        the paragraphs join in the last one's properties, the blocks between go
        whole (a table, a content control), a last mark the range holds goes
        with it, joining the paragraph after (``paragraph.get_range().delete()``
        now deletes the paragraph, where 0.1.0 emptied it), and a note or a
        comment whose reference goes leaves when the write ends. A range into
        or out of a block content control joins nothing across the control's
        edge: the start's paragraph takes the end's properties. Tracked,
        nothing joins: the text is a ``w:del``, every mark the range holds is
        marked deleted, with no ``w:pPrChange``, a table's rows and cells'
        marks are deleted, and a whole control is deleted between Word's
        ``w:customXmlDelRange*`` markers. The range is left collapsed where the
        text closed up (tracked, where the deletion starts).

        **A content control showing its placeholder** (CR-019 sections 4.3
        and 4.4, core-ts's D3 and D4): the control's own ``"Content"`` range
        keeps it and its flag, its placeholder's runs going in the bare form
        (marked deleted under tracking); its ``"Whole"`` range takes a
        run-level one; a paragraph's ``"Content"`` range over a block
        control's only paragraph takes the control, its paragraph left in its
        place, emptied (tracked, the control marked deleted in Word's form).

        Raises:
            SpanError: ``range.partial_table`` for a range starting or ending
                inside a table, as Word refuses it (core-ts check 52);
                ``range.partial_control`` inside a content control, but one
                level down into or out of one; ``range.not_trackable`` under
                change tracking for one holding a whole ``w:altChunk`` or kept
                ``mc:AlternateContent``. Nothing has changed.
        """
        if self.of_control is not None:
            # a control's own range takes a control showing its placeholder as Word does (CR-019 section
            # 4.4, core-ts's D3 and D4)
            control, part = self.of_control
            if control._delete_through(part):
                self.end = self.start
                return
        elif self.whole and not self.mark and not self.spans_paragraphs:
            # a paragraph's 'Content' range over a block control's only paragraph takes the control, as
            # Word's took one showing its placeholder (CR-019 section 4.3, D3; check 63 cases 12 and 38)
            control = self.paragraph._placeholder_control
            if (
                control is not None
                and len(control.paragraphs) == 1
                and self.start == 0
                and self.end >= len(self.paragraph.text)
            ):
                landed = control._take_whole()
                self._land_at(landed.paragraph, 0)
                return
        if self.spans_paragraphs or self.holds_end_mark:
            self._delete_span()
            return
        with recording(self.paragraph.parent_body, "delete") as change:
            change.touched(self.paragraph)
            change.text(before=self.text, after="")
            # a range holding its paragraph's whole content takes what stands at its edges, an empty
            # paragraph's comment reference among them (core-ts's WHOLE, its CR-014 section 12.6)
            self.paragraph.splice(self.start, self.end, "", whole=self.mark or self.whole)
            self.end = self.start

    def _delete_span(self) -> None:
        """:meth:`delete` across paragraphs, or of a range holding its mark (:func:`~.span_edits.delete_span`)."""
        from docx4j_py.model.content.span_edits import (
            assert_trackable,
            delete_span,
            extent_of,
            refuse_tracked_content,
        )

        # an end in a kept branch is read-only before anything else is asked (CR-013 section 4.5)
        self._refuse_read_only("delete", ends_only=True)
        extent = extent_of(self, "delete", cross_controls=True)
        self._refuse_edge_pieces("delete")
        body = self.paragraph.parent_body
        if self.paragraph.change_tracker is not None:
            assert_trackable(extent, "delete")
            refuse_tracked_content(extent, "delete", body)
        with recording(body, "delete") as change:
            change.text(before=self.text, after="")
            change.touched(self.paragraph, self.end_paragraph)
            landing = delete_span(self, extent, False)
            if landing is not None:
                self._land_at(landing.paragraph, landing.offset)
                change.touched(landing.paragraph)

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
        """The span as XML: a copy of the paragraph trimmed to the span; across paragraphs, the blocks it holds.

        The paragraph is deep-copied (CR-001's ``deep_copy``, which never
        follows the parent pointer), the text outside the span is spliced out
        of the copy, **untracked** whatever the package's mode (core-ts
        ``202a01d``: until CR-013 Phase C the trim was a tracked splice under
        change tracking, so a word's XML held the paragraph's other words as
        ``w:del``), and the copy is serialised. The document is not touched.

        **Across paragraphs** (CR-013 Phase C; core-ts CR-014 section 13): the
        blocks :meth:`delete` takes, one after another --- the first paragraph
        from the range's start, every block between whole (a table, a content
        control), the last paragraph to the range's end, left out where the
        range ends at its start --- which :meth:`Body.insert_xml` takes back.

        Raises:
            SpanError: ``range.partial_table`` or ``range.partial_control``, a
                range starting or ending inside a table or content control it
                doesn't hold whole: a fragment of part of one isn't written.
        """
        return "".join(to_xml(block) for block in self.blocks())

    def blocks(self) -> list[Any]:
        """Detached copies of the blocks this range holds, its paragraphs trimmed to it. Extension.

        core-ts's ``blocksOfRange`` (``fragment.mts`` at ``202a01d``): within
        one paragraph, that paragraph trimmed to the range; across paragraphs,
        the blocks :meth:`delete` takes (:func:`~.span_edits.extent_of`), the
        first and last paragraphs trimmed. What :meth:`get_xml` and
        :meth:`to_api_script` write.

        Raises:
            SpanError: ``range.partial_table`` or ``range.partial_control``.
        """
        from docx4j_py.model.content.span_edits import extent_of

        if not self.spans_paragraphs:
            return [_trimmed(self.paragraph, self.start, self.end)]
        extent = extent_of(self, "get_xml")
        out: list[Any] = []
        for index in range(extent.first, extent.last + 1):
            block = extent.items[index]
            if index == extent.first and extent.head is not None:
                out.append(_trimmed(extent.head, self.start, len(extent.head.text)))
            elif index == extent.last and extent.tail is not None:
                if self.end > 0:
                    out.append(_trimmed(extent.tail, 0, self.end))
            else:
                copy = deep_copy(block)
                copy.parent = None
                out.append(copy)
        return out

    def to_api_script(self, **options: Any) -> str:
        """The content-API calls that would produce this span's runs (CR-003 section 3.11).

        Python source against a name ``body``, executable as it stands; see
        :func:`docx4j_py.model.content.api_script.to_api_script` for the
        options (``variable``, ``location``, ``pictures``, ``limit``,
        ``addresses``). Within one paragraph `variable` names a paragraph, as
        a range emits its runs only; **across paragraphs** (CR-013 Phase C;
        core-ts ``202a01d``) it names a body, and each block the range holds
        (:meth:`blocks`) is written as a body's block is.

        Raises:
            SpanError: ``range.partial_table`` or ``range.partial_control``, as
                :meth:`get_xml`.
        """
        from docx4j_py.model.content.api_script import to_api_script

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


def _unlimited(options: dict[str, Any]) -> dict[str, Any]:
    """A search's options without ``limit``: a replacement's refusal reads every hit (CR-016 Phase C)."""
    return {name: value for name, value in options.items() if name != "limit"}


def _trimmed(paragraph: Paragraph, start: int, end: int) -> Any:
    """A detached copy of `paragraph` holding only its text from `start` to `end`, with the runs and their formatting.

    Trimmed through a body that reads as the paragraph's own does (its
    package's styles decide which runs are hidden, and so the offsets) and
    writes untracked (core-ts's ``CopyBody``, ``202a01d``): trimming a copy is
    no edit of the document.
    """
    from docx4j_py.model.content.body import CopyBody
    from docx4j_py.model.content.paragraph import Paragraph as ParagraphView
    from docx4j_py.wml import Body as BodyElement

    copy = deep_copy(paragraph.element)
    holder = BodyElement(content=ChildList([copy]))
    copy.parent = holder
    body = CopyBody(
        paragraph.parent_body.part,
        holder,
        prefix="fragment",
        package=paragraph.parent_body.package,
    )
    view = ParagraphView(copy, holder.content, body)
    segments = view.segments()
    if any(segment.branch is not None or segment.math for segment in segments):
        _trim_around_branches(view, start, end)
    else:
        length = len(view.text)
        if end < length:
            view.splice(end, length, "")
        if start > 0:
            view.splice(0, start, "")
    copy.parent = None
    return copy


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


