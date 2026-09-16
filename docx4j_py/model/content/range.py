"""``Range``: Office JS ``Word.Range``, a span within one paragraph.

CR-003 section 3.2 and decided question 3: the offsets are **code points** into
the paragraph's accepted-view text, and only a ``Range`` carries them. Positions
are recomputed from the tree on every use, so a range stays valid across the
edits made through it; an edit made elsewhere in the same paragraph may shift
it, as it does in Word.

Two rules of section 4 land here: a ``Range``'s original view is not offered
(``get_text(view="original")`` raises --- a range's offsets are accepted-view
offsets, and the original view has different ones), and :attr:`Range.font`
splits the runs at the boundaries so that a write touches the span and nothing
else. Phase B splits and does not merge back: two adjacent runs whose ``w:rPr``
ends up identical stay two runs, which Word tolerates and a later phase can
tidy.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, deep_copy
from docx4j_py.model.content.enums import (
    Location,
    ParagraphLocation,
    RangeLocation,
    TextView,
)
from docx4j_py.model.content.errors import ContentError, SpanError
from docx4j_py.model.content.font import Font
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.text_model import (
    RUN_HOLDER_NAMES,
    find_all,
    search_pattern,
)
from docx4j_py.traversal import element_name
from docx4j_py.wml import to_xml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = ["Range"]

#: How much of the text a ``repr`` shows.
_PREVIEW = 30


class Range:
    """A subset of Office JS ``Word.Range``: ``[start, end)`` in one paragraph."""

    __slots__ = ("end", "paragraph", "start")

    def __init__(self, paragraph: Paragraph, start: int, end: int) -> None:
        """Build the span over a paragraph, from `start` to `end` in code points."""
        #: The paragraph the span is in.
        self.paragraph = paragraph
        #: The offset of the first character.
        self.start = start
        #: One past the last character.
        self.end = end

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """The same span of the same paragraph."""
        return (
            isinstance(other, Range)
            and other.paragraph == self.paragraph
            and (other.start, other.end) == (self.start, self.end)
        )

    def __hash__(self) -> int:
        """Hashes by the paragraph's element and the two offsets."""
        return hash((id(self.paragraph.element), self.start, self.end))

    def __str__(self) -> str:
        """The span's text."""
        return self.text

    def __repr__(self) -> str:
        """``<Range 12:21 'quick brown'>``."""
        text = self.text
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        return f"<Range {self.start}:{self.end} {preview!r}>"

    def __len__(self) -> int:
        """The number of code points the span covers."""
        return max(0, self.end - self.start)

    # -- text --------------------------------------------------------------

    @property
    def text(self) -> str:
        """The span's text."""
        return self.paragraph.text[self.start : self.end]

    @text.setter
    def text(self, value: str) -> None:
        self.insert_text(value, location="Replace")

    def get_text(self, *, view: TextView = "accepted") -> str:
        """The span's text. The original view is not offered on a range.

        Raises:
            SpanError: `view` is ``"original"``; read that view on the
                paragraph, whose offsets it belongs to (CR-003 section 4).
        """
        if view == "original":
            raise SpanError(
                "a Range's offsets are accepted-view offsets, so it has no original view",
                code="range.no_original_view",
                hint="call get_text(view='original') on the paragraph instead",
            )
        return self.text

    @property
    def paragraphs(self) -> list[Paragraph]:
        """The paragraph this span is in, as a list (Office JS's shape)."""
        return [self.paragraph]

    @property
    def runs(self) -> list[Any]:
        """The runs the span covers; a run only partly inside counts."""
        out: list[Any] = []
        seen: set[int] = set()
        for segment in self.paragraph.segments():
            if segment.end <= self.start or segment.start >= self.end:
                continue
            if id(segment.run) in seen:
                continue
            seen.add(id(segment.run))
            out.append(segment.run)
        return out

    # -- style and formatting ---------------------------------------------

    @property
    def style(self) -> str:
        """The paragraph's style display name; Office JS reports that on a range."""
        return self.paragraph.style

    @style.setter
    def style(self, value: str) -> None:
        self.paragraph.style = value

    @property
    def style_id(self) -> str:
        """The paragraph's style id. Extension, as on :class:`Paragraph`."""
        return self.paragraph.style_id

    @style_id.setter
    def style_id(self, value: str) -> None:
        self.paragraph.style_id = value

    @property
    def style_built_in(self) -> str:
        """The paragraph's ``Word.Style`` value, or ``"Other"``."""
        return self.paragraph.style_built_in

    @style_built_in.setter
    def style_built_in(self, value: str) -> None:
        self.paragraph.style_built_in = value

    @property
    def font(self) -> Font:
        """The direct formatting of exactly this span.

        Asking for it splits nothing; a read reports the first run the span
        touches. A **write** splits the runs at the boundaries first --- the
        start backwards and the end forwards to the nearest grapheme boundary,
        so a cluster is never cut in two (CR-003 section 3.12) --- so that only
        the span is touched.
        """

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
        )

    # -- spans that cross a run holder (CR-003 section 3.4) ----------------

    def holder_boundaries(self) -> list[int]:
        """The offsets at which this span enters or leaves a run holder.

        A ``w:hyperlink``, a run-level ``w:sdt``, a ``w:ins`` and the other run
        holders own their runs; an operation that has to sit in exactly one of
        them --- a content control at range level (Phase C), a comment anchor
        (Phase G) --- cannot span two. These are the offsets to split at.
        """
        out: list[int] = []
        previous: Any = None
        for segment in self.paragraph.segments():
            if segment.end <= self.start or segment.start >= self.end:
                continue
            holder = getattr(segment.run, "parent", None)
            name = element_name(holder)
            key = id(holder) if name in RUN_HOLDER_NAMES else None
            if previous is not None and key != previous[0]:
                out.append(max(self.start, segment.start))
            previous = (key, holder)
        return out

    def require_one_holder(self, operation: str = "this operation") -> None:
        """Refuse a span that crosses a run holder, saying where to split.

        Raises:
            SpanError: the span covers runs in more than one holder; the message
                names the offsets to split at.
        """
        boundaries = self.holder_boundaries()
        if not boundaries:
            return
        where = ", ".join(str(offset) for offset in boundaries)
        raise SpanError(
            f"{operation} needs a span inside one run holder, and "
            f"{self.start}:{self.end} crosses one at {where}",
            code="range.crosses_holder",
            hint=f"split the span at {boundaries[0]}, or use get_range() on each part",
        )

    # -- editing -----------------------------------------------------------

    def insert_text(self, text: str, *, location: Location = "Replace") -> Range:
        """Insert text relative to the span.

        Args:
            text: what to insert.
            location: ``"Replace"`` (the default, as Office JS has it for a
                range), ``"Before"`` / ``"Start"`` for in front of the span, or
                ``"After"`` / ``"End"`` for behind it.

        Returns:
            The range of the inserted text. This range goes on covering the
            same text: a replacement resizes it, text before it shifts it, and
            text after it leaves it where it was.
        """
        if location in ("Replace", "Before", "Start", "After", "End"):
            with recording(self.paragraph.parent_body, "insert_text") as change:
                change.touched(self.paragraph)
                change.text(before=self.text)
                if location == "Replace":
                    out = self.paragraph.splice(self.start, self.end, text)
                    self.end = self.start + len(text)
                elif location in ("Before", "Start"):
                    out = self.paragraph.splice(self.start, self.start, text)
                    self.start += len(text)
                    self.end += len(text)
                else:
                    out = self.paragraph.splice(self.end, self.end, text)
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
        """A new paragraph before or after this span's paragraph."""
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
        """
        from docx4j_py.model.content.ooxml import insert_ooxml_into_range

        return insert_ooxml_into_range(self, ooxml, location=location)

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` within this span, as ranges of the paragraph."""
        base = self.start
        limit = options.pop("limit", None)
        hits = find_all(self.text, search_pattern(text, **options))
        if limit is not None:
            hits = hits[:limit]
        return [Range(self.paragraph, base + start, base + end) for start, end in hits]

    def replace_text(self, find: str, replace: str, **options: Any) -> int:
        """Replace every match within the span, last first; returns the count."""
        with recording(self.paragraph.parent_body, "replace_text") as change:
            matches = self.search(find, **options)
            for match in reversed(matches):
                match.insert_text(replace, location="Replace")
            if matches:
                change.touched(self.paragraph)
                change.text(before=find, after=replace)
            return len(matches)

    def find(self, text: str, *, context: int = 40, limit: int = 20, **options: Any) -> list[Any]:
        """Every match within this span, as hits with addresses and context."""
        from docx4j_py.model.content.reports import hit_for

        base = self.start
        hits = find_all(self.text, search_pattern(text, **options))
        return [
            hit_for(self.paragraph, base + start, base + end, context=context)
            for start, end in hits[:limit]
        ]

    # -- comments (CR-003 section 3.9, Phase G) -----------------------------

    def get_comments(self) -> list[Any]:
        """The comments whose markers overlap this span (Office JS ``getComments``)."""
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

        The markers are **hoisted out of a ``w:ins`` or ``w:del``** the anchor
        run sits in, because a comment is not a revision (CR-003 section 4).

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

        Returns:
            The new :class:`~docx4j_py.model.content.comments.Comment`.
        """
        from docx4j_py.model.content.comments import insert_comment_into

        return insert_comment_into(self, text)

    def delete(self) -> None:
        """Remove the span's text, leaving an empty range where it was."""
        with recording(self.paragraph.parent_body, "delete") as change:
            change.touched(self.paragraph)
            change.text(before=self.text, after="")
            self.paragraph.splice(self.start, self.end, "")
            self.end = self.start

    def get_range(self, location: RangeLocation = "Whole") -> Range:
        """A range over the whole span, or the empty one at either end."""
        if location == "Start":
            return Range(self.paragraph, self.start, self.start)
        if location == "End":
            return Range(self.paragraph, self.end, self.end)
        return Range(self.paragraph, self.start, self.end)

    def get_xml(self) -> str:
        """The span as a ``w:p``: a copy of the paragraph trimmed to the span.

        The paragraph is deep-copied (CR-001's ``deep_copy``, which never
        follows the parent pointer), the text outside the span is spliced out
        of the copy, and the copy is serialised. The document is not touched.
        """
        from docx4j_py.model.content.body import Body
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
        length = len(view.text)
        if self.end < length:
            view.splice(self.end, length, "")
        if self.start > 0:
            view.splice(0, self.start, "")
        return to_xml(copy)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: the address, the text and the two offsets. Extension."""
        return {
            "address": self.paragraph.address,
            "text": self.text,
            "start": self.start,
            "end": self.end,
        }
