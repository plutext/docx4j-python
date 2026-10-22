"""``Document.add_comment`` and the ``Comment`` it returns (python-docx 1.2).

CR-003 section 3.6, Phase J, over CR-003 Phase G (section 15). python-docx 1.2
added comments in the shape ``document.add_comment(runs, text, author,
initials)``: the range is delimited by the **first and last run** given, and a
single run serves as both. Phase G's ``Range.insert_comment`` is the verb
underneath, so all this module does is turn a pair of runs into a
:class:`~docx4j_py.model.content.Range`, put the author on the package for the
length of the call (Phase G reads ``pkg.author``, since there is no signed-in
user) and wrap the comment Phase G returns.

The one thing python-docx allows and this does not is a first and last run in
**different paragraphs**: Phase G's ``Range`` is one paragraph's, and the
markers go around a span inside it. Section 20 records it, with the message the
error carries.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Iterator, Sequence
from typing import TYPE_CHECKING, Any

from docx4j_py.docx._base import BlockItemContainer, Story
from docx4j_py.model.content.comments import Author
from docx4j_py.model.content.comments import Comment as ContentComment

if TYPE_CHECKING:  # pragma: no cover - annotations only
    from docx4j_py.docx.styles import ParagraphStyle
    from docx4j_py.docx.text import Paragraph, Run

__all__ = ["Comment", "Comments"]


class Comment(BlockItemContainer):
    """One ``w:comment``: python-docx 1.2's ``Comment``, a block container."""

    __slots__ = ("_view",)

    def __init__(self, view: ContentComment) -> None:
        """Wrap the content API's :class:`~docx4j_py.model.content.Comment`."""
        self._view = view

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:comment`` are equal."""
        return isinstance(other, Comment) and other._view == self._view

    def __hash__(self) -> int:
        """Hashes by the comment's element."""
        return hash(self._view)

    def __repr__(self) -> str:
        """``<docx4j_py.docx Comment 3 by 'Ada' 'Is this right?'>``."""
        text = self.text.replace("\n", " ")
        preview = text[:30] + "…" if len(text) > 30 else text
        return f"<docx4j_py.docx Comment {self.comment_id} by {self.author!r} {preview!r}>"

    @property
    def view(self) -> ContentComment:
        """The content API's comment. An extension: replies and resolution are on it."""
        return self._view

    @property
    def element(self) -> Any:
        """The typed ``w:comment``."""
        return self._view.element

    @property
    def _story(self) -> Story:
        """The comments part, as a story of its own."""
        return Story(self._view.comment_body)

    @property
    def _container(self) -> Any:
        """The ``w:comment``, whose block children are the comment's paragraphs."""
        return self._view.element

    @property
    def part(self) -> Any:
        """The comments part."""
        return self._view.parts.comments_part

    # -- metadata ----------------------------------------------------------

    @property
    def comment_id(self) -> int:
        """``w:comment/@w:id``."""
        return self._view.id

    @property
    def author(self) -> str:
        """``w:comment/@w:author``."""
        return self._view.author_name

    @author.setter
    def author(self, value: str) -> None:
        from docx4j_py.model.content.reports import recording_on

        with recording_on(
            self._view.body.package, "docx.comment.author", self.part
        ):
            self._view.element.author = value

    @property
    def initials(self) -> str | None:
        """``w:comment/@w:initials``."""
        return self._view.initials or None

    @initials.setter
    def initials(self, value: str | None) -> None:
        from docx4j_py.model.content.reports import recording_on

        with recording_on(self._view.body.package, "docx.comment.initials", self.part):
            self._view.element.initials = value

    @property
    def timestamp(self) -> dt.datetime | None:
        """``w:comment/@w:date``."""
        return self._view.creation_date

    @property
    def text(self) -> str:
        """The comment's text, a line per paragraph: python-docx's, not the text model's (CR-011 Q5).

        python-docx 1.2.0's ``Comment.text``: each paragraph's text (its runs and
        hyperlinks) a ``\\n`` apart. Not the content API's ``content``, whose
        paragraphs are Office JS's ``\\r`` apart since CR-013 Phase A.
        """
        from docx4j_py.docx.text import holder_text

        return "\n".join(holder_text(paragraph.element) for paragraph in self._view.paragraphs)

    # -- content -----------------------------------------------------------

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """A paragraph at the end of the comment; ``CommentText`` unless `style` says."""
        paragraph = super().add_paragraph(text, style)
        if style is None:
            paragraph.view.style_id = "CommentText"
        return paragraph


class Comments:
    """``document.comments``: the comments the document carries."""

    __slots__ = ("_document",)

    def __init__(self, document: Any) -> None:
        """Wrap the document whose comments these are."""
        self._document = document

    def __repr__(self) -> str:
        """``<docx4j_py.docx Comments, 3 comments>``."""
        return f"<docx4j_py.docx Comments, {len(self)} comments>"

    @property
    def _all(self) -> list[ContentComment]:
        """Every comment the part holds, replies included, by id.

        ``body.get_comments()`` nests replies under their parent, which is the
        right answer for reading a thread and the wrong one for python-docx's
        ``Comments``, which is flat.
        """
        from docx4j_py.openpackaging.parts.wml.comments import comment_parts_of

        main = self._document.part
        parts = comment_parts_of(main)
        if parts is None or parts.comments is None:
            return []
        body = self._document.part.body
        return [
            ContentComment(element, parts, body)
            for element in getattr(parts.comments, "comment", None) or ()
        ]

    def __len__(self) -> int:
        """How many comments there are."""
        return len(self._all)

    def __iter__(self) -> Iterator[Comment]:
        """Each comment, in the order the part holds them."""
        return (Comment(view) for view in self._all)

    def get(self, comment_id: int) -> Comment | None:
        """The comment with `comment_id`, or None."""
        for view in self._all:
            if view.id == comment_id:
                return Comment(view)
        return None

    def add_comment(
        self, text: str = "", author: str = "", initials: str | None = ""
    ) -> Comment:
        """A comment with no anchor in the body, as python-docx's collection adds one.

        Word shows nothing for a comment with no ``w:commentRangeStart``, so
        this is the half of python-docx's API that ``Document.add_comment``
        completes; the facade offers it because python-docx does, and the
        markers can be put in later with ``run.view`` and Phase G.
        """
        raise NotImplementedError(
            "docx4j_py.docx adds a comment through Document.add_comment(runs, text), "
            "which anchors it; Phase G has no verb for an unanchored comment "
            "(CR-003 section 20)"
        )


def add_comment(
    document: Any,
    runs: Run | Sequence[Run],
    text: str | None = "",
    author: str = "",
    initials: str | None = "",
) -> Comment:
    """What ``Document.add_comment`` is. See its docstring."""
    import re

    from docx4j_py.docx.text import Run as FacadeRun
    from docx4j_py.docx.text import check_python_docx_text
    from docx4j_py.model.content.comments import insert_comment_into
    from docx4j_py.model.content.errors import ContentError
    from docx4j_py.model.content.range import Range

    # python-docx 1.2.0's: a newline starts a paragraph, a carriage return is a
    # w:br in it, any other control character refused (CR-013 Phase A)
    check_python_docx_text(text)

    items = [runs] if isinstance(runs, FacadeRun) else list(runs)
    if not items:
        raise ContentError(
            "add_comment needs at least one run to anchor the comment on",
            code="comment.no_runs",
            hint="pass a Run, or paragraph.runs",
        )
    first, last = items[0].element, items[-1].element
    paragraph = items[0].view_paragraph
    if items[-1].view_paragraph.element is not paragraph.element:
        raise ContentError(
            "the first and last run of a comment must be in the same paragraph",
            code="comment.spans_paragraphs",
            hint=(
                "comment each paragraph separately, or use the content API: "
                "paragraph.view.insert_comment(text)"
            ),
        )
    segments = paragraph.segments()
    starts = [s.start for s in segments if s.run is first]
    ends = [s.end for s in segments if s.run is last]
    start = min(starts) if starts else 0
    end = max(ends) if ends else start
    package = paragraph.parent_body.package
    previous = getattr(package, "_author", None)
    if author or initials:
        package.author = Author(author, initials=initials)
    try:
        view = insert_comment_into(
            Range(paragraph, start, end), (text or "").replace("\r", "\x0b"), marks=re.compile("\n")
        )
    finally:
        package.author = previous
    return Comment(view)
