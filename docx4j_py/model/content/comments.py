"""``Comment``: Office JS ``Word.Comment`` over the five comment parts.

CR-003 section 3.9 and section 4's "Comments" rules, Phase G. The view half;
the parts are :mod:`docx4j_py.openpackaging.parts.wml.comments`, which knows
nothing of this module (CR-003 section 5: the content layer imports the parts
layer and never the other way round).

::

    pkg.author = Author("Claude", initials="C")
    hit = pkg.find("quick brown fox")[0]
    hit.range(pkg.body).insert_comment("Changed because the source says 'red'.")
    pkg.body.get_comments()[0].reply("Agreed.")

What a comment *is*, in the file: a ``w:comment`` in ``/word/comments.xml``, a
``w:commentRangeStart`` / ``w:commentRangeEnd`` pair and a reference run in the
body, a ``w15:commentEx`` keyed by the ``w14:paraId`` of the comment's first
paragraph (which carries ``w15:done`` and, for a reply, ``w15:paraIdParent``),
a ``w16cid:commentId`` giving it a durable id, a ``w15:person`` recording its
author, and --- only when the document already has the part --- a
``w16cex:commentExtensible`` carrying a UTC date.

The rules of section 4 that land here:

* **one type for comments and replies.** Office JS has ``CommentReply`` with a
  subset of ``Comment``'s members; a reply is a ``Comment`` here, because it is
  one in the file. :attr:`Comment.parent` is the extension.
* :attr:`Comment.id` is the OOXML ``w:id``, an ``int``, not Office JS's opaque
  string; :attr:`~Comment.creation_date` may be None.
* :meth:`Comment.get_range` is a :class:`~docx4j_py.model.content.range.Range`
  **per paragraph**, and a comment whose markers sit at block level gets none.
  An empty range gets a reference run only, which is all Word requires.
* **no ``w16cex`` part is created**, only kept in step.
* the styles part is unmarshalled **only when it lacks the comment styles**.
* a comment is **not a revision**, and on text in an insertion it is written in
  **Word's own form** (CR-010 section 12.8; docx4j-core-ts CR-002 section 38
  item 1): the ``w:commentRangeStart`` inside the ``w:ins``, before the first
  commented run; the ``w:ins`` closed after the last one, the rest of it a
  second ``w:ins`` with a new id; the ``w:commentRangeEnd`` and the reference
  run between the two. The anchor stays narrow and the reference run is no
  insertion. Accepting the insertion leaves the comment on its text; rejecting
  it removes the comment, as current Word does. A ``w:del``, ``w:moveFrom`` or
  ``w:moveTo`` keeps the earlier hoist: the markers go outside it.
* comment ids are their own space, above the highest ``w:comment/@w:id`` and the
  highest id on a marker; revision ids (Phase F) are a different counter.
"""

from __future__ import annotations

import dataclasses
import datetime
import re
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, link_parents
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.styles import ensure_style, style_ids_of
from docx4j_py.model.content.text_model import (
    _REVISIONS,
    block_children_of,
    drop_unwritable,
    hidden_runs_of,
    segments_of,
    text_run,
)
from docx4j_py.model.values import integer, on_off
from docx4j_py.namespaces import WML_NS
from docx4j_py.openpackaging.parts.wml.comments import (
    CommentParts,
    comment_parts_of,
    create_comment_parts,
    durable_ids_in_use,
    remove_comment_id,
    remove_extensible,
    set_comment_id,
    set_extensible,
)
from docx4j_py.traversal import element_name, run_items_of
from docx4j_py.wml import P, R, to_xml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "COMMENT_REFERENCE_STYLE",
    "COMMENT_STYLES",
    "COMMENT_TEXT_STYLE",
    "DEFAULT_AUTHOR",
    "Author",
    "Comment",
    "CommentMarker",
    "author_of",
    "comments_of",
    "ensure_comment_styles",
    "initials_of",
    "insert_comment_into",
    "markers_of",
    "markers_of_paragraph",
    "next_comment_id",
]

#: The paragraph style Word gives a comment's own paragraphs.
COMMENT_TEXT_STYLE = "CommentText"
#: The character style of the reference run, in the comment and in the body.
COMMENT_REFERENCE_STYLE = "CommentReference"
#: What :func:`ensure_comment_styles` puts in a document that lacks them.
#: ``CommentTextChar`` is there because docx4j's ``KnownStyles.xml`` gives
#: ``CommentText`` a ``w:link`` to it, and a definition should not dangle.
COMMENT_STYLES: tuple[str, ...] = (COMMENT_TEXT_STYLE, "CommentTextChar", COMMENT_REFERENCE_STYLE)

#: How much of the comment's text a ``repr`` shows.
_PREVIEW = 30

#: How many ids are drawn before the generator is declared stuck.
_MAX_ATTEMPTS = 64


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


W_R = _w("r")
W_COMMENT_RANGE_START = _w("commentRangeStart")
W_COMMENT_RANGE_END = _w("commentRangeEnd")
W_COMMENT_REFERENCE = _w("commentReference")

#: The four run-level revision holders, by element name. A marker beside a run
#: in one is hoisted out of it (section 4: a comment is not a revision), except
#: the range markers of a comment on text in a ``w:ins``, which go where Word
#: puts them (:func:`_place_around`).
_REVISION_NAMES = frozenset(_REVISIONS)

W_INS = _w("ins")


# ---------------------------------------------------------------------------
# who the comment is by
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class Author:
    """Who a comment is by. ``pkg.author = Author("Claude", initials="C")``.

    There is no signed-in user here, so the package carries the identity
    (CR-003 section 3.4: "the audit trail is Word's own"). **Phase F's tracked
    changes read the same setting**, so setting it once covers both halves of
    the audit trail.
    """

    name: str
    """``w:comment/@w:author``, and ``w15:person/@w15:author``."""
    initials: str | None = None
    """``w:comment/@w:initials``; derived from the name when not given."""
    email: str | None = None
    """Written into ``w15:presenceInfo/@w15:userId``; Office JS ``authorEmail``."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return {"name": self.name, "initials": initials_of(self), "email": self.email}

    def __str__(self) -> str:
        """The author's name."""
        return self.name


def initials_of(author: Author) -> str:
    """The initials Word would show: one letter per word of the name, at most three."""
    if author.initials:
        return author.initials
    words = [word for word in author.name.split() if word]
    return "".join(word[0].upper() for word in words[:3])


def _default_author() -> Author:
    """``Author("docx4j-python")``: the ``Application`` a created document names."""
    from docx4j_py.openpackaging.packages.wordprocessingml_package import (
        APPLICATION_NAME,
    )

    return Author(APPLICATION_NAME)


#: The identity a package that has not been given one writes.
DEFAULT_AUTHOR = _default_author()


def author_of(package: Any) -> Author:
    """``pkg.author``, or :data:`DEFAULT_AUTHOR` when the package has none."""
    found = getattr(package, "author", None) if package is not None else None
    return found if isinstance(found, Author) else DEFAULT_AUTHOR


# ---------------------------------------------------------------------------
# the markers in the body
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class CommentMarker:
    """One ``w:commentRangeStart`` / ``End`` / ``w:commentReference`` in a body.

    Carries everything an edit needs: the item, the live list holding it, the
    run and paragraph it is in, and its offset in the paragraph's text.
    """

    id: int
    """The ``w:id`` it names, which is the comment's."""
    kind: str
    """``"start"``, ``"end"`` or ``"reference"``."""
    item: Any
    """The element itself."""
    owner: list
    """The live list holding :attr:`item`."""
    run: Any = None
    """For a reference, the ``w:r`` it is in."""
    run_owner: list | None = None
    """For a reference, the list holding that run."""
    paragraph: Any = None
    """The ``w:p`` the marker is in; None for a block-level marker."""
    paragraph_container: list | None = None
    """The block list holding that paragraph."""
    offset: int | None = None
    """The marker's offset in the paragraph's accepted text; None at block level."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the id, the kind and the offset."""
        return {"id": self.id, "kind": self.kind, "offset": self.offset}


def _index_of(items: list, element: Any) -> int:
    """The index of an element **by identity**; ``-1`` when it is not there.

    ``list.index`` compares by value, and two runs of the same text are equal
    in this model, so every lookup here is by identity.
    """
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def _marker_id(item: Any) -> int | None:
    """A marker's ``w:id`` as Word reads it (CR-021): ``"3.0"`` is 3; a refused form None."""
    return integer(getattr(item, "id", None))


def _markers_in_paragraph(
    out: list[CommentMarker], paragraph: Any, container: Any, hidden: Any = None
) -> None:
    """Collect one paragraph's markers, each with its offset in the paragraph's ``text``.

    The offsets are :func:`~.text_model.segments_of`'s, through its ``visit``
    hook, so that one reader counts (CR-013 Phase A, section 4.1): a note
    reference, a ``w:ptab``, an equation and hidden text count as
    ``Paragraph.text`` counts them, and `hidden` is the paragraph's
    (:func:`~.text_model.hidden_runs_of`). A marker inside a ``w:ins`` gets its
    offset in the **paragraph** (Word's form, CR-010 section 12.8); one in a
    ``w:del`` is not read, the offsets being the accepted view's; one in a kept
    branch is one branch's, and is not listed (CR-011 Phase F).
    """

    def visit(item: Any, at: int, owner: list, run: Any, run_owner: Any, branch: Any) -> None:
        if branch is not None:
            return
        name = element_name(item)
        if run is None:
            if name in (W_COMMENT_RANGE_START, W_COMMENT_RANGE_END):
                identifier = _marker_id(item)
                if identifier is not None:
                    out.append(
                        CommentMarker(
                            id=identifier,
                            kind="start" if name == W_COMMENT_RANGE_START else "end",
                            item=item,
                            owner=owner,
                            paragraph=paragraph,
                            paragraph_container=container,
                            offset=at,
                        )
                    )
            return
        if name == W_COMMENT_REFERENCE:
            identifier = _marker_id(item)
            if identifier is not None:
                out.append(
                    CommentMarker(
                        id=identifier,
                        kind="reference",
                        item=item,
                        owner=owner,
                        run=run,
                        run_owner=run_owner,
                        paragraph=paragraph,
                        paragraph_container=container,
                        offset=at,
                    )
                )

    segments_of(paragraph, hidden=hidden, visit=visit)


def markers_of_paragraph(element: P, container: list, resolver: Any = None) -> list[CommentMarker]:
    """Every comment marker in one paragraph, with its offset in the text.

    `resolver` is the package's property resolver, for the hidden runs the
    offsets leave out (CR-013 Phase A); without one, a run's own ``w:vanish``.
    """
    out: list[CommentMarker] = []
    _markers_in_paragraph(out, element, container, hidden_runs_of(resolver, element))
    return out


def markers_of(container: Any, resolver: Any = None) -> list[CommentMarker]:
    """Every comment marker under a container, in document order.

    A table, a row, a cell and a content control are all one
    :func:`~docx4j_py.model.content.text_model.block_children_of` step, so the
    walk has no special case for any of them. `resolver` is
    :func:`markers_of_paragraph`'s. `container` may be a block list itself (a
    story's, CR-013 Phase B).
    """
    out: list[CommentMarker] = []

    def visit(items: list) -> None:
        for item in items:
            if isinstance(item, P):
                _markers_in_paragraph(out, item, items, hidden_runs_of(resolver, item))
                continue
            name = element_name(item)
            if name in (W_COMMENT_RANGE_START, W_COMMENT_RANGE_END):
                identifier = _marker_id(item)
                if identifier is not None:
                    out.append(
                        CommentMarker(
                            id=identifier,
                            kind="start" if name == W_COMMENT_RANGE_START else "end",
                            item=item,
                            owner=items,
                        )
                    )
                continue
            children = block_children_of(item)
            if children is not None:
                visit(children)

    children = container if isinstance(container, list) else block_children_of(container)
    if children is not None:
        visit(children)
    return out


def comment_ids_of(container: Any) -> list[int]:
    """The comment ids under a container, in document order, first use first."""
    seen: set[int] = set()
    out: list[int] = []
    for marker in markers_of(container):
        if marker.id not in seen:
            seen.add(marker.id)
            out.append(marker.id)
    return out


def _remove_markers(container: Any, identifier: int) -> None:
    """Remove every marker of a comment, and a reference run left empty."""
    for marker in markers_of(container):
        if marker.id != identifier:
            continue
        index = _index_of(marker.owner, marker.item)
        if index >= 0:
            del marker.owner[index]
        if (
            marker.kind == "reference"
            and marker.run is not None
            and marker.run_owner is not None
            and not getattr(marker.run, "content", None)
        ):
            index = _index_of(marker.run_owner, marker.run)
            if index >= 0:
                del marker.run_owner[index]


# ---------------------------------------------------------------------------
# ids
# ---------------------------------------------------------------------------


def next_comment_id(comments: Any, container: Any) -> int:
    """The next free ``w:id`` for a comment: above every comment and every marker.

    Derived from the document's own state rather than from
    ``pkg.id_generator()``, as Phase C's part names and relationship ids are
    (CR-003 section 14.2 item 10): the same document and the same calls give the
    same bytes, and the numbers stay readable. **A separate space from the
    revision ids** Phase F allocates (section 4).
    """
    highest = -1
    for comment in getattr(comments, "comment", None) or ():
        value = getattr(comment, "id", None)
        if isinstance(value, int) and value > highest:
            highest = value
    for marker in markers_of(container):
        highest = max(highest, marker.id)
    return highest + 1


def _para_ids_taken(parts: CommentParts, body: Body) -> set[str]:
    """Every ``w14:paraId`` a new comment paragraph must not collide with.

    The body's own cached set (CR-003 section 12.2 item 4), widened with the ids
    the comments part already uses --- comment paragraphs and body paragraphs
    share one id space in Word, and the w15 and w16cid entries are keyed by it.
    """
    from docx4j_py.model.content.addresses import _taken

    taken = _taken(body.package, body)
    for comment in getattr(parts.comments, "comment", None) or ():
        for item in getattr(comment, "content", None) or ():
            para_id = getattr(item, "para_id", None)
            if para_id:
                taken.add(str(para_id).upper())
    return taken


def _hex_id(package: Any, taken: set[str], derive_from: Any = ()) -> str:
    """An eight-digit hexadecimal id, as Word writes ``w14:paraId``.

    From the package's seedable generator, so that the same seed and the same
    calls give the same bytes (CR-003 section 3.4).
    """
    generator = package.id_generator(derive_from=derive_from or taken)
    for _attempt in range(_MAX_ATTEMPTS):
        candidate = f"{generator.randrange(1, 0x7FFFFFFF):08X}"
        if candidate not in taken:
            taken.add(candidate)
            return candidate
    raise ContentError(  # pragma: no cover - 64 collisions in a 2^31 space
        "could not allocate a paragraph id",
        code="comment.no_id",
        hint="set pkg.id_seed to a different value",
    )


# ---------------------------------------------------------------------------
# the styles (CR-003 section 14.9: a built-in style the document lacks)
# ---------------------------------------------------------------------------


def ensure_comment_styles(package: Any) -> tuple[str, ...]:
    """Define the comment styles the document lacks; returns the parts touched.

    ``CommentText``, ``CommentTextChar`` and ``CommentReference``, from docx4j's
    ``KnownStyles.xml`` through
    :func:`~docx4j_py.model.content.styles.ensure_style`, which is the rule
    CR-003 section 14.9 settled: Word renders a dangling ``w:pStyle`` as Normal,
    so a style that is referred to has to be defined.

    **The styles part is unmarshalled only when it lacks one of them**
    (CR-003 section 4): the ids are read from the part's bytes with lxml first,
    so a document that already has the styles keeps ``/word/styles.xml`` byte
    for byte.
    """
    from docx4j_py.model.content.errors import StyleError

    part = getattr(package, "style_definitions_part", None) if package is not None else None
    if part is None:
        return ()
    defined = style_ids_of(part)
    if all(style_id in defined for style_id in COMMENT_STYLES):
        return ()
    touched: set[str] = set()
    for style_id in COMMENT_STYLES:
        try:
            ensure_style(package, style_id, touched=touched, defined=defined)
        except StyleError as error:  # pragma: no cover - KnownStyles.xml has all three
            if error.code != "style.not_creatable":
                raise
    return tuple(sorted(touched))


# ---------------------------------------------------------------------------
# building the pieces Word writes
# ---------------------------------------------------------------------------


def _reference_style_run(content: list) -> R:
    """A run in the ``CommentReference`` character style."""
    from docx4j_py.wml import el

    run = R(r_pr=el.rPr(r_style=el.rStyle(val=COMMENT_REFERENCE_STYLE)), content=ChildList(content))
    link_parents(run)
    return run


def _reference_run(identifier: int) -> R:
    """The run that anchors a comment in the body: ``w:commentReference``."""
    from docx4j_py.wml import el

    return _reference_style_run([el.commentReference(id=identifier)])


def _annotation_ref_run() -> R:
    """The run Word puts first in a comment: ``w:annotationRef``."""
    from docx4j_py.wml import el

    return _reference_style_run([el.annotationRef()])


def _resolver_of(body: Any) -> Any:
    """A body's package's property resolver, for the hidden runs a marker's offset leaves out (CR-013)."""
    from docx4j_py.model.property_resolver import resolver_of

    package = getattr(body, "package", None)
    return resolver_of(package) if package is not None else None


def _comment_paragraph(text: str, first: bool) -> P:
    """A paragraph of a comment: ``CommentText``, opened by the annotation reference.

    The text through the writers' map (CR-013 Q2): a line break U+000B a
    ``w:br``, and the rest of Office JS's characters their elements.
    """
    from docx4j_py.wml import el

    content: list[Any] = []
    if first:
        content.append(_annotation_ref_run())
    if text:
        content.append(text_run(text))
    paragraph = P(p_pr=el.pPr(p_style=el.pStyle(val=COMMENT_TEXT_STYLE)), content=ChildList(content))
    link_parents(paragraph)
    return paragraph


#: A comment's paragraph marks in its text: ``\r`` as ``content`` reads them, ``\n``
#: as 0.1.0 wrote them, and ``\r\n`` as one (CR-013 section 4.1; core-ts's D5).
_COMMENT_MARKS = re.compile("\r\n|\r|\n")


def _comment_paragraphs(text: str, *, marks: Any = None) -> list[P]:
    """One paragraph per line, the first opened by the annotation reference run.

    Both ``\\r`` and ``\\n`` are paragraph marks, so that ``Comment.content``
    written back makes the paragraphs it read (CR-013 section 4.1); `marks`, a
    compiled pattern, splits otherwise (the python-docx facade's ``\\n``).
    """
    lines = (marks or _COMMENT_MARKS).split(text)
    return [_comment_paragraph(line, index == 0) for index, line in enumerate(lines)]


def _append(owner: Any, field: str, item: Any) -> None:
    """Append to a model list and wire the parents, as every verb here does."""
    items = getattr(owner, field, None)
    if items is None:
        items = ChildList([], owner=owner, field=field)
        setattr(owner, field, items)
    items.append(item)
    link_parents(item)
    item.parent = owner


def _insert_at(owner: list, index: int, items: list, parent: Any) -> None:
    """Insert run-level items into a live list and wire their parents."""
    for offset, item in enumerate(items):
        owner.insert(index + offset, item)
        link_parents(item)
        item.parent = parent


def _marker_site(paragraph: Paragraph, segment: Any, *, after: bool) -> tuple[list, int, Any]:
    """Where a marker for a segment's run goes when it is not put inside a ``w:ins``.

    Beside the run, but **outside** the ``w:ins``, ``w:del``, ``w:moveFrom`` or
    ``w:moveTo`` the run sits in (section 4: a comment is not a revision, and a
    reference run inside one would be an insertion of its own). Since CR-010
    section 12.8 :func:`_place_around` asks this only for an empty range's
    reference run, and for a range marker whose run is not directly in a
    ``w:ins``: the range markers of a comment on inserted text go where Word
    puts them (:func:`_insertion_of`, :func:`_close_insertion`). A deletion's
    form is not measured and a move's halves are not to be split, so those keep
    this hoist (docx4j-core-ts CR-002 section 38 item 1).
    """
    # beside the run, or beside the kept element its text is read from (CR-011 Phase F)
    site, owner, site_index = segment.site()
    holder = getattr(site, "parent", None)
    if element_name(holder) in _REVISION_NAMES:
        outer = getattr(holder, "parent", None)
        items = run_items_of(outer) if outer is not None else None
        if items is not None:
            index = _index_of(items, holder)
            if index >= 0:
                return items, index + (1 if after else 0), outer
    parent = holder or paragraph.element
    return owner, site_index + (1 if after else 0), parent


def _insertion_of(segment: Any) -> Any:
    """The ``w:ins`` a segment's run sits **directly** in, or None.

    Directly: the run is an item of the ``w:ins``'s own list, which is checked
    by identity as well as through the run's ``parent``, so that a stale pointer
    (CR-003 section 16.9) cannot send a marker into the wrong list. A run in a
    hyperlink or a content control inside a ``w:ins`` is not directly in it.
    """
    holder = getattr(segment.run, "parent", None)
    if element_name(holder) != W_INS or run_items_of(holder) is not segment.run_owner:
        return None
    return holder


def _close_insertion(paragraph: Paragraph, insertion: Any, index: int, items: list) -> bool:
    """Close a ``w:ins`` before its item at `index` and put `items` after it.

    Word's form for the end of a comment on inserted text (CR-010 section 12.8,
    measured in docx4j's ``revisions/editor/review/measure-word15.docx``): the
    ``w:ins``'s items from `index` on move into a second ``w:ins`` with the same
    author and date and a **new** ``w:id``, and `items` go between the two; with
    nothing from `index` on, `items` simply go after the ``w:ins``. docx4j-core-ts's
    ``Paragraph.splitHolderAround``.

    The new id comes from the one annotation id space, through the package's
    tracker, or one made for the call while tracking is off (a comment on an
    insertion someone else made takes Word's form either way). False, with
    nothing changed, when the ``w:ins`` is not in its parent's list.
    """
    from docx4j_py.model.content.hyperlinks import _id_source, _revision_like

    # every parent read before anything moves (CR-003 section 16.9)
    outer_parent = getattr(insertion, "parent", None)
    outer = run_items_of(outer_parent) if outer_parent is not None else None
    at = _index_of(outer, insertion) if outer is not None else -1
    content = run_items_of(insertion)
    if outer is None or at < 0 or content is None:
        return False
    pieces = list(items)
    if index < len(content):
        tail = list(content[index:])
        # the first half lets go of the tail before the second adopts it
        del content[index:]
        pieces.append(_revision_like(insertion, tail, _id_source(paragraph)))
    _insert_at(outer, at + 1, pieces, outer_parent)
    return True


def _place_around(span: Range, identifier: int) -> None:
    """Write the markers of a new comment around a span, splitting at its boundaries.

    On text in a ``w:ins``, Word's form (CR-010 section 12.8): the start inside
    the ``w:ins`` before the first commented run, the end and the reference run
    after the ``w:ins`` that holds the last, which is split there when more of
    it follows. Elsewhere, and in a ``w:del`` or a move, the markers go beside
    the runs, outside any revision (:func:`_marker_site`).
    """
    if span.spans_paragraphs:
        _place_across(span, identifier)
        return
    paragraph = span.paragraph
    # a run holding the span holds nothing else, a tab or a break beside it moved out
    # too, as a font write splits (core-ts's splitRunsAround, its CR-014 section 11.2;
    # CR-013 Phase C): until then a marker went beside the whole run, widening the anchor
    if span.end > span.start:
        start, end = paragraph.split_runs_around(span.start, span.end)
    else:
        start = end = paragraph.split_at(span.start, prefer="back")
    segments = paragraph.segments()
    covered = [s for s in segments if s.end > start and s.start < end] if end > start else []

    from docx4j_py.wml import el

    reference = _reference_run(identifier)
    if not covered:
        # an empty span or an empty paragraph: the reference run only, which is
        # all Word requires (docx4j's CommentsSample says the same)
        after = next((s for s in segments if s.start >= start), None)
        before = next((s for s in reversed(segments) if s.end <= start), None)
        if after is not None:
            owner, index, parent = _marker_site(paragraph, after, after=False)
        elif before is not None:
            owner, index, parent = _marker_site(paragraph, before, after=True)
        else:
            owner, index, parent = paragraph.element.content, len(paragraph.element.content), (
                paragraph.element
            )
        _insert_at(owner, index, [reference], parent)
        return

    first, last = covered[0], covered[-1]
    # both looked up before the split moves anything
    first_insertion, last_insertion = _insertion_of(first), _insertion_of(last)
    closing = [el.commentRangeEnd(id=identifier), reference]
    if last_insertion is None or not _close_insertion(
        paragraph, last_insertion, last.run_index + 1, closing
    ):
        owner, index, parent = _marker_site(paragraph, last, after=True)
        _insert_at(owner, index, closing, parent)
    # the split moved only what follows the last run, so the first's index holds
    opening = el.commentRangeStart(id=identifier)
    if first_insertion is not None:
        _insert_at(first.run_owner, first.run_index, [opening], first_insertion)
    else:
        owner, index, parent = _marker_site(paragraph, first, after=False)
        _insert_at(owner, index, [opening], parent)


def _place_across(span: Range, identifier: int) -> None:
    """The markers of a comment over a range across paragraphs, as Word writes them.

    docx4j-core-ts's ``placeAcross`` (``Comment.mts`` at ``87bbb38``; Word's
    check 52 cases 79 to 85; CR-013 Phase C): the start before the range's
    first character, in its first paragraph, and the end and the reference run
    after its last character, in its last paragraph, before a mark the range
    holds. A table between, or a cell either end is in, stays as it is: the
    markers stand in the paragraphs. On a block control's content range Word
    put the start before the control's first paragraph, in ``w:sdtContent``
    (case 82); here it is inside that paragraph, which reads the same (core-ts
    CR-014 section 9.2). An insertion at either end is treated as
    :func:`_place_around` treats one: the start inside it, before the first
    commented run; the end after the last, the rest of the insertion split off.
    The runs are split as a ``font`` write splits them, so a marker never goes
    beside a run holding text outside the range.
    """
    from docx4j_py.wml import el

    first = span.paragraph
    last = span.end_paragraph
    opening = el.commentRangeStart(id=identifier)
    closing = [el.commentRangeEnd(id=identifier), _reference_run(identifier)]
    # the end first; the two are in different paragraphs, so neither's offsets move with the other's markers
    end = span.end
    if end > 0:
        _start, end = last.split_runs_around(0, end)
    before = [s for s in last.segments() if s.start < end]
    last_segment = before[-1] if before else None
    insertion = _insertion_of(last_segment) if last_segment is not None else None
    if insertion is None or not _close_insertion(last, insertion, last_segment.run_index + 1, closing):
        if last_segment is not None:
            owner, index, parent = _marker_site(last, last_segment, after=True)
        else:
            # nothing of the last paragraph in the range: the end at its start
            segments = last.segments()
            if segments:
                owner, index, parent = _marker_site(last, segments[0], after=False)
            else:
                owner, index, parent = last.element.content, 0, last.element
        _insert_at(owner, index, closing, parent)
    start = span.start
    length = len(first.text)
    if start < length:
        start, _end = first.split_runs_around(start, length)
    segments = first.segments()
    first_segment = next((s for s in segments if s.end > start), None)
    insertion = _insertion_of(first_segment) if first_segment is not None else None
    if first_segment is not None and insertion is not None:
        _insert_at(first_segment.run_owner, first_segment.run_index, [opening], insertion)
        return
    if first_segment is not None:
        owner, index, parent = _marker_site(first, first_segment, after=False)
    elif segments:
        # nothing of the first paragraph in the range: the start at its end
        owner, index, parent = _marker_site(first, segments[-1], after=True)
    else:
        owner, index, parent = first.element.content, len(first.element.content), first.element
    _insert_at(owner, index, [opening], parent)


def _place_after(body: Body, parent_id: int, identifier: int) -> None:
    """Write the markers of a reply beside its parent's, as Word nests them.

    The three insertions are made **last first** --- the reference run, then the
    ``w:commentRangeEnd``, then the ``w:commentRangeStart`` --- because the three
    are usually in one list and an insertion in front of a marker would move the
    ones after it.
    """
    from docx4j_py.wml import el

    markers = [m for m in markers_of(body.container) if m.id == parent_id]
    start = next((m for m in markers if m.kind == "start"), None)
    end = next((m for m in reversed(markers) if m.kind == "end"), None)
    reference = next((m for m in reversed(markers) if m.kind == "reference"), None)
    if reference is not None and reference.run_owner is not None:
        index = _index_of(reference.run_owner, reference.run) + 1
        _insert_at(
            reference.run_owner, index, [_reference_run(identifier)], _parent_of(reference.run)
        )
    elif end is not None:
        index = _index_of(end.owner, end.item) + 1
        _insert_at(end.owner, index, [_reference_run(identifier)], _parent_of(end.item))
    else:
        raise ContentError(
            f"comment {parent_id} has no markers in this body, so a reply cannot be anchored",
            code="comment.no_anchor",
            hint="reply to a comment that is anchored in this body, or insert a new one",
        )
    if end is not None:
        index = _index_of(end.owner, end.item) + 1
        _insert_at(end.owner, index, [el.commentRangeEnd(id=identifier)], _parent_of(end.item))
    if start is not None:
        index = _index_of(start.owner, start.item) + 1
        _insert_at(
            start.owner, index, [el.commentRangeStart(id=identifier)], _parent_of(start.item)
        )


def _parent_of(element: Any) -> Any:
    return getattr(element, "parent", None)


# ---------------------------------------------------------------------------
# the view
# ---------------------------------------------------------------------------


class Comment:
    """A subset of Office JS ``Word.Comment`` over a ``w:comment`` and its markers."""

    __slots__ = ("_body_view", "_parent", "_replies", "body", "element", "parts")

    def __init__(self, element: Any, parts: CommentParts, body: Body) -> None:
        """Build the view over the ``w:comment``, the parts and the anchoring body."""
        #: The ``w:comment``. Extension.
        self.element = element
        #: The document's comment parts.
        self.parts = parts
        #: The :class:`~docx4j_py.model.content.body.Body` the comment is anchored in.
        self.body = body
        self._parent: Comment | None = None
        self._replies: list[Comment] = []
        self._body_view: Any = None

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:comment`` are equal (CR-003 section 3.1)."""
        return isinstance(other, Comment) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self.element))

    def __str__(self) -> str:
        """The comment's text."""
        return self.content

    def __repr__(self) -> str:
        """``<Comment 3 body/2 'Claude' 'Changed because the …'>``."""
        text = self.content
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        where = self.anchor_address or "-"
        flag = " resolved" if self.resolved else ""
        return f"<Comment {self.id} {where} {self.author_name!r} {preview!r}{flag}>"

    # -- reading -----------------------------------------------------------

    @property
    def id(self) -> int:
        """``w:id``. Extension: Office JS's ``id`` is an opaque string."""
        value = integer(getattr(self.element, "id", None))  # as Word reads it (CR-021)
        return value if value is not None else -1

    @property
    def author_name(self) -> str:
        """``w:author``, as Office JS's ``authorName``."""
        return getattr(self.element, "author", None) or ""

    @property
    def initials(self) -> str:
        """``w:initials``. Extension: Office JS has no initials."""
        return getattr(self.element, "initials", None) or ""

    @property
    def author_email(self) -> str:
        """The author's email from ``w:people``; ``""`` when none is recorded.

        Word stores it as ``w15:presenceInfo/@w15:userId``, and for an Active
        Directory provider as ``S::name@example.com::<guid>``; the address is
        taken out of that form.
        """
        for person in getattr(self.parts.people, "person", None) or ():
            if (getattr(person, "author", None) or "") != self.author_name:
                continue
            info = getattr(person, "presence_info", None)
            user_id = (getattr(info, "user_id", None) if info is not None else None) or ""
            if user_id.startswith("S::"):
                rest = user_id[3:]
                return rest.partition("::")[0]
            return user_id
        return ""

    @property
    def creation_date(self) -> datetime.datetime | None:
        """When the comment was made, in UTC, as Word means it; None when it carries no date.

        Its ``w16cex:dateUtc`` when the document has one for it, which is UTC;
        else ``w:date``, which Word writes on the author's local wall clock
        whatever its ``Z`` says, read as local time in this process's zone
        (CR-010 section 12.9; docx4j-core-ts ``Comment.creationDate``). The
        extensible entry is found through the paragraph ids, the last paragraph's
        first, which is the one Word keys it by.
        """
        from docx4j_py.model.revisions.change import word_date
        from docx4j_py.openpackaging.parts.wml.comments import dates_utc

        dates = dates_utc(self.parts) if self.parts is not None else {}
        if dates:
            for item in reversed(getattr(self.element, "content", None) or ()):
                para_id = item.para_id if isinstance(item, P) else None
                found = dates.get(para_id.upper()) if para_id else None
                if found is not None:
                    return found
        return word_date(getattr(self.element, "date", None))

    @property
    def para_id(self) -> str | None:
        """The ``w14:paraId`` of the first paragraph: the w15 and w16cid key. Extension."""
        for item in getattr(self.element, "content", None) or ():
            if isinstance(item, P):
                return item.para_id
        return None

    @property
    def comment_body(self) -> Any:
        """The comment's own paragraphs as a :class:`Body`. Extension.

        The whole content API inside a comment: ``comment.comment_body.text``,
        ``to_markdown()``, ``insert_paragraph`` and the rest.
        """
        from docx4j_py.model.content.body import Body

        if self._body_view is None:
            self._body_view = Body(
                self.parts.comments_part,
                self.element,
                prefix=f"comment:{self.id}",
                package=self.body.package,
            )
        return self._body_view

    @property
    def paragraphs(self) -> list[Paragraph]:
        """The comment's paragraphs. Extension."""
        return self.comment_body.paragraphs

    @property
    def content(self) -> str:
        """The comment's text as Office JS's ``comment.content`` gives it. Setting it replaces the paragraphs.

        Its body's ``Body.text``: the paragraphs a mark ``\\r`` apart, in Office
        JS's characters (CR-013 Phase A; core-ts check 51). Written, both
        ``\\r`` and ``\\n`` are paragraph marks, ``\\r\\n`` one, as Word's
        ``comment.content`` makes them (core-ts check 55, cases 43 to 45), so
        that the content read and written back makes the paragraphs it read
        (CR-013 section 4.1); each other character of Office JS's is the
        element it reads from (Q2), and U+0002 and the other characters XML 1.0
        excludes are dropped (section 11, E3).

        Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text. It reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, CR-011 section 15).
        """
        return self.comment_body.text

    @content.setter
    def content(self, text: str) -> None:
        text = drop_unwritable(text)  # CR-013 section 11, E3: what no element carries is dropped
        with recording(self.body, "set_comment_content") as change:
            change.text(before=self.content, after=text)
            para_id = self.para_id
            paragraphs = _comment_paragraphs(text)
            if para_id:
                paragraphs[0].para_id = para_id
            self.element.content = ChildList(paragraphs, owner=self.element, field="content")
            link_parents(self.element)
            self._body_view = None
            self._report(change)

    @property
    def resolved(self) -> bool:
        """``w15:done``: whether the thread is marked resolved.

        Setting it creates ``/word/commentsExtended.xml`` when the document has
        none, and gives the comment a ``w14:paraId`` when it has none, because
        that is the key the entry is filed under.
        """
        entry = self._comment_ex()
        return on_off(getattr(entry, "done", None) if entry is not None else None, default=False)

    @resolved.setter
    def resolved(self, value: bool) -> None:
        from docx4j_py.w15 import el as w15el

        with recording(self.body, "resolve_comment" if value else "reopen_comment") as change:
            change.text(before=f"resolved={self.resolved}")
            if value:
                # CR-003 sections 15.7 and 17.11: w15:done is Word 2013's, and a
                # Word 2010 document loses it on save with a Compatibility
                # Checker note ("comments which have been collapsed")
                from docx4j_py.model.content.compatibility import warn_below

                warn_below(
                    getattr(self.body, "package", None),
                    change,
                    "a resolved comment (w15:done)",
                )
            parts = self._writable()
            para_id = self._ensure_para_id()
            entry = self._comment_ex()
            if entry is None:
                entry = w15el.commentEx(para_id=para_id)
                _append(parts.comments_ex, "comment_ex", entry)
            entry.done = "1" if value else "0"
            change.text(after=f"resolved={value}")
            self._report(change)

    @property
    def replies(self) -> list[Comment]:
        """The replies to this comment, in document order. Office JS ``replies``.

        Threaded through ``w15:paraIdParent``. Filled in by
        :meth:`Body.get_comments` and by :meth:`reply`; a view built any other
        way answers with what it has been told.
        """
        return list(self._replies)

    @property
    def parent(self) -> Comment | None:
        """The comment this one replies to, or None. Extension."""
        return self._parent

    @property
    def anchor_address(self) -> str | None:
        """The address of the paragraph this comment is anchored in, or None. Extension."""
        for marker in markers_of(self.body.container):
            if marker.id == self.id and marker.paragraph is not None:
                return self.body.paragraph_for(marker.paragraph).address
        return None

    def get_range(self) -> Range | None:
        """The commented range: one :class:`Range` from the comment's start to its end (Office JS ``getRange``).

        CR-013 Phase B (D2, Q3; core-ts's at ``ef57729``): across paragraphs
        where the comment spans them, as Office JS's (core-ts check 52: a
        comment from one paragraph into the next reads ``BODY1B\\rLIST1A``). A
        start marker between blocks (Word writes one before a control's first
        paragraph) starts the range at the next paragraph, and an end marker
        between blocks ends it at the previous one's end. None when the
        comment has no markers in this body. Until 0.2.0 this returned a list
        a paragraph at a time, which is :meth:`get_ranges` now: a caller that
        indexed ``get_range()[0]`` reads ``get_range()``, or ``get_ranges()``.
        """
        from docx4j_py.model.content.range import Range

        markers = [m for m in markers_of(self.body.container, _resolver_of(self.body)) if m.id == self.id]
        start = next((m for m in markers if m.kind == "start"), None)
        end = next((m for m in reversed(markers) if m.kind == "end"), None)
        begins = self._anchor(start, "start") if start is not None else None
        ends = self._anchor(end, "end") if end is not None else None
        if begins is None or ends is None:
            ranges = self.get_ranges()
            return ranges[0] if ranges else None
        paragraphs = self.body.paragraphs
        index_of = {id(p.element): i for i, p in enumerate(paragraphs)}
        a = index_of.get(id(begins[0].element))
        b = index_of.get(id(ends[0].element))
        if a is None or b is None or b < a or (a == b and ends[1] < begins[1]):
            ranges = self.get_ranges()
            return ranges[0] if ranges else None
        if a == b:
            return Range(begins[0], begins[1], ends[1])
        return Range(begins[0], begins[1], ends[1], end_paragraph=ends[0])

    def _anchor(self, marker: CommentMarker, side: str) -> tuple[Paragraph, int] | None:
        """Where a range marker stands: its paragraph and offset; a block-level one's the next or the previous paragraph's edge."""
        from docx4j_py.model.content.paragraph import Paragraph
        from docx4j_py.model.content.story import first_and_last_paragraph

        if marker.paragraph is not None and marker.paragraph_container is not None:
            paragraph = Paragraph(marker.paragraph, marker.paragraph_container, self.body)
            if marker.offset is not None:
                return paragraph, marker.offset
            return paragraph, 0 if side == "start" else len(paragraph.text)
        at = _index_of(marker.owner, marker.item)
        if at < 0:
            return None
        indices = range(at + 1, len(marker.owner)) if side == "start" else range(at - 1, -1, -1)
        for index in indices:
            first, last = first_and_last_paragraph(marker.owner, index)
            place = first if side == "start" else last
            if place is None:
                continue
            paragraph = Paragraph(place.element, place.container, self.body)
            return paragraph, 0 if side == "start" else len(paragraph.text)
        return None

    def get_ranges(self) -> list[Range]:
        """The commented span a paragraph at a time: a :class:`Range` **per paragraph** it touches. Extension.

        What :meth:`get_range` returned until CR-013 Phase B (D2; core-ts's
        ``getRanges``). Empty for a comment whose markers sit at block level,
        and for one whose markers are in another body.
        """
        from docx4j_py.model.content.range import Range

        markers = [
            m
            for m in markers_of(self.body.container, _resolver_of(self.body))
            if m.id == self.id and m.paragraph is not None
        ]
        if not markers:
            return []
        paragraphs = self.body.paragraphs
        index_of = {id(p.element): i for i, p in enumerate(paragraphs)}
        start = next((m for m in markers if m.kind == "start"), None)
        end = next((m for m in reversed(markers) if m.kind == "end"), None)
        if start is None or end is None:
            marker = start or end or markers[0]
            position = index_of.get(id(marker.paragraph))
            if position is None:
                return []
            paragraph = paragraphs[position]
            at = marker.offset or 0
            return [Range(paragraph, at, len(paragraph.text) if marker.kind == "start" else at)]
        first = index_of.get(id(start.paragraph))
        last = index_of.get(id(end.paragraph))
        if first is None or last is None or last < first:
            return []
        out: list[Range] = []
        for position in range(first, last + 1):
            paragraph = paragraphs[position]
            out.append(
                Range(
                    paragraph,
                    start.offset or 0 if position == first else 0,
                    (end.offset if end.offset is not None else len(paragraph.text))
                    if position == last
                    else len(paragraph.text),
                )
            )
        return out

    def get_xml(self) -> str:
        """The ``w:comment`` as XML, with docx4j's prefixes. Extension."""
        return to_xml(self.element)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a comment thread."""
        out: dict[str, Any] = {
            "id": self.id,
            "author_name": self.author_name,
            "initials": self.initials,
            "content": self.content,
        }
        email = self.author_email
        if email:
            out["author_email"] = email
        date = self.creation_date
        if date is not None:
            out["creation_date"] = date.isoformat()
        if self.para_id:
            out["para_id"] = self.para_id
        address = self.anchor_address
        if address:
            out["address"] = address
        if self.resolved:
            out["resolved"] = True
        if self._replies:
            out["replies"] = [reply.to_dict() for reply in self._replies]
        return out

    # -- editing -----------------------------------------------------------

    def reply(self, text: str) -> Comment:
        """A reply in the same thread, anchored beside this comment's markers.

        A reply is a ``Comment`` of its own in the file, linked to this one by
        ``w15:paraIdParent``, which is how Word threads them.

        Raises:
            ContentError: this comment is no longer in the document, or it has
                no markers in this body to nest a reply inside. U+0002 and the
                other characters XML 1.0 excludes are dropped, as
                :meth:`content`'s setter drops them.
        """
        text = drop_unwritable(text)  # CR-013 section 11, E3
        with recording(self.body, "reply_to_comment") as change:
            self._require_present()
            parts = self._writable()
            identifier = next_comment_id(parts.comments, self.body.container)
            _place_after(self.body, self.id, identifier)
            reply = _add_comment(parts, self.body, identifier, text, self)
            self._replies.append(reply)
            reply._parent = self
            change.text(after=text)
            change.touched(self.anchor_address)
            self._report(change)
            return reply

    def delete(self) -> None:
        """Remove the comment **and its replies**: markers, runs and every entry.

        The thread is read from ``w15:paraIdParent`` rather than from
        :attr:`replies`, so deleting a comment read back from the document
        removes its replies whether or not this view was told about them.
        """
        with recording(self.body, "delete_comment") as change:
            change.text(before=self.content, after="")
            change.touched(self.anchor_address)
            parts = self._writable()
            for comment in self._thread():
                comment._remove_self(parts)
            if self._parent is not None:
                self._parent._replies = [c for c in self._parent._replies if c is not self]
                self._parent = None
            self._report(change)

    # -- internals ---------------------------------------------------------

    def add_reply(self, reply: Comment) -> None:
        """Record a reply on this view. Used by :func:`comments_of` while threading."""
        reply._parent = self
        self._replies.append(reply)

    def _comment_ex(self) -> Any:
        """This comment's ``w15:commentEx`` entry, or None."""
        para_id = self.para_id
        if not para_id:
            return None
        for entry in getattr(self.parts.comments_ex, "comment_ex", None) or ():
            if (getattr(entry, "para_id", None) or "").upper() == para_id.upper():
                return entry
        return None

    def _writable(self) -> CommentParts:
        """The parts a write needs, creating the four the document may lack."""
        parts = create_comment_parts(self.parts.main)
        _note_added(self.body, parts)
        self.parts = parts
        return parts

    def _report(self, change: Any) -> None:
        """Name every comment part the call touched in ``parts_touched``."""
        _record_parts(change, self.parts)

    def _require_present(self) -> None:
        comments = getattr(self.parts.comments, "comment", None) or ()
        if not any(item is self.element for item in comments):
            raise ContentError(
                f"comment {self.id} is no longer in this document",
                code="comment.gone",
                hint="call get_comments() again; a deleted comment cannot be replied to",
            )

    def _thread(self) -> list[Comment]:
        """This comment and, depth first, every comment that replies to it."""
        by_parent: dict[str, list[Any]] = {}
        for entry in getattr(self.parts.comments_ex, "comment_ex", None) or ():
            parent = getattr(entry, "para_id_parent", None)
            if parent:
                by_parent.setdefault(parent.upper(), []).append(entry)
        by_para_id = {}
        for comment in getattr(self.parts.comments, "comment", None) or ():
            view = Comment(comment, self.parts, self.body)
            para_id = view.para_id
            if para_id:
                by_para_id[para_id.upper()] = view

        out = [self]
        queue = [self]
        while queue:
            current = queue.pop(0)
            para_id = current.para_id
            if not para_id:
                continue
            for entry in by_parent.get(para_id.upper(), ()):
                child = by_para_id.get((getattr(entry, "para_id", None) or "").upper())
                if child is not None and not any(c.id == child.id for c in out):
                    out.append(child)
                    queue.append(child)
        return out

    def _remove_self(self, parts: CommentParts) -> None:
        """Remove one comment: its markers, its ``w:comment`` and its side entries."""
        _remove_markers(self.body.container, self.id)
        comments = getattr(parts.comments, "comment", None)
        if comments is not None:
            for index, item in enumerate(comments):
                if item is self.element:
                    del comments[index]
                    break
        para_id = self.para_id
        if not para_id:
            return
        entries = getattr(parts.comments_ex, "comment_ex", None)
        if entries is not None:
            for index, entry in enumerate(entries):
                if (getattr(entry, "para_id", None) or "").upper() == para_id.upper():
                    del entries[index]
                    break
        durable_id = remove_comment_id(parts.ids_part, para_id)
        if durable_id:
            remove_extensible(parts.extensible_part, durable_id)

    def _ensure_para_id(self) -> str:
        """The first paragraph's ``w14:paraId``, assigned when there is none."""
        existing = self.para_id
        if existing:
            return existing
        content = getattr(self.element, "content", None)
        if content is None:
            content = ChildList([], owner=self.element, field="content")
            self.element.content = content
        first = next((item for item in content if isinstance(item, P)), None)
        if first is None:
            first = _comment_paragraph("", True)
            content.insert(0, first)
            link_parents(first)
            first.parent = self.element
        taken = _para_ids_taken(self.parts, self.body)
        first.para_id = _hex_id(self.body.package, taken)
        self._body_view = None
        return first.para_id


# ---------------------------------------------------------------------------
# the verbs
# ---------------------------------------------------------------------------


def _record_parts(change: Any, parts: CommentParts) -> None:
    """Name every comment part in the report's ``parts_touched``."""
    names = getattr(change, "parts", None)
    if names is None:
        return
    for name in parts.part_names:
        if name not in names:
            names.append(name)


def _note_added(body: Body, parts: CommentParts) -> None:
    """Tell a dry run about the parts this call created (CR-003 section 14.4)."""
    from docx4j_py.model.content.picture import note_added_part

    package = getattr(body, "package", None)
    for part, relationship, source, added_content_type in parts.added:
        note_added_part(
            package, part, relationship, source, added_content_type=added_content_type
        )


def _main_part_of(body: Body) -> Any:
    """The document part the comments belong to: the body's own, or the package's main one."""
    from docx4j_py.openpackaging.parts.wml import MainDocumentPart

    part = getattr(body, "part", None)
    wrapped = getattr(part, "_wrapped", part)
    if isinstance(wrapped, MainDocumentPart):
        return part
    package = getattr(body, "package", None)
    main = getattr(package, "main_document_part", None)
    if main is None:
        raise ContentError(
            "comments need a main document part, and this body has none",
            code="comment.no_main_part",
            hint="reach the body through a WordprocessingMLPackage",
        )
    return main


def _add_person(parts: CommentParts, author: Author) -> None:
    """Record the author in ``w:people``, as Word does; the email is the presence info."""
    from docx4j_py.w15 import el as w15el

    people = parts.people
    if people is None:
        return
    for person in getattr(people, "person", None) or ():
        if (getattr(person, "author", None) or "") == author.name:
            return
    # w15:contact is declared required and Word omits it, so it is left out
    person = w15el.person(
        author=author.name,
        presence_info=w15el.presenceInfo(
            provider_id="None", user_id=author.email or author.name
        ),
    )
    _append(people, "person", person)


def _add_comment(
    parts: CommentParts,
    body: Body,
    identifier: int,
    text: str,
    parent: Comment | None,
    marks: Any = None,
) -> Comment:
    """The ``w:comment`` and every side entry of a new comment; the markers are written.

    `marks` is :func:`_comment_paragraphs`'.
    """
    from docx4j_py.model.content.tracking import now_to_the_minute, xml_date
    from docx4j_py.w15 import el as w15el
    from docx4j_py.wml import el

    author = author_of(body.package)
    # one instant for w:date (the local wall clock, as Word writes it) and
    # w16cex:dateUtc (UTC), to the minute as Word stamps it (CR-010 section 12.9)
    now = now_to_the_minute()
    paragraphs = _comment_paragraphs(text, marks=marks)
    taken = _para_ids_taken(parts, body)
    para_id = _hex_id(body.package, taken)
    paragraphs[0].para_id = para_id

    comment = el.comment(
        id=identifier,
        author=author.name,
        initials=initials_of(author),
        date=xml_date(now),
        content=ChildList(paragraphs),
    )
    _append(parts.comments, "comment", comment)

    entry = w15el.commentEx(para_id=para_id, done="0")
    if parent is not None:
        entry.para_id_parent = parent._ensure_para_id()
    if parts.comments_ex is not None:
        _append(parts.comments_ex, "comment_ex", entry)

    if parts.ids_part is not None:
        durable_id = _hex_id(
            body.package, durable_ids_in_use(parts.ids_part), derive_from=taken
        )
        set_comment_id(parts.ids_part, para_id, durable_id)
        set_extensible(parts.extensible_part, durable_id, now)
        parts.dates_utc_cache = None

    _add_person(parts, author)
    return Comment(comment, parts, body)


def insert_comment_into(span: Range, text: str, *, marks: Any = None) -> Comment:
    """``Range.insert_comment`` and ``Paragraph.insert_comment``. See their docstrings.

    The comment's text is a paragraph per line, ``\\r`` and ``\\n`` both
    marks (CR-013 section 4.1), each of Office JS's characters its element
    (Q2); U+0002 and the other characters XML 1.0 excludes are dropped, as
    Word drops them (CR-013 section 11, E3; core-ts check 55 cases 38 to 40).
    `marks`, a compiled pattern, splits the lines otherwise (the python-docx
    facade's ``\\n``, after its own refusal of those characters).
    """
    text = drop_unwritable(text)
    body = span.paragraph.parent_body
    if span.spans_paragraphs:
        # across paragraphs (CR-013 Phase C) the markers go in the first and the last paragraph,
        # as Word's (core-ts check 52 cases 79 to 85): those two refuse a kept branch's text
        from docx4j_py.model.content.reports import refuse_read_only

        for paragraph in (span.paragraph, span.end_paragraph):
            refuse_read_only(paragraph.parent_body, "insert_comment")
        pieces = span._pieces()
        for piece in (pieces[0], pieces[-1]):
            piece.paragraph.refuse_branch_text(piece.start, piece.end, "insert_comment")
    with recording(body, "insert_comment") as change:
        if not span.spans_paragraphs:
            span.refuse_branch_text("insert_comment")
            span.require_one_holder("insert_comment")
        main = _main_part_of(body)
        parts = create_comment_parts(main)
        _note_added(body, parts)
        for name in ensure_comment_styles(body.package):
            names = getattr(change, "parts", None)
            if names is not None and name not in names:
                names.append(name)
        identifier = next_comment_id(parts.comments, body.container)
        _place_around(span, identifier)
        comment = _add_comment(parts, body, identifier, text, None, marks)
        change.touched(span.paragraph, span.end_paragraph)
        change.text(after=text)
        _record_parts(change, parts)
        return comment


def comments_of(scope: Any) -> list[Comment]:
    """``get_comments()`` on a body, a paragraph, a range or a content control.

    The comments anchored in the scope, in document order, with replies nested
    under their parent: a reply is in the list only when its parent is not.
    """
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

    if isinstance(scope, Range):
        body = scope.paragraph.parent_body
        ids = _ids_in_range(scope)
    elif isinstance(scope, Paragraph):
        body = scope.parent_body
        ids = _unique(
            m.id
            for m in markers_of_paragraph(scope.element, scope.container, _resolver_of(scope.parent_body))
        )
    elif isinstance(scope, Body):
        body = scope
        ids = comment_ids_of(scope.container)
    else:  # a ContentControl, or anything else with an element in a body
        body = scope.parent_body
        ids = comment_ids_of(scope.element)

    parts = comment_parts_of(_main_part_of(body))
    if parts is None:
        return []
    views: dict[int, Comment] = {}
    for comment in getattr(parts.comments, "comment", None) or ():
        view = Comment(comment, parts, body)
        if view.id >= 0 and view.id not in views:
            views[view.id] = view
    by_para_id = {v.para_id.upper(): v for v in views.values() if v.para_id}
    for view in views.values():
        entry = view._comment_ex()
        parent_para_id = getattr(entry, "para_id_parent", None) if entry is not None else None
        parent = by_para_id.get(parent_para_id.upper()) if parent_para_id else None
        if parent is not None and parent is not view:
            parent.add_reply(view)

    in_scope = set(ids)
    out: list[Comment] = []
    for identifier in ids:
        view = views.get(identifier)
        if view is None:
            continue
        if view.parent is not None and view.parent.id in in_scope:
            continue
        out.append(view)
    return out


def _ids_in_range(span: Range) -> list[int]:
    """The comments a span touches: those whose extent over the story overlaps ``[start, end]``, both edges inclusive.

    CR-013 Phase B (section 2.5 item 1): a comment's extent runs from its
    first marker to its last over the story, each as ``(paragraph, offset)``
    in story order, a block-level marker at the next paragraph's start (a
    start) or the previous one's end (an end). Until then
    only the span's own paragraph's markers were read, so a range on a
    comment's last paragraph, which holds only its end marker, listed
    nothing. Within one paragraph's markers the answer is as it was.
    """
    from docx4j_py.model.content.story import story_layout, story_of

    items, _owner = story_of(span.paragraph.element)
    order_of: dict[int, int] = {}
    layout = story_layout(items, lambda _element, _container: "")
    for index, place in enumerate(layout.places):
        order_of[id(place.element)] = index
    first = order_of.get(id(span.paragraph.element))
    last = order_of.get(id(span.end_paragraph.element))
    if first is None or last is None:
        return []
    low, high = (first, span.start), (last, span.end)
    resolver = _resolver_of(span.paragraph.parent_body)
    spans: dict[int, list[tuple[int, float]]] = {}
    order: list[int] = []
    for marker in markers_of(items, resolver):
        if marker.paragraph is not None:
            index = order_of.get(id(marker.paragraph))
            at: tuple[int, float] | None = None if index is None else (index, marker.offset or 0)
        else:
            at = _block_marker_at(marker, order_of)
        if at is None:
            continue
        found = spans.get(marker.id)
        if found is None:
            spans[marker.id] = [at, at]
            order.append(marker.id)
        else:
            found[0] = min(found[0], at)
            found[1] = max(found[1], at)
    return [i for i in order if spans[i][0] <= high and spans[i][1] >= low]


def _block_marker_at(marker: CommentMarker, order_of: dict[int, int]) -> tuple[int, float] | None:
    """A block-level marker's place in story order: a start at the next paragraph's start, an end at the previous one's end."""
    from docx4j_py.model.content.story import first_and_last_paragraph

    at = _index_of(marker.owner, marker.item)
    if at < 0:
        return None
    if marker.kind == "start":
        for index in range(at + 1, len(marker.owner)):
            first, _last = first_and_last_paragraph(marker.owner, index)
            if first is not None and id(first.element) in order_of:
                return order_of[id(first.element)], 0
        return None
    for index in range(at - 1, -1, -1):
        _first, last = first_and_last_paragraph(marker.owner, index)
        if last is not None and id(last.element) in order_of:
            return order_of[id(last.element)], float("inf")
    return None


def _unique(ids: Any) -> list[int]:
    seen: set[int] = set()
    out: list[int] = []
    for identifier in ids:
        if identifier not in seen:
            seen.add(identifier)
            out.append(identifier)
    return out
