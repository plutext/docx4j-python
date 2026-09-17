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
* a comment is **not a revision**: its markers are hoisted out of the ``w:ins``
  or ``w:del`` the anchor run sits in, so accepting or rejecting that revision
  leaves the comment where it is.
* comment ids are their own space, above the highest ``w:comment/@w:id`` and the
  highest id on a marker; revision ids (Phase F) are a different counter.
"""

from __future__ import annotations

import dataclasses
import datetime
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, link_parents
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.styles import ensure_style, style_ids_of
from docx4j_py.model.content.text_model import (
    _REVISIONS,
    RUN_HOLDER_NAMES,
    block_children_of,
    item_text,
)
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

#: The four run-level revision holders, by element name. A marker inside one is
#: hoisted out of it (section 4: a comment is not a revision).
_REVISION_NAMES = frozenset(_REVISIONS)


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
    value = getattr(item, "id", None)
    return int(value) if isinstance(value, int) else None


def _visit_runs(out: list[CommentMarker], items: Any, paragraph: Any, container: Any) -> int:
    """Collect the markers of one run list; returns the text length it covered."""
    position = 0
    if not items:
        return position
    for element in items:
        name = element_name(element)
        if name == W_R:
            content = getattr(element, "content", None)
            if not isinstance(content, list):
                continue
            for child in content:
                qname = element_name(child)
                if qname == W_COMMENT_REFERENCE:
                    identifier = _marker_id(child)
                    if identifier is not None:
                        out.append(
                            CommentMarker(
                                id=identifier,
                                kind="reference",
                                item=child,
                                owner=content,
                                run=element,
                                run_owner=items,
                                paragraph=paragraph,
                                paragraph_container=container,
                                offset=position,
                            )
                        )
                    continue
                text = item_text(qname, child)
                if text:
                    position += len(text)
            continue
        if name in (W_COMMENT_RANGE_START, W_COMMENT_RANGE_END):
            identifier = _marker_id(element)
            if identifier is not None:
                out.append(
                    CommentMarker(
                        id=identifier,
                        kind="start" if name == W_COMMENT_RANGE_START else "end",
                        item=element,
                        owner=items,
                        paragraph=paragraph,
                        paragraph_container=container,
                        offset=position,
                    )
                )
            continue
        shown = _REVISIONS.get(name or "")
        if shown is not None:
            # the accepted view, which is what the offsets are in: a w:ins and a
            # w:moveTo hold text and markers, a w:del and a w:moveFrom do not
            if shown == "accepted":
                position += _visit_runs(out, run_items_of(element), paragraph, container)
        elif name in RUN_HOLDER_NAMES:
            position += _visit_runs(out, run_items_of(element), paragraph, container)
    return position


def markers_of_paragraph(element: P, container: list) -> list[CommentMarker]:
    """Every comment marker in one paragraph, with its offset in the text."""
    out: list[CommentMarker] = []
    _visit_runs(out, run_items_of(element), element, container)
    return out


def markers_of(container: Any) -> list[CommentMarker]:
    """Every comment marker under a container, in document order.

    A table, a row, a cell and a content control are all one
    :func:`~docx4j_py.model.content.text_model.block_children_of` step, so the
    walk has no special case for any of them.
    """
    out: list[CommentMarker] = []

    def visit(items: list) -> None:
        for item in items:
            if isinstance(item, P):
                _visit_runs(out, run_items_of(item), item, items)
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

    children = block_children_of(container)
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


def _comment_paragraph(text: str, first: bool) -> P:
    """A paragraph of a comment: ``CommentText``, opened by the annotation reference."""
    from docx4j_py.wml import el
    from docx4j_py.wml import r as text_run

    content: list[Any] = []
    if first:
        content.append(_annotation_ref_run())
    if text:
        content.append(text_run(text))
    paragraph = P(p_pr=el.pPr(p_style=el.pStyle(val=COMMENT_TEXT_STYLE)), content=ChildList(content))
    link_parents(paragraph)
    return paragraph


def _comment_paragraphs(text: str) -> list[P]:
    """One paragraph per line, the first opened by the annotation reference run."""
    return [_comment_paragraph(line, index == 0) for index, line in enumerate(text.split("\n"))]


def _append(owner: Any, field: str, item: Any) -> None:
    """Append to a model list and wire the parents, as every verb here does."""
    items = getattr(owner, field, None)
    if items is None:
        items = ChildList([], owner=owner)
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
    """Where a marker for a segment's run goes: beside the run, outside a revision.

    Section 4: a comment is **not** a revision, so its markers are hoisted out of
    the ``w:ins`` or ``w:del`` the anchor run sits in. Accepting or rejecting
    that revision then leaves the comment where it is; the cost is that a comment
    on part of an insertion widens to the whole of it, which is what Word shows
    once the insertion is accepted anyway.
    """
    holder = getattr(segment.run, "parent", None)
    if element_name(holder) in _REVISION_NAMES:
        outer = getattr(holder, "parent", None)
        items = run_items_of(outer) if outer is not None else None
        if items is not None:
            index = _index_of(items, holder)
            if index >= 0:
                return items, index + (1 if after else 0), outer
    owner = segment.run_owner
    parent = getattr(segment.run, "parent", None) or paragraph.element
    return owner, segment.run_index + (1 if after else 0), parent


def _place_around(span: Range, identifier: int) -> None:
    """Write the markers of a new comment around a span, splitting at its boundaries."""
    paragraph = span.paragraph
    start = paragraph.split_at(span.start, prefer="back")
    end = paragraph.split_at(span.end, prefer="forward") if span.end > span.start else start
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

    owner, index, parent = _marker_site(paragraph, covered[-1], after=True)
    _insert_at(owner, index, [el.commentRangeEnd(id=identifier), reference], parent)
    owner, index, parent = _marker_site(paragraph, covered[0], after=False)
    _insert_at(owner, index, [el.commentRangeStart(id=identifier)], parent)


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
        value = getattr(self.element, "id", None)
        return int(value) if isinstance(value, int) else -1

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
        """``w:date`` as a ``datetime``, or None: a comment may carry none."""
        value = getattr(self.element, "date", None)
        if value is None:
            return None
        to_datetime = getattr(value, "to_datetime", None)
        return to_datetime() if to_datetime is not None else None

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
        """The comment's text, a paragraph per line. Setting it replaces them."""
        return self.comment_body.text

    @content.setter
    def content(self, text: str) -> None:
        with recording(self.body, "set_comment_content") as change:
            change.text(before=self.content, after=text)
            para_id = self.para_id
            paragraphs = _comment_paragraphs(text)
            if para_id:
                paragraphs[0].para_id = para_id
            self.element.content = ChildList(paragraphs, owner=self.element)
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
        value = getattr(entry, "done", None) if entry is not None else None
        value = getattr(value, "value", value)
        return str(value) in ("1", "true", "on")

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

    def get_range(self) -> list[Range]:
        """The commented span: a :class:`Range` **per paragraph** it touches.

        Office JS returns one ``Range``; CR-003 section 4 says a list, because a
        comment may span paragraphs and this API's ``Range`` is within one.
        Empty for a comment whose markers sit at block level, and for one whose
        markers are in another body.
        """
        from docx4j_py.model.content.range import Range

        markers = [
            m for m in markers_of(self.body.container) if m.id == self.id and m.paragraph is not None
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
                no markers in this body to nest a reply inside.
        """
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
            content = ChildList([], owner=self.element)
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
) -> Comment:
    """The ``w:comment`` and every side entry of a new comment; the markers are written."""
    from docx4j_xsdata.models.datatype import XmlDateTime

    from docx4j_py.w15 import el as w15el
    from docx4j_py.wml import el

    author = author_of(body.package)
    now = datetime.datetime.now(datetime.UTC).replace(microsecond=0)
    paragraphs = _comment_paragraphs(text)
    taken = _para_ids_taken(parts, body)
    para_id = _hex_id(body.package, taken)
    paragraphs[0].para_id = para_id

    comment = el.comment(
        id=identifier,
        author=author.name,
        initials=initials_of(author),
        date=XmlDateTime.from_datetime(now),
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

    _add_person(parts, author)
    return Comment(comment, parts, body)


def insert_comment_into(span: Range, text: str) -> Comment:
    """``Range.insert_comment`` and ``Paragraph.insert_comment``. See their docstrings."""
    body = span.paragraph.parent_body
    with recording(body, "insert_comment") as change:
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
        comment = _add_comment(parts, body, identifier, text, None)
        change.touched(span.paragraph)
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
        ids = _unique(m.id for m in markers_of_paragraph(scope.element, scope.container))
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
    """The comments a span touches: those whose markers overlap ``[start, end]``."""
    spans: dict[int, list[int]] = {}
    order: list[int] = []
    for marker in markers_of_paragraph(span.paragraph.element, span.paragraph.container):
        at = marker.offset or 0
        found = spans.get(marker.id)
        if found is None:
            spans[marker.id] = [at, at]
            order.append(marker.id)
        else:
            found[0] = min(found[0], at)
            found[1] = max(found[1], at)
    return [i for i in order if spans[i][0] <= span.end and spans[i][1] >= span.start]


def _unique(ids: Any) -> list[int]:
    seen: set[int] = set()
    out: list[int] = []
    for identifier in ids:
        if identifier not in seen:
            seen.add(identifier)
            out.append(identifier)
    return out
