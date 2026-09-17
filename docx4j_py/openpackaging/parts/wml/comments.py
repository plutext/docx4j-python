"""The five comment parts of a document, loaded and created. CR-003 Phase G.

The part half of CR-003 section 3.9; the views and the verbs are
:mod:`docx4j_py.model.content.comments`. The split is section 5's
("half view, half part") and section 5's import rule: **the parts layer never
imports the content API**, so nothing here knows what a ``Comment`` is, and
nothing here touches the styles part --- that is
:func:`docx4j_py.model.content.comments.ensure_comment_styles`, because the
rule for a built-in style a document lacks lives in
:mod:`docx4j_py.model.content.styles` (CR-003 section 14.9).

The five parts, and which of them a *read* pays for:

``/word/comments.xml`` (``w:comments``)
    the comments themselves. Unmarshalled by a read.
``/word/commentsExtended.xml`` (``w15:commentsEx``)
    the thread links (``w15:paraIdParent``) and the resolved flags
    (``w15:done``), keyed by the ``w14:paraId`` of a comment's first paragraph.
    Unmarshalled by a read.
``/word/people.xml`` (``w15:people``)
    the authors, whose ``w15:presenceInfo/@w15:userId`` carries the email.
    Unmarshalled by a read.
``/word/commentsIds.xml`` (``w16cid:commentsIds``)
    the durable ids. **Only touched by a write**, and held as an lxml tree:
    ``w16cid`` is outside this build's schema closure (CR-001 section 13.5), so
    the registry gives it a
    :class:`~docx4j_py.openpackaging.parts.default_xml_part.DefaultXmlPart`.
``/word/commentsExtensible.xml`` (``w16cex:commentsExtensible``)
    a UTC date per durable id, likewise an lxml tree. **Never created**, only
    kept in step when the document already has one (CR-003 section 4).

So a document whose comments are never read keeps all five byte for byte, and
one whose comments are read but not edited keeps the last two.
"""

from __future__ import annotations

import datetime
from typing import TYPE_CHECKING, Any

from lxml import etree

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.parts.default_xml_part import DefaultXmlPart
from docx4j_py.openpackaging.parts.namespaces import Namespaces

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.parts.part import Part

__all__ = [
    "COMMENTS_EXTENSIBLE_NS",
    "COMMENTS_IDS_NS",
    "CommentParts",
    "comment_parts_of",
    "create_comment_parts",
    "durable_id_for",
    "durable_ids_in_use",
    "remove_comment_id",
    "remove_extensible",
    "set_comment_id",
    "set_extensible",
]

#: ``w16cid``, the durable-id namespace, and ``w16cex``, the extensible one.
COMMENTS_IDS_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
COMMENTS_EXTENSIBLE_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex"

#: What a newly created ``w:comments`` declares as ignorable. Its paragraphs
#: carry ``w14:paraId`` and its side parts are keyed by it, which is why Word
#: names ``w14`` and ``w15`` here; both are in docx4j's prefix table, so
#: :attr:`~docx4j_py.openpackaging.parts.xml_part.XmlPart.xml` can declare them.
NEW_COMMENTS_IGNORABLE = "w14 w15"

_MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"


def _w16cid(local: str) -> str:
    return f"{{{COMMENTS_IDS_NS}}}{local}"


def _w16cex(local: str) -> str:
    return f"{{{COMMENTS_EXTENSIBLE_NS}}}{local}"


class CommentParts:
    """The comment parts of one document, with what a verb needs to report.

    A plain holder, made by :func:`comment_parts_of` (a read: three parts) or
    :func:`create_comment_parts` (a write: four, the fifth only if present).
    """

    __slots__ = (
        "added",
        "comments",
        "comments_ex",
        "comments_part",
        "extensible_part",
        "ids_part",
        "main",
        "people",
    )

    def __init__(self, main: Any, comments_part: Any, comments: Any) -> None:
        """Hold the main document part, the comments part and its ``w:comments``."""
        #: The ``MainDocumentPart`` the comments belong to.
        self.main = main
        #: ``/word/comments.xml``.
        self.comments_part = comments_part
        #: The ``w:comments``.
        self.comments = comments
        #: The ``w15:commentsEx``, or None when the document has no such part.
        self.comments_ex: Any = None
        #: The ``w15:people``, or None.
        self.people: Any = None
        #: ``/word/commentsIds.xml`` as an lxml part, or None.
        self.ids_part: Any = None
        #: ``/word/commentsExtensible.xml`` as an lxml part, or None.
        self.extensible_part: Any = None
        #: ``(part, relationship, source, added_content_type)`` per part created
        #: here, so that a caller can report them and a dry run can un-add them
        #: (CR-003 section 14.4).
        self.added: list[tuple[Any, Any, Any, bool]] = []

    @property
    def part_names(self) -> list[str]:
        """The names of every part a write touches, for ``parts_touched``."""
        out: list[str] = [str(self.comments_part.part_name)]
        for part in (
            self.main.comments_extended_part,
            self.ids_part,
            self.main.people_part,
            self.extensible_part,
        ):
            if part is not None:
                out.append(str(part.part_name))
        return out

    def __repr__(self) -> str:
        """``<CommentParts /word/comments.xml, 3 comments>``."""
        count = len(getattr(self.comments, "comment", None) or ())
        return f"<CommentParts {self.comments_part.part_name}, {count} comments>"


# ---------------------------------------------------------------------------
# loading and creating
# ---------------------------------------------------------------------------


def comment_parts_of(main: Any) -> CommentParts | None:
    """The comment parts a read needs, or None when the document has none.

    Unmarshals ``w:comments``, ``w15:commentsEx`` and ``w15:people`` --- the
    three a :class:`~docx4j_py.model.content.comments.Comment` view reads ---
    and leaves ``w16cid`` and ``w16cex`` alone (CR-003 section 3.9: "reads
    unmarshal three parts, writes five").
    """
    comments_part = getattr(main, "comments_part", None)
    if comments_part is None:
        return None
    parts = CommentParts(main, comments_part, comments_part.contents)
    extended = getattr(main, "comments_extended_part", None)
    if extended is not None:
        parts.comments_ex = extended.contents
    people = getattr(main, "people_part", None)
    if people is not None:
        parts.people = people.contents
    parts.ids_part = getattr(main, "comments_ids_part", None)
    parts.extensible_part = getattr(main, "comments_extensible_part", None)
    return parts


def create_comment_parts(main: Any) -> CommentParts:
    """As :func:`comment_parts_of`, creating any of the first four that is absent.

    Each new part gets its relationship and its content type through
    ``add_target_part``, exactly as
    :func:`docx4j_py.model.content.picture.add_image` adds an image part, and is
    recorded in :attr:`CommentParts.added` so that the caller can report it and
    a dry run can take it away again (CR-003 section 14.4).

    ``w16cex:commentsExtensible`` is **not** created (CR-003 section 4); when the
    document has one it is kept in step.
    """
    from docx4j_py.openpackaging.parts.wml import (
        CommentsExtendedPart,
        CommentsPart,
        PeoplePart,
    )

    added: list[tuple[Any, Any, Any, bool]] = []

    comments_part = getattr(main, "comments_part", None)
    if comments_part is None:
        comments_part = CommentsPart()
        comments_part.set_contents(_empty_comments())
        _add(main, comments_part, added)

    if getattr(main, "comments_extended_part", None) is None:
        part = CommentsExtendedPart()
        part.set_contents(_empty_comments_ex())
        _add(main, part, added)

    if getattr(main, "comments_ids_part", None) is None:
        part = DefaultXmlPart(
            "/word/commentsIds.xml",
            ContentTypes.WORDPROCESSINGML_COMMENTS_IDS,
            Namespaces.COMMENTS_IDS,
        )
        part.set_tree(etree.Element(_w16cid("commentsIds"), nsmap={"w16cid": COMMENTS_IDS_NS}))
        _add(main, part, added)

    if getattr(main, "people_part", None) is None:
        part = PeoplePart()
        part.set_contents(_empty_people())
        _add(main, part, added)

    parts = comment_parts_of(main)
    assert parts is not None  # the comments part exists by now
    parts.added = added
    return parts


def _add(main: Any, part: Part, added: list[tuple[Any, Any, Any, bool]]) -> None:
    """Relate a new part from the main document part and note what that added."""
    package = main.package
    had_override = package.content_type_manager.get_override_content_type(part.part_name)
    rel = main.add_target_part(part)
    added.append((part, rel, main, had_override is None))


def _empty_comments() -> Any:
    """``<w:comments mc:Ignorable="w14 w15"/>``, as Word writes a fresh one."""
    from docx4j_py.wml import Comments

    comments = Comments()
    comments.other_attributes[_MC_IGNORABLE] = NEW_COMMENTS_IGNORABLE
    return comments


def _empty_comments_ex() -> Any:
    """``<w15:commentsEx/>``. No ``mc:Ignorable``: ``w15`` is the root's own namespace."""
    from docx4j_py.w15 import CommentsEx

    return CommentsEx()


def _empty_people() -> Any:
    """``<w15:people/>``."""
    from docx4j_py.w15 import People

    return People()


# ---------------------------------------------------------------------------
# w16cid:commentsIds, as lxml (CR-001 section 13.5: outside the schema closure)
# ---------------------------------------------------------------------------


def _entries(part: Any, name: str) -> list[Any]:
    """The child elements of an lxml comment part, or an empty list."""
    if part is None:
        return []
    return list(part.tree.findall(name))


def _touch(part: Any) -> None:
    """Adopt an lxml part's tree after editing it in place (CR-003 Phase E).

    :class:`~docx4j_py.openpackaging.parts.default_xml_part.DefaultXmlPart`
    writes a part that was only *parsed* back byte for byte, so anything that
    edits :attr:`tree` has to say it did.
    """
    mark = getattr(part, "mark_modified", None)
    if mark is not None:
        mark()


def durable_ids_in_use(part: Any) -> set[str]:
    """Every ``w16cid:durableId`` the part already carries, upper-cased."""
    out: set[str] = set()
    for entry in _entries(part, _w16cid("commentId")):
        value = entry.get(_w16cid("durableId"))
        if value:
            out.add(value.upper())
    return out


def durable_id_for(part: Any, para_id: str) -> str | None:
    """The durable id recorded for a comment's paragraph id, or None."""
    for entry in _entries(part, _w16cid("commentId")):
        if (entry.get(_w16cid("paraId")) or "").upper() == para_id.upper():
            return entry.get(_w16cid("durableId"))
    return None


def set_comment_id(part: Any, para_id: str, durable_id: str) -> None:
    """Add or update the ``w16cid:commentId`` entry of a paragraph id."""
    if part is None:
        return
    for entry in _entries(part, _w16cid("commentId")):
        if (entry.get(_w16cid("paraId")) or "").upper() == para_id.upper():
            entry.set(_w16cid("durableId"), durable_id)
            _touch(part)
            return
    entry = etree.SubElement(part.tree, _w16cid("commentId"))
    entry.set(_w16cid("paraId"), para_id)
    entry.set(_w16cid("durableId"), durable_id)
    _touch(part)


def remove_comment_id(part: Any, para_id: str) -> str | None:
    """Remove the entry of a paragraph id; returns the durable id it carried."""
    if part is None:
        return None
    for entry in _entries(part, _w16cid("commentId")):
        if (entry.get(_w16cid("paraId")) or "").upper() == para_id.upper():
            durable_id = entry.get(_w16cid("durableId"))
            part.tree.remove(entry)
            _touch(part)
            return durable_id
    return None


# ---------------------------------------------------------------------------
# w16cex:commentsExtensible, as lxml. Never created, only kept in step.
# ---------------------------------------------------------------------------


def set_extensible(part: Any, durable_id: str, when: datetime.datetime) -> None:
    """Add or update the ``w16cex:commentExtensible`` entry of a durable id."""
    if part is None:
        return
    stamp = when.astimezone(datetime.UTC).replace(microsecond=0, tzinfo=None).isoformat() + "Z"
    for entry in _entries(part, _w16cex("commentExtensible")):
        if (entry.get(_w16cex("durableId")) or "").upper() == durable_id.upper():
            entry.set(_w16cex("dateUtc"), stamp)
            _touch(part)
            return
    entry = etree.SubElement(part.tree, _w16cex("commentExtensible"))
    entry.set(_w16cex("durableId"), durable_id)
    entry.set(_w16cex("dateUtc"), stamp)
    _touch(part)


def remove_extensible(part: Any, durable_id: str) -> None:
    """Remove the entry of a durable id, if there is one."""
    if part is None:
        return
    for entry in _entries(part, _w16cex("commentExtensible")):
        if (entry.get(_w16cex("durableId")) or "").upper() == durable_id.upper():
            part.tree.remove(entry)
            _touch(part)
            return
