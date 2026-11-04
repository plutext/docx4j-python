"""A content API deletion and a story part that cannot be read: CR-031 Phase A3, section 4.6b.

A delete, a replace or a block removed records what its content referenced and
settles it when the write ends (:func:`~docx4j_py.model.content.span_edits.settle_removals`,
through the revisions engine's ``Removals``), after its edit, so it cannot
refuse first as a resolution does (``tests/revisions/test_unreadable_story_part.py``).
At the settle, a story part that cannot be read is neither consulted nor
changed: a comment or a note it might reference is kept, its markers with it,
and the part is written byte for byte; an unreadable comments or notes part
loses no entry. At worst a comment entry stays that a readable part would have
let go.

**The fixture** is ``samples/Headers.docx`` made in code (:func:`built`): in the
body, comment 3's range and reference, then a paragraph holding footnote 1's
reference; ``word/header1.xml`` referencing comment 3 too, where asked; and the
part a test cuts short by its last three bytes, after the reference.
"""

from __future__ import annotations

import io
import zipfile
from collections.abc import Callable
from typing import Any

import pytest
from conftest import SAMPLES

from docx4j_py import load

SOURCE = SAMPLES / "Headers.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

HEADER = "word/header1.xml"
FOOTNOTES = "word/footnotes.xml"
COMMENTS = "word/comments.xml"

#: At the body's start: comment 3's range and reference, then footnote 1's reference, each in a paragraph of its own.
_BODY = (
    b'<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Anchor</w:t></w:r><w:commentRangeEnd w:id="3"/>'
    b'<w:r><w:commentReference w:id="3"/></w:r></w:p>'
    b'<w:p><w:r><w:t>Noted</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
)
_HEADER_REFERENCE = b'<w:p><w:r><w:t>Header reference</w:t></w:r><w:r><w:commentReference w:id="3"/></w:r></w:p>'
_NOTE = b'<w:footnote w:id="1"><w:p><w:r><w:t>The note</w:t></w:r></w:p></w:footnote>'
_COMMENTS = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:comments xmlns:w="{W}">'
    '<w:comment w:id="3" w:author="A" w:date="2026-11-04T00:00:00Z"><w:p><w:r><w:t>The comment</w:t></w:r></w:p>'
    "</w:comment></w:comments>"
).encode()
_RELATIONSHIP = (
    b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"'
    b' Target="comments.xml"/>'
)
_OVERRIDE = (
    b'<Override PartName="/word/comments.xml" '
    b'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
)

_BUILT: dict[tuple[str | None, bool], bytes] = {}


def built(cut: str | None = None, *, header_reference: bool = True) -> bytes:
    """The fixture, the entry `cut` (when given) short by its last three bytes; the header's reference where asked."""
    key = (cut, header_reference)
    if key not in _BUILT:
        edits: dict[str, Callable[[bytes], bytes]] = {
            "word/document.xml": lambda data: data.replace(b"<w:body>", b"<w:body>" + _BODY, 1),
            FOOTNOTES: lambda data: data.replace(b"</w:footnotes>", _NOTE + b"</w:footnotes>"),
            "word/_rels/document.xml.rels": lambda data: data.replace(
                b"</Relationships>", _RELATIONSHIP + b"</Relationships>"
            ),
            "[Content_Types].xml": lambda data: data.replace(b"</Types>", _OVERRIDE + b"</Types>"),
        }
        if header_reference:
            edits[HEADER] = lambda data: data.replace(b"</w:hdr>", _HEADER_REFERENCE + b"</w:hdr>")
        out = io.BytesIO()
        with zipfile.ZipFile(SOURCE) as source, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
            entries = [(info.filename, source.read(info)) for info in source.infolist()]
            for name, data in [*entries, (COMMENTS, _COMMENTS)]:
                edited = edits[name](data) if name in edits else data
                assert name not in edits or edited != data, name
                target.writestr(name, edited[:-3] if name == cut else edited)
        _BUILT[key] = out.getvalue()
    return _BUILT[key]


def entry(data: bytes, name: str) -> bytes | None:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read(name) if name in archive.namelist() else None


def _comment_ids(data: bytes) -> list[int]:
    return [c.id for c in load(data).main_document_part.comments_part.contents.comment]


#: The deletes that take the body's reference to comment 3: the span route (``queue_removed``, the paragraph joined
#: to the next) and the paragraph's own splice (``queue_comment_removal``).
DELETES: dict[str, Callable[[Any], None]] = {
    "the paragraph's Whole range": lambda pkg: pkg.body.paragraphs[0].get_range("Whole").delete(),
    "the paragraph's Content range": lambda pkg: pkg.body.paragraphs[0].get_range("Content").delete(),
    "a replace of the paragraph's Whole range": lambda pkg: (
        pkg.body.paragraphs[0].get_range("Whole").insert_text("Replaced", location="Replace")
    ),
}


@pytest.mark.parametrize("delete", list(DELETES))
def test_a_delete_keeps_a_comment_a_cut_header_might_reference(delete: str) -> None:
    data = built(HEADER)
    pkg = load(data)
    header = pkg.get_part("/" + HEADER)
    DELETES[delete](pkg)
    assert "commentReference" not in pkg.main_document_part.get_xml().split("Noted")[0]
    # the settle neither read the header nor changed it
    assert not header.is_unmarshalled and header.unreadable is None
    saved = pkg.save()
    assert entry(saved, HEADER) == entry(data, HEADER)
    assert entry(saved, COMMENTS) == entry(data, COMMENTS)
    assert _comment_ids(saved) == [3]


@pytest.mark.parametrize("delete", list(DELETES))
def test_a_delete_keeps_the_comment_even_where_the_cut_header_holds_no_reference(delete: str) -> None:
    """An unknown part might reference any comment: nothing is lost, and the readable markers stay with it."""
    data = built(HEADER, header_reference=False)
    pkg = load(data)
    DELETES[delete](pkg)
    saved = pkg.save()
    assert _comment_ids(saved) == [3]
    assert entry(saved, HEADER) == entry(data, HEADER)


@pytest.mark.parametrize(
    ("header_reference", "kept"),
    [(True, [3]), (False, [])],
    ids=["referenced from the header", "referenced from nothing else"],
)
def test_with_the_header_intact_the_header_s_reference_decides(header_reference: bool, kept: list[int]) -> None:
    """The control: a readable header keeps the comment by its reference, and without one the comment goes."""
    data = built(header_reference=header_reference)
    pkg = load(data)
    DELETES["the paragraph's Whole range"](pkg)
    saved = pkg.save()
    if kept:
        assert _comment_ids(saved) == kept
    else:
        assert entry(saved, COMMENTS) is None  # with the last comment, the comment parts go (CR-010 section 12.8)
    assert entry(saved, HEADER) == entry(data, HEADER)


@pytest.mark.parametrize("cut", [None, FOOTNOTES], ids=["intact", "footnotes cut"])
def test_a_delete_keeps_a_note_in_a_cut_footnotes_part(cut: str | None) -> None:
    """Before Phase A3, the settle read the cut part and raised, after the edit; the intact part loses the note."""
    data = built(cut)
    pkg = load(data)
    pkg.body.paragraphs[1].get_range("Whole").delete()
    assert "footnoteReference" not in pkg.main_document_part.get_xml()
    saved = pkg.save()
    if cut is None:
        assert [n.id for n in load(saved).main_document_part.footnotes_part.contents.footnote] == [-1, 0]
    else:
        assert entry(saved, FOOTNOTES) == entry(data, FOOTNOTES)
        assert pkg.main_document_part.footnotes_part.unreadable is None


def test_a_delete_with_a_cut_comments_part_removes_no_entry() -> None:
    """The comments part itself cannot be read: it is not changed, and the body's markers stay with its entry."""
    data = built(COMMENTS, header_reference=False)
    pkg = load(data)
    comments = pkg.main_document_part.comments_part
    DELETES["the paragraph's Whole range"](pkg)
    assert not comments.is_unmarshalled and comments.unreadable is None
    saved = pkg.save()
    assert entry(saved, COMMENTS) == entry(data, COMMENTS)


def test_a_cut_header_keeps_a_note_too() -> None:
    """Every note is kept where any story part cannot be read, as every comment is, whatever the part's kind.

    The settle unions the references of every story part, a header's among them, as it always has; a header
    Word would give no note reference is unknown all the same. Before Phase A3 the note went.
    """
    data = built(HEADER)
    pkg = load(data)
    pkg.body.paragraphs[1].get_range("Whole").delete()
    assert "footnoteReference" not in pkg.main_document_part.get_xml()
    assert not pkg.main_document_part.footnotes_part.is_unmarshalled
    saved = pkg.save()
    assert entry(saved, FOOTNOTES) == entry(data, FOOTNOTES)
    assert [n.id for n in load(saved).main_document_part.footnotes_part.contents.footnote] == [-1, 0, 1]
