"""A resolution and a story part it cannot read: CR-031 Phase A3, section 4.6b.

``Removals`` settles what a resolution removed against what still stands in
every story, and scans a story part nobody has read in its bytes, with lxml
(CR-010 section 4). Before Phase A3 it read one whose XML is not well formed as
holding nothing (CR-031 section 11.4 row 8), so ``accept_all`` removed a
comment that a cut header still referenced, and saved the header as read. Now:

- **a resolution of every story refuses first.** ``resolve_all`` (``accept_all``,
  ``reject_all``, ``Body.accept_all`` and ``reject_all``) and a single change's
  ``accept()`` and ``reject()`` check every story part the flush may consult
  before the listing, and one that cannot be read raises its own refusal, its
  ``unreadable`` (D8), with nothing changed;
- **the check unmarshals nothing**, and a raw parse that succeeds is remembered
  on the package by part name, so a second resolution parses none of them again;
- ``Removals._index`` of a part that cannot be parsed is None, "unknown", never
  an empty index;
- a dry run's body resolves **locally** (CR-010 section 10.3), and checks nothing.

A content API deletion, which settles after its edit and keeps instead, is
``tests/content/test_unreadable_story_part_deletes.py``'s.

**The fixture** is ``samples/Headers.docx`` made in code (:func:`built`):
comment 3 referenced from a tracked deletion in the body and again from
``word/header1.xml``; footnote 1 referenced from another tracked deletion; and
the part a test cuts short by its last three bytes, as section 11.4 row 8's
measurement cut the header, after the reference.
"""

from __future__ import annotations

import io
import zipfile
from collections.abc import Callable
from typing import Any

import pytest
from revision_helpers import ROOT, revisions

from docx4j_py import load
from docx4j_py.model.revisions import TrackedChangeFilter
from docx4j_py.model.revisions import removals as removals_module
from docx4j_py.model.revisions.removals import Removals
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.parts.xml_part import syntax_error_of

SOURCE = ROOT / "samples" / "Headers.docx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

HEADER = "word/header1.xml"
FOOTNOTES = "word/footnotes.xml"
COMMENTS = "word/comments.xml"

#: Two tracked deletions at the body's start: comment 3's reference, then footnote 1's.
_DELETIONS = (
    b'<w:p><w:del w:id="901" w:author="A" w:date="2026-11-04T00:00:00Z"><w:r><w:commentReference w:id="3"/>'
    b"</w:r></w:del></w:p>"
    b'<w:p><w:del w:id="902" w:author="A" w:date="2026-11-04T00:00:00Z"><w:r><w:footnoteReference w:id="1"/>'
    b"</w:r></w:del></w:p>"
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

_BUILT: dict[str | None, bytes] = {}


def built(cut: str | None = None) -> bytes:
    """The fixture, with the entry `cut` (when given) short by its last three bytes."""
    if cut not in _BUILT:
        edits: dict[str, Callable[[bytes], bytes]] = {
            "word/document.xml": lambda data: data.replace(b"<w:body>", b"<w:body>" + _DELETIONS, 1),
            HEADER: lambda data: data.replace(b"</w:hdr>", _HEADER_REFERENCE + b"</w:hdr>"),
            FOOTNOTES: lambda data: data.replace(b"</w:footnotes>", _NOTE + b"</w:footnotes>"),
            "word/_rels/document.xml.rels": lambda data: data.replace(
                b"</Relationships>", _RELATIONSHIP + b"</Relationships>"
            ),
            "[Content_Types].xml": lambda data: data.replace(b"</Types>", _OVERRIDE + b"</Types>"),
        }
        out = io.BytesIO()
        with zipfile.ZipFile(SOURCE) as source, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
            entries = [(info.filename, source.read(info)) for info in source.infolist()]
            for name, data in [*entries, (COMMENTS, _COMMENTS)]:
                edited = edits[name](data) if name in edits else data
                assert name not in edits or edited != data, name
                target.writestr(name, edited[:-3] if name == cut else edited)
        _BUILT[cut] = out.getvalue()
    return _BUILT[cut]


def entry(data: bytes, name: str) -> bytes | None:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return archive.read(name) if name in archive.namelist() else None


def _state(pkg: Any) -> tuple[bytes, ...]:
    """The main part's XML, its tree's, and the comments and footnotes parts' as they would be written."""
    main = pkg.main_document_part
    main.contents  # noqa: B018 - the listing reads it: what is compared is its tree's XML
    return main.xml, main.comments_part.xml, main.footnotes_part.xml


def _story_parts(pkg: Any) -> list[Any]:
    main = pkg.main_document_part
    return [*main.header_parts(), main.footnotes_part, main.endnotes_part, main.comments_part]


@pytest.fixture
def raw_parses(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    """The names of the parts ``Removals`` parses raw, in order, while the test runs."""
    parsed: list[str] = []
    original = removals_module._raw_root

    def counting(part: Any) -> Any:
        parsed.append(str(part.part_name))
        return original(part)

    monkeypatch.setattr(removals_module, "_raw_root", counting)
    return parsed


# ---------------------------------------------------------------------------
# the fixture, and the control
# ---------------------------------------------------------------------------


def test_the_fixture_holds_what_the_tests_need() -> None:
    pkg = load(built())
    assert len(revisions.tracked_changes(pkg)) == 2
    assert [c.id for c in pkg.main_document_part.comments_part.contents.comment] == [3]
    assert b'<w:commentReference w:id="3"/>' in entry(built(), HEADER)
    for cut in (HEADER, FOOTNOTES):
        # the cut is after the reference and the note: only the closing tag goes
        assert entry(built(cut), cut) == entry(built(), cut)[:-3]
        part = load(built(cut)).get_part("/" + cut)
        with pytest.raises(Docx4JException) as raised:
            part.read_contents()
        assert syntax_error_of(raised.value) is not None


def test_with_the_header_intact_comment_3_is_kept_by_its_header_reference() -> None:
    """The control: accepting the body's deletion takes its reference, and the header's keeps the comment."""
    data = built()
    pkg = load(data)
    assert revisions.accept_all(pkg) == 2
    main = pkg.main_document_part
    assert [c.id for c in main.comments_part.contents.comment] == [3]
    # the note's only reference went with its deletion, and the note with it
    assert [n.id for n in main.footnotes_part.contents.footnote] == [-1, 0]
    saved = pkg.save()
    assert entry(saved, HEADER) == entry(data, HEADER)
    assert all(not part.is_unmarshalled for part in main.header_parts())


# ---------------------------------------------------------------------------
# a resolution refuses first
# ---------------------------------------------------------------------------

#: Every route to a resolution of every story (section 4.6b's first bullet).
ROUTES: dict[str, Callable[[Any], Any]] = {
    "accept_all": revisions.accept_all,
    "reject_all": revisions.reject_all,
    "accept_all, filtered": lambda pkg: revisions.accept_all(pkg, TrackedChangeFilter(authors=["A"])),
    "TrackedChange.accept": lambda pkg: revisions.tracked_changes(pkg)[0].accept(),
    "TrackedChange.reject": lambda pkg: revisions.tracked_changes(pkg)[0].reject(),
    "Body.accept_all": lambda pkg: pkg.body.accept_all(),
    "Body.reject_all": lambda pkg: pkg.body.reject_all(),
    "the content API's accept": lambda pkg: pkg.body.get_tracked_changes()[0].accept(),
    "the content API's reject": lambda pkg: pkg.body.get_tracked_changes()[0].reject(),
}


@pytest.mark.parametrize("route", list(ROUTES))
def test_a_resolution_refuses_a_cut_header_with_its_own_refusal_and_changes_nothing(route: str) -> None:
    data = built(HEADER)
    pkg = load(data)
    header = pkg.get_part("/" + HEADER)
    before = _state(pkg)
    with pytest.raises(Docx4JException) as raised:
        ROUTES[route](pkg)
    assert raised.value is header.unreadable
    assert str(raised.value).startswith("Problem with part /word/header1.xml: ")
    assert syntax_error_of(raised.value) is not None
    # nothing changed, and nothing read that was not read before
    assert _state(pkg) == before
    assert not header.is_unmarshalled
    assert all(not part.is_unmarshalled for part in _story_parts(pkg))
    assert len(revisions.tracked_changes(pkg)) == 2
    saved = pkg.save()
    for name in (HEADER, COMMENTS, FOOTNOTES):
        assert entry(saved, name) == entry(data, name), name


def test_a_second_refusal_raises_the_record_again_without_a_parse(raw_parses: list[str]) -> None:
    pkg = load(built(HEADER))
    header = pkg.get_part("/" + HEADER)
    with pytest.raises(Docx4JException) as first:
        revisions.accept_all(pkg)
    assert raw_parses.count("/word/header1.xml") == 1
    raw_parses.clear()
    with pytest.raises(Docx4JException) as second:
        revisions.reject_all(pkg)
    assert second.value is first.value is header.unreadable
    assert "/word/header1.xml" not in raw_parses


def test_a_part_whose_read_failed_raises_its_record(raw_parses: list[str]) -> None:
    """A part another reader already failed to read is not parsed raw: its record is raised again (D8)."""
    pkg = load(built(HEADER))
    header = pkg.get_part("/" + HEADER)
    with pytest.raises(Docx4JException) as read:
        header.read_contents()
    with pytest.raises(Docx4JException) as raised:
        revisions.accept_all(pkg)
    assert raised.value is read.value is header.unreadable
    assert "/word/header1.xml" not in raw_parses


@pytest.mark.parametrize(
    "route",
    [revisions.accept_all, lambda pkg: revisions.tracked_changes(pkg)[1].accept()],
    ids=["accept_all", "the note's TrackedChange.accept"],
)
def test_a_cut_footnotes_part_refuses_an_accept_that_removes_a_note_reference(route: Callable[[Any], Any]) -> None:
    """Before Phase A3 the accept removed the reference, then raised at the note: the body changed, the save after it."""
    data = built(FOOTNOTES)
    pkg = load(data)
    footnotes = pkg.main_document_part.footnotes_part
    before = _state(pkg)
    with pytest.raises(Docx4JException) as raised:
        route(pkg)
    assert raised.value is footnotes.unreadable
    assert str(raised.value).startswith("Problem with part /word/footnotes.xml: ")
    assert _state(pkg) == before
    assert len(revisions.tracked_changes(pkg)) == 2
    assert entry(pkg.save(), FOOTNOTES) == entry(data, FOOTNOTES)


def test_a_dry_run_resolves_locally_and_checks_nothing() -> None:
    """CR-010 section 10.3: a trial's body touches no part outside its own, so it needs none of them."""
    pkg = load(built(HEADER))
    header = pkg.get_part("/" + HEADER)
    before = _state(pkg)
    with pkg.dry_run() as trial:
        assert trial.body.accept_all() == 2
        assert trial.body.get_tracked_changes() == []
    assert header.unreadable is None and not header.is_unmarshalled
    assert _state(pkg) == before
    assert len(revisions.tracked_changes(pkg)) == 2


# ---------------------------------------------------------------------------
# the check unmarshals nothing, and is remembered
# ---------------------------------------------------------------------------


def test_a_second_resolution_parses_no_story_part_raw(raw_parses: list[str]) -> None:
    data = built()
    pkg = load(data)
    unread = {str(part.part_name) for part in _story_parts(pkg)}
    assert revisions.accept_all(pkg) == 2
    # the check parsed each story part nobody had read, and the flush read the header for comment 3's reference
    assert unread <= set(raw_parses)
    assert pkg._well_formed_sources == unread
    raw_parses.clear()
    assert revisions.accept_all(pkg) == 0
    assert revisions.reject_all(pkg) == 0
    assert pkg.body.accept_all() == 0
    assert raw_parses == []
    # a well-formed header nobody read is still not unmarshalled, and is written byte for byte
    main = pkg.main_document_part
    assert all(not part.is_unmarshalled for part in main.header_parts())
    saved = pkg.save()
    for part in main.header_parts():
        name = str(part.part_name).lstrip("/")
        assert entry(saved, name) == entry(data, name), name


def test_the_check_unmarshals_no_well_formed_part(raw_parses: list[str]) -> None:
    """A reject keeps both references, so nothing is settled: the check alone read the story parts."""
    data = built()
    pkg = load(data)
    assert revisions.reject_all(pkg) == 2
    assert all(not part.is_unmarshalled and part.unreadable is None for part in _story_parts(pkg))
    assert set(raw_parses) == {str(part.part_name) for part in _story_parts(pkg)}
    assert len(raw_parses) == len(set(raw_parses))
    saved = pkg.save()
    for part in _story_parts(pkg):
        name = str(part.part_name).lstrip("/")
        assert entry(saved, name) == entry(data, name), name


def test_bytes_set_since_the_check_are_checked_again() -> None:
    """The memo holds a part's source bytes: bytes set since are parsed afresh, and refused."""
    pkg = load(built())
    assert revisions.reject_all(pkg) == 2
    assert "/word/header1.xml" in pkg._well_formed_sources
    header = pkg.get_part("/" + HEADER)
    header.set_bytes(entry(built(HEADER), HEADER))
    with pytest.raises(Docx4JException) as raised:
        revisions.accept_all(pkg)
    assert raised.value is header.unreadable


# ---------------------------------------------------------------------------
# the index
# ---------------------------------------------------------------------------


def test_the_index_of_a_part_that_cannot_be_parsed_is_unknown() -> None:
    intact, cut = load(built()), load(built(HEADER))
    assert Removals(intact)._index(intact.get_part("/" + HEADER)) == {"commentReference": {3}}
    header = cut.get_part("/" + HEADER)
    assert Removals(cut)._index(header) is None
    # nor is a part whose read failed indexed from its bytes
    with pytest.raises(Docx4JException):
        header.read_contents()
    assert Removals(cut)._index(header) is None
