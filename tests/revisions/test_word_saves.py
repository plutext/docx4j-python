"""CR-010's oracle: Accept All and Reject All against Word's own saves. CR-010 section 3.6.

The port of docx4j-core-tests' ``TrackedChangesWordTest`` (docx4j CR-043
section 12): each Word-made fixture of the agreement run resolved whole --- by
:func:`docx4j_py.model.revisions.accept_all` / ``reject_all``, every story ---
saved and loaded again, and compared structurally with Word 365's own save of
the same file (``description.describe``, the port of ``DocumentDescription``).
Nothing may be left, in any story or branch.

Since Phase B (CR-010 section 5) the comparison shows the table, row, cell and
section properties and the records' presence (``records=True``), as docx4j's
phase 2 made it: the probe's C23 (a row indent: ``w:tblPrExChange``,
``w:trPrChange``, ``w:tcPrChange`` and a merged grid) and C07 (a paragraph's
section) both ways, check 18's table and the body's section, check 28's two
tables.

CR-043's U21 to U23 (docx4j ``a8bc6cda9``, Word 365's saves of 21 October
2026; CR-010 section 12.9) are ``u21AcceptAll``, ``u22RejectOne`` and
``u23RejectAll``: a comment on a deletion whose markers lie outside it, the
only comment removed by one rejected insertion, and a comment on two inserted
paragraphs, its start in the first and its end in the second.
"""

from __future__ import annotations

import pytest
from description import differing
from revision_helpers import (
    assert_like_word,
    describe,
    fixture,
    revisions,
    round_trip,
    unresolved,
)

from docx4j_py import load

#: (tracked, Word's save, accept)
CASES = [
    ("cr043/cr043-tracked.docx", "cr043/cr043-accept-all.docx", True),
    ("cr043/cr043-tracked.docx", "cr043/cr043-reject-all.docx", False),
    ("tracked-changes-equations.docx", "cr043/equations-accept-all.docx", True),
    ("tracked-changes-equations.docx", "cr043/equations-reject-all.docx", False),
    ("editor/revisions/revisions-word15.docx", "core-ts/check18/accept-all.docx", True),
    ("editor/revisions/revisions-word15.docx", "core-ts/check18/reject-all.docx", False),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word15-accepted.docx", True),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word2010-accepted.docx", True),
    # Word 365 removes the comment on the inserted paragraph whose markers lie outside its w:ins
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word365-rejected.docx", False),
    # docx4j-python's artefact 8 before fix 1: a comment on inserted text, every marker outside the w:ins
    ("python/artefact8/8-tracked-changes.docx", "python/artefact8/8-word-accepted.docx", True),
    ("python/artefact8/8-tracked-changes.docx", "python/artefact8/8-word-rejected.docx", False),
    # check 28: bold over two runs, and two tables' records (w:trPrChange included)
    ("core-ts/check28/28a-changes.docx", "core-ts/check28/28a-accept-all.docx", True),
    ("core-ts/check28/28a-changes.docx", "core-ts/check28/28a-reject-all.docx", False),
    # check 26: a paragraph centred, and one indented, rejected: back to the input Word saved first
    ("core-ts/check26/26a-centred.docx", "core-ts/check26/26-input.docx", False),
    ("core-ts/check26/26d-indented.docx", "core-ts/check26/26-input.docx", False),
    # check 33: a whole paragraph deleted, accepted: nothing carries to the next
    ("core-ts/check33/33a-deleted.docx", "core-ts/check33/33b-accepted.docx", True),
    # check 25: a list change rejected, the level Word writes back
    ("core-ts/check25/25a-bullets.docx", "core-ts/check25/25b-bullets-reject.docx", False),
    ("core-ts/check25/25e-bullet-demote.docx", "core-ts/check25/25f-bullet-demote-reject.docx", False),
    # CR-043 U21: Accept All of a deletion whose comment's markers lie outside the w:del; a control comment stays
    (
        "cr043/comments/c43-u21-accept-all-deletion-comment-outside.docx",
        "cr043/comments/c43-u21-word365-accepted.docx",
        True,
    ),
    # CR-043 U23: Reject All of two inserted paragraphs commented whole, the range across them; a control stays
    (
        "cr043/comments/c43-u23-reject-all-two-paragraph-insertion-commented.docx",
        "cr043/comments/c43-u23-word365-rejected.docx",
        False,
    ),
]


def _id(case: tuple) -> str:
    _tracked, golden, accept = case
    return f"{'accept' if accept else 'reject'}-{golden.rpartition('/')[2][:-5]}"


@pytest.mark.parametrize("case", CASES, ids=[_id(c) for c in CASES])
def test_resolving_every_story_gives_words_save(case):
    tracked, golden, accept = case
    package = fixture(tracked)
    (revisions.accept_all if accept else revisions.reject_all)(package)
    left = unresolved(package)
    assert left == [], f"{tracked}: {left}"
    assert_like_word(package, fixture(golden), records=True, label=golden)


#: The fixtures whose revisions are all in the main document: there, a body's
#: accept_all() is the document's (CR-010 section 3.3).
BODY_CASES = [case for case in CASES if not case[0].startswith(("cr043/cr043", "editor/revisions"))]


@pytest.mark.parametrize("case", BODY_CASES, ids=[_id(c) for c in BODY_CASES])
def test_a_bodys_accept_all_gives_words_save_where_the_body_is_the_document(case):
    tracked, golden, accept = case
    package = fixture(tracked)
    body = package.body
    count = body.accept_all() if accept else body.reject_all()
    assert count > 0
    assert body.get_tracked_changes() == []
    assert_like_word(package, fixture(golden), records=True, label=golden)


def test_the_probe_resolved_leaves_nothing_listed():
    """Every change the probe lists is resolved and counted, its table's and its section's records with them."""
    for verb in (revisions.accept_all, revisions.reject_all):
        package = fixture("cr043/cr043-tracked.docx")
        listed = revisions.tracked_changes(package)
        assert len(listed) == 37
        records = [str(change.kind) for change in listed if change.kind in ("table_properties", "section_properties")]
        assert records == ["table_properties", "section_properties"]
        assert verb(package) == 37
        assert revisions.tracked_changes(package) == []
        assert unresolved(package) == []


#: Word 2010's and Word 15's Reject All of the review file: they keep the comment on the inserted
#: paragraph whose markers lie outside its ``w:ins``, its range markers gone, where Word 365 and
#: the engine remove it (CR-010 section 12.8, fix 2).
OLDER_REJECTED = [
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word15-rejected.docx"),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word2010-rejected.docx"),
]

#: The comment those saves keep, found by its text (Word renumbers comments on save).
OLDER_COMMENT = "Changed from arrears, as agreed."


def _without_comment(word, text):
    """Word's save with one comment taken away whole, as the engine's ``Removals`` takes one."""
    from docx4j_py.model.revisions.removals import Removals
    from docx4j_py.traversal import text_of

    found = [c.id for c in word.main_document_part.comments_part.contents.comment if text in text_of(c)]
    assert len(found) == 1, f"no comment {text!r}"
    older = Removals(word)
    older.remove_comment(found[0])
    older.flush()
    return word


@pytest.mark.parametrize("verb", ["package", "body"])
@pytest.mark.parametrize("case", OLDER_REJECTED, ids=[_id((*c, False)) for c in OLDER_REJECTED])
def test_older_words_reject_all_differs_by_that_one_comment(case, verb):
    """docx4j's ``assertAllButComment``: the older Words' save, without the comment, is ours."""
    tracked, golden = case
    package = fixture(tracked)
    if verb == "package":
        revisions.reject_all(package)
    else:
        package.body.reject_all()
    assert unresolved(package) == []
    theirs = describe(round_trip(_without_comment(fixture(golden), OLDER_COMMENT)))
    ours = describe(round_trip(package))
    assert ours == theirs, f"{golden}: differs from Word's save but the comment {OLDER_COMMENT!r}"
    # and with it, by that comment's lines alone
    assert differing(describe(fixture(golden)), ours) > 0


_COMMENT_PARTS = ("/word/comments.xml", "/word/commentsExtended.xml", "/word/commentsIds.xml", "/word/people.xml")


def test_artefact_8_rejected_takes_the_comment_parts_and_people_with_its_last_comment():
    """docx4j's ``python8RejectAll``: no comment left, so no comment parts, and no revision left to need people.xml."""
    import io
    import zipfile

    package = fixture("python/artefact8/8-tracked-changes.docx")
    revisions.reject_all(package)
    data = package.save()
    saved = load(io.BytesIO(data))
    main = saved.main_document_part
    assert main.comments_part is None
    assert main.comments_extended_part is None
    assert main.comments_ids_part is None
    assert main.people_part is None
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = set(archive.namelist())
        types = archive.read("[Content_Types].xml").decode()
        rels = archive.read("word/_rels/document.xml.rels").decode()
    for name in _COMMENT_PARTS:
        assert saved.get_part(name) is None, name
        assert name[1:] not in names, name
        assert name not in types, name
        assert name.rpartition("/")[2] not in rels, name


COMMENTS = "cr043/comments/"


@pytest.mark.parametrize("route", ["engine", "content"])
def test_u22_one_insertion_rejected_takes_the_only_comment_and_its_parts_but_not_people(route):
    """docx4j ``u22RejectOne`` (CR-043 U22, Word 365): ``people.xml`` stays while a tracked deletion is left.

    Only the change whose text is ``C43_U22_INSERTED`` is rejected, through the
    engine's change and through the content API's.
    """
    tracked = COMMENTS + "c43-u22-reject-one-last-comment-revisions-left.docx"
    package = fixture(tracked)
    listed = revisions.tracked_changes(package) if route == "engine" else package.body.get_tracked_changes()
    insertion = [change for change in listed if change.text == "C43_U22_INSERTED"]
    assert len(insertion) == 1 and insertion[0].type == "Added"
    insertion[0].reject()
    word = fixture(COMMENTS + "c43-u22-word365-rejected-one.docx")
    assert_like_word(package, word, records=True, label=f"{tracked}, the insertion rejected")
    saved = round_trip(package)
    main = saved.main_document_part
    assert main.comments_part is None
    assert main.comments_extended_part is None and main.comments_ids_part is None
    assert main.people_part is not None, "people.xml stays while a tracked change is left"
    assert [change.text for change in revisions.tracked_changes(saved)] == ["C43_U22_STILL_TRACKED_DELETION"]


def test_u23_reject_all_removes_the_comment_across_two_paragraphs_and_keeps_the_control():
    """docx4j ``u23RejectAll``: its start in the first inserted paragraph, its end in the second; one comment left."""
    from docx4j_py.traversal import text_of

    package = fixture(COMMENTS + "c43-u23-reject-all-two-paragraph-insertion-commented.docx")
    revisions.reject_all(package)
    comments = round_trip(package).main_document_part.comments_part.contents.comment
    assert [text_of(comment) for comment in comments] == ["C43_U23_CONTROL"]
