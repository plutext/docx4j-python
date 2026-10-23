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

docx4j-core-ts's check 54 (CR-010 section 13; docx4j ``29dbf5964``, its CR-043
section 17) is ``check54AcceptAll``: Word 365's save after 43 tracked writes,
accepted all, against Word's own Accept All, the whole document; and since
docx4j ``e55a83b8e`` (CR-010 section 13.3a) ``check54RejectAll``, against Word's
Reject All. Its files are CR-013 Phase C's, in ``tests/office_js/``, with Word's
Reject All beside them. docx4j's text-left probe (``e55a83b8e``,
``fixtures/revisions/cr043/s17/``) is ``textLeftProbeAcceptAll`` and
``textLeftProbeRejectAll``.
"""

from __future__ import annotations

import pytest
from description import differing
from revision_helpers import (
    FIXTURES,
    assert_like_word,
    describe,
    fixture,
    revisions,
    round_trip,
    unresolved,
)

from docx4j_py import load
from docx4j_py.model.revisions.trees import View

#: Office JS's checks' files and Word's saves of them (CR-013 Phase C's fixtures, with ``SHA1SUMS``).
OFFICE_JS = FIXTURES.parent.parent / "office_js"

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


@pytest.mark.parametrize("verb", ["package", "body"])
def test_check_54_accept_all_gives_words_accept_all(verb):
    """docx4j ``check54AcceptAll`` (``29dbf5964``): Word 365's Accept All of 43 tracked writes, the whole document.

    Among them the three rules of CR-010 section 13: an emptied paragraph
    before a table goes (case 9); a control's last paragraph deleted whole
    keeps a bare mark (13); a deleted mark before a block control joins
    nothing and takes the control's first paragraph's properties (19). And a
    control deleted whole in Word's ``w:customXmlDelRange*`` form goes with its
    markers, the mark before it then joining what follows (15, 17; CR-013
    ``e7f18a7``). Lists are named by first appearance, so the numbering ids
    Word's save renumbers compare once rule 2 drops the orphaned ``w:numPr``,
    as docx4j found. Word's Reject All was not saved.
    """
    package = load(OFFICE_JS / "54-writes-word365.docx")
    if verb == "package":
        assert revisions.accept_all(package) > 0
    else:
        assert package.body.accept_all() > 0
    assert unresolved(package) == []
    assert_like_word(package, load(OFFICE_JS / "54-writes-word365-accepted.docx"), records=True, label="check 54")


@pytest.mark.parametrize("verb", ["package", "body"])
def test_check_54_reject_all_gives_words_reject_all(verb):
    """docx4j ``check54RejectAll`` (``e55a83b8e``): Word 365's Reject All of check 54 is the original everywhere, as the engine's."""
    package = load(OFFICE_JS / "54-writes-word365.docx")
    if verb == "package":
        assert revisions.reject_all(package) > 0
    else:
        assert package.body.reject_all() > 0
    assert unresolved(package) == []
    assert_like_word(package, load(OFFICE_JS / "54-writes-word365-rejected.docx"), records=True, label="check 54")


#: docx4j's text-left probe (``e55a83b8e``, its CR-043 section 17), with its builder and ``SHA1SUMS``.
S17 = "cr043/s17/"


def _table_props_aside(description: str) -> str:
    """docx4j's ``tablePropsAside``: the table properties out, and the probe's region 4 out.

    Word's save adds ``w:tblInd`` and ``w:tblCellMar`` to a table that lacks
    them. Region 4, a deleted mark on a cell's last paragraph, is not an edit
    Word makes; on it Word's Accept All dropped the one-cell table and kept its
    paragraph as text, which docx4j does not follow, and neither does this
    package (U1 keeps the cell's paragraph).
    """
    import re

    description = re.sub(r"tblPr\[.*?\] grid", "tblPr[] grid", description)
    return re.sub(r"(?s)P \| «BEGIN4 .*?P \| «END4»\n", "", description)


@pytest.mark.parametrize("accept", [True, False], ids=["accept", "reject"])
def test_the_text_left_probe_resolved_gives_words_save(accept):
    """docx4j ``textLeftProbeAcceptAll`` and ``textLeftProbeRejectAll`` (``e55a83b8e``), table properties and region 4 aside.

    Word 365's Accept All: text left before a table joins the table's first
    cell's first paragraph (case 1); a control's last paragraph with text left
    keeps a bare mark (2); an emptied paragraph before a control goes (3); text
    left before a control takes its first paragraph's properties, a number
    among them (5). Its Reject All: the original.
    """
    package = fixture(S17 + "17-text-left.docx")
    (revisions.accept_all if accept else revisions.reject_all)(package)
    assert unresolved(package) == []
    golden = fixture(S17 + ("17-text-left-accepted.docx" if accept else "17-text-left-rejected.docx"))
    ours, word = describe(round_trip(package)), describe(golden)
    assert "«BEGIN4 " in word and "«BEGIN5 " in word
    assert _table_props_aside(ours) == _table_props_aside(word)


def test_the_text_left_probe_is_docx4js_as_copied():
    """The four files of docx4j's ``revisions/cr043/s17/`` at ``e55a83b8e``, byte for byte, as ``SHA1SUMS`` lists them."""
    import hashlib

    lines = (FIXTURES / S17 / "SHA1SUMS").read_text(encoding="utf-8").splitlines()
    manifest = {name: digest for digest, name in (line.split("  ", 1) for line in lines if line)}
    assert sorted(manifest) == [
        "17-text-left-accepted.docx",
        "17-text-left-rejected.docx",
        "17-text-left.docx",
        "build-17-text-left.py",
    ]
    for name, digest in manifest.items():
        assert hashlib.sha1((FIXTURES / S17 / name).read_bytes()).hexdigest() == digest, name


def _accepted_text(item):
    """A block's text in the accepted view: a paragraph's, else None."""
    from docx4j_py.model.revisions import trees

    return trees.paragraph_text(item, View.ACCEPTED) if isinstance(item, trees.cls("P")) else None


def test_check_54s_three_regions_after_accept_all_are_words():
    """Cases 9, 13 and 19 as Word's Accept All left them, where the engine left them otherwise before section 13."""
    package = load(OFFICE_JS / "54-writes-word365.docx")
    revisions.accept_all(package)
    body = round_trip(package).main_document_part.contents.body.content
    texts = [_accepted_text(item) for item in body]
    # case 9: the emptied paragraph gone, the table straight after the case's label (an empty w:jc paragraph was left)
    nine = texts.index("BEGIN9 WT delete tracked whole body9a")
    assert type(body[nine + 1]).__name__ == "Tbl"
    # case 13: the control's emptied last paragraph a bare mark (its ListParagraph and w:numPr were left)
    thirteen = body[texts.index("BEGIN13 WC delete tracked whole list13a") + 1]
    last = thirteen.sdt_content.content[-1]
    assert last.p_pr is None and last.content == []
    # case 19: the head in the control's first paragraph's properties, nothing joined (it kept Heading1)
    nineteen = texts.index("BEGIN19 CCI delete tracked from head19b to body19a")
    head, control = body[nineteen + 1], body[nineteen + 2]
    assert texts[nineteen + 1] == "HEAD19A "
    assert head.p_pr.p_style is None and head.p_pr.jc.val.value == "center"
    assert type(control).__name__ == "SdtBlock"
