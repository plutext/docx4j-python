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
"""

from __future__ import annotations

import pytest
from revision_helpers import assert_like_word, fixture, revisions, unresolved

#: (tracked, Word's save, accept)
CASES = [
    ("cr043/cr043-tracked.docx", "cr043/cr043-accept-all.docx", True),
    ("cr043/cr043-tracked.docx", "cr043/cr043-reject-all.docx", False),
    ("tracked-changes-equations.docx", "cr043/equations-accept-all.docx", True),
    ("tracked-changes-equations.docx", "cr043/equations-reject-all.docx", False),
    ("editor/revisions/revisions-word15.docx", "core-ts/check18/accept-all.docx", True),
    ("editor/revisions/revisions-word15.docx", "core-ts/check18/reject-all.docx", False),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word15-accepted.docx", True),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word15-rejected.docx", False),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word2010-accepted.docx", True),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word2010-rejected.docx", False),
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
