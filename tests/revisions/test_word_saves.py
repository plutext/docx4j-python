"""CR-010's oracle: Accept All and Reject All against Word's own saves. CR-010 section 3.6.

The port of docx4j-core-tests' ``TrackedChangesWordTest`` (docx4j CR-043
section 12): each Word-made fixture of the agreement run resolved whole --- by
:func:`docx4j_py.model.revisions.accept_all` / ``reject_all``, every story ---
saved and loaded again, and compared structurally with Word 365's own save of
the same file (``description.describe``, the port of ``DocumentDescription``).
Nothing this release resolves may be left, in any story or branch.

Phase A compares **without** the table, row, cell and section properties
(``records=False``), as docx4j's phase 1 did: those records are CR-010 Phase
B's. Where a fixture holds none, the comparison is made with them shown too.
"""

from __future__ import annotations

import pytest
from revision_helpers import assert_like_word, fixture, revisions, unresolved

#: (tracked, Word's save, accept, whether the fixture holds no Phase B records)
CASES = [
    ("cr043/cr043-tracked.docx", "cr043/cr043-accept-all.docx", True, False),
    ("cr043/cr043-tracked.docx", "cr043/cr043-reject-all.docx", False, False),
    ("tracked-changes-equations.docx", "cr043/equations-accept-all.docx", True, True),
    ("tracked-changes-equations.docx", "cr043/equations-reject-all.docx", False, True),
    ("editor/revisions/revisions-word15.docx", "core-ts/check18/accept-all.docx", True, False),
    ("editor/revisions/revisions-word15.docx", "core-ts/check18/reject-all.docx", False, False),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word15-accepted.docx", True, True),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word15-rejected.docx", False, True),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word2010-accepted.docx", True, True),
    ("editor/review/review-for-word.docx", "editor/review/review-for-word-word2010-rejected.docx", False, True),
    ("core-ts/check28/28a-changes.docx", "core-ts/check28/28a-accept-all.docx", True, False),
    ("core-ts/check28/28a-changes.docx", "core-ts/check28/28a-reject-all.docx", False, False),
    # check 26: a paragraph centred, and one indented, rejected: back to the input Word saved first
    ("core-ts/check26/26a-centred.docx", "core-ts/check26/26-input.docx", False, True),
    ("core-ts/check26/26d-indented.docx", "core-ts/check26/26-input.docx", False, True),
    # check 33: a whole paragraph deleted, accepted: nothing carries to the next
    ("core-ts/check33/33a-deleted.docx", "core-ts/check33/33b-accepted.docx", True, True),
    # check 25: a list change rejected, the level Word writes back
    ("core-ts/check25/25a-bullets.docx", "core-ts/check25/25b-bullets-reject.docx", False, True),
    ("core-ts/check25/25e-bullet-demote.docx", "core-ts/check25/25f-bullet-demote-reject.docx", False, True),
]


def _id(case: tuple) -> str:
    _tracked, golden, accept, _records = case
    return f"{'accept' if accept else 'reject'}-{golden.rpartition('/')[2][:-5]}"


@pytest.mark.parametrize("case", CASES, ids=[_id(c) for c in CASES])
def test_resolving_every_story_gives_words_save(case):
    tracked, golden, accept, no_records = case
    package = fixture(tracked)
    (revisions.accept_all if accept else revisions.reject_all)(package)
    left = unresolved(package)
    assert left == [], f"{tracked}: {left}"
    word = fixture(golden)
    assert_like_word(package, word, records=False, label=f"{golden}")
    if no_records:
        assert_like_word(package, word, records=True, label=f"{golden}, records shown")


#: The fixtures whose revisions are all in the main document: there, a body's
#: accept_all() is the document's (CR-010 section 3.3).
BODY_CASES = [case for case in CASES if not case[0].startswith(("cr043/cr043", "editor/revisions"))]


@pytest.mark.parametrize("case", BODY_CASES, ids=[_id(c) for c in BODY_CASES])
def test_a_bodys_accept_all_gives_words_save_where_the_body_is_the_document(case):
    tracked, golden, accept, _records = case
    package = fixture(tracked)
    body = package.body
    count = body.accept_all() if accept else body.reject_all()
    assert count > 0
    assert body.get_tracked_changes() == [] or all(
        change.kind in ("table_properties", "section_properties") for change in body.get_tracked_changes()
    )
    assert_like_word(package, fixture(golden), records=False, label=golden)


def test_the_probe_resolved_leaves_only_phase_bs_records_listed():
    """What Phase A leaves: the table's and the section's records, listed, never a clean document."""
    for verb in (revisions.accept_all, revisions.reject_all):
        package = fixture("cr043/cr043-tracked.docx")
        count = verb(package)
        assert count == 35
        left = revisions.tracked_changes(package)
        assert sorted(str(change.kind) for change in left) == ["section_properties", "table_properties"]
        assert not any(change.resolvable for change in left)
