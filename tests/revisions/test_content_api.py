"""The content API's reads against the engine, over the agreement fixtures. CR-010 sections 6 and 12.

- ``OutlineStats.tracked_changes`` counts the engine's listing: a body's
  outline what that body's ``get_tracked_changes()`` lists, the package's
  outline the body's and every header's and footer's (the stories an outline
  shows), and nothing once ``revisions.accept_all`` has run.
- ``describe()``'s ``tracking_on`` is ``w:trackRevisions``, which accepting
  and rejecting never change (docx4j CR-043 section 3.3).
- The text and Markdown views are **not yet** the engine's: ``get_text()`` and
  ``to_markdown(view="accepted")`` filter runs by their revision holders, so a
  deleted paragraph mark is not joined and a paragraph Word's Accept All
  removes is still a line. The user moved the agreement to CR-011, after
  docx4j CR-046 phase 0's goldens (the portfolio's text-extraction paper,
  sections 3.1 and 10; CR-010 section 12). The fixtures where the views differ
  today are pinned ``xfail(strict=True)``, measured, so that CR-011 closing
  them shows; where they agree, the agreement is held.
"""

from __future__ import annotations

from collections import Counter

import pytest
from revision_helpers import fixture, revisions, round_trip

from docx4j_py.model.content.addresses import package_bodies

#: The agreement run's eleven Word-made fixtures (CR-010 section 3.6).
TRACKED = [
    "cr043/cr043-tracked.docx",
    "tracked-changes-equations.docx",
    "editor/revisions/revisions-word15.docx",
    "editor/review/review-for-word.docx",
    "core-ts/check28/28a-changes.docx",
    "core-ts/check26/26a-centred.docx",
    "core-ts/check26/26d-indented.docx",
    "core-ts/check33/33a-deleted.docx",
    "core-ts/check25/25a-bullets.docx",
    "core-ts/check25/25e-bullet-demote.docx",
    "core-ts/check22/22-deleted.docx",
]


def _id(name: str) -> str:
    return name.rpartition("/")[2][:-5]


# ---------------------------------------------------------------------------
# the counts
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", TRACKED, ids=[_id(n) for n in TRACKED])
def test_the_outline_counts_the_engines_listing(name):
    package = fixture(name)
    by_part = Counter(id(change.part) for change in revisions.tracked_changes(package))
    shown = 0
    for body in package_bodies(package):
        listed = len(body.get_tracked_changes())
        assert body.outline().stats.tracked_changes == listed, body.prefix
        assert listed == by_part[id(body.part)], body.prefix
        if body.prefix == "body" or body.prefix.startswith(("header:", "footer:")):
            shown += listed
    assert package.outline().stats.tracked_changes == shown
    assert len(package.get_tracked_changes()) == sum(by_part.values()) == revisions.report(package).total

    revisions.accept_all(package)
    assert package.outline().stats.tracked_changes == 0
    assert round_trip(package).outline().stats.tracked_changes == 0


@pytest.mark.parametrize("accept", [True, False], ids=["accept", "reject"])
@pytest.mark.parametrize("name", TRACKED, ids=[_id(n) for n in TRACKED])
def test_tracking_on_is_the_setting_and_resolving_leaves_it(name, accept):
    package = fixture(name)
    before = package.describe().tracking_on
    assert before is (package.change_tracking_mode != "Off")
    (revisions.accept_all if accept else revisions.reject_all)(package)
    assert package.describe().tracking_on is before
    assert round_trip(package).describe().tracking_on is before


# ---------------------------------------------------------------------------
# the views: CR-011's, pinned where they differ today
# ---------------------------------------------------------------------------

#: Measured 2026-10-18: where the accepted view differs from the text after
#: ``revisions.accept_all``, and why (the same five for both views; the line
#: counts are the text view's).
DIFFERS = {
    "cr043-tracked": "C06's, C09's and C10b's deleted marks not joined; a deleted table's and row's cells as 6 empty lines",
    "revisions-word15": "a deleted mark not joined; a moved paragraph's source and a deleted row's cells as 4 empty lines",
    "review-for-word": "a deleted mark not joined; a deleted row's cells and a deleted paragraph as 4 empty lines",
    "33a-deleted": "a deleted heading left (an empty line in the text, '##' in the Markdown)",
    "22-deleted": "deleted marks not joined ('Four.\\nFive \\neight.' for 'Four.Five eight.'); 'Two.' as an empty line",
}


def _view_cases():
    for name in TRACKED:
        why = DIFFERS.get(_id(name))
        marks = [pytest.mark.xfail(strict=True, reason=f"CR-011: {why}")] if why else []
        yield pytest.param(name, marks=marks, id=_id(name))


def _accepted(name):
    package = fixture(name)
    revisions.accept_all(package)
    return package


@pytest.mark.parametrize("name", list(_view_cases()))
def test_get_text_reads_as_accept_all_leaves_the_document(name):
    assert fixture(name).body.get_text(view="accepted") == _accepted(name).body.get_text()


@pytest.mark.parametrize("name", list(_view_cases()))
def test_the_accepted_markdown_reads_as_accept_all_leaves_the_document(name):
    assert fixture(name).body.to_markdown(view="accepted") == _accepted(name).body.to_markdown()
