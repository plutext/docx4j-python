"""Every Word-made fixture of docx4j's, resolved whole and one change at a time. CR-010 section 3.6.

The two sweeping tests of docx4j-core-tests' ``TrackedChangesWordTest``, over
every ``.docx`` under its ``src/test/resources/revisions/`` (128 files: checks
16 to 36, the editor's revisions and review sessions, the CR-043 probe), read
from the sibling checkout and skipped without it:

* accepted all and rejected all, nothing this release resolves is left in any
  story or branch, and the result saves and loads again;
* resolved one change at a time, listing again after each, accepting and
  rejecting in turn: each resolution takes its change away (the listing
  shrinks), and the result saves and loads again.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from revision_helpers import SIBLING, revisions, round_trip, unresolved

from docx4j_py import load

pytestmark = pytest.mark.skipif(
    not SIBLING.is_dir(), reason=f"docx4j's revisions fixtures are not checked out at {SIBLING}"
)

FILES = sorted(SIBLING.rglob("*.docx")) if SIBLING.is_dir() else []


def _id(path: Path) -> str:
    return str(path.relative_to(SIBLING))


def test_there_are_more_than_a_hundred():
    assert len(FILES) > 100


@pytest.mark.parametrize("path", FILES, ids=[_id(p) for p in FILES])
def test_every_fixture_resolves_whole(path):
    for verb in (revisions.accept_all, revisions.reject_all):
        package = load(path)
        verb(package)
        assert unresolved(package) == [], f"{verb.__name__} of {path.name}"
        round_trip(package)


@pytest.mark.parametrize("path", FILES, ids=[_id(p) for p in FILES])
def test_every_change_one_by_one(path):
    package = load(path)
    accept = True
    for _round in range(500):
        changes = [change for change in revisions.tracked_changes(package) if change.resolvable]
        if not changes:
            break
        first = changes[0]
        (first.accept if accept else first.reject)()
        accept = not accept
        after = [change for change in revisions.tracked_changes(package) if change.resolvable]
        assert len(after) < len(changes), f"{first!r} left {len(after)} of {len(changes)}"
    else:
        raise AssertionError("no end")
    round_trip(package)
