"""One change accepted or rejected, against Word's Review tab doing the same. CR-010 section 3.6.

The port of docx4j-core-tests' ``TrackedChangesOneChangeTest``: Accept This
Change and Reject Change on the Word-made fixtures --- docx4j-core-ts's checks
18, 20 and 22 --- compared with Word's save after the same click. The change is
the one the check's tester clicked in, found by its text and type; the rest of
the document keeps its revisions, which the comparison includes.

Check 22's files are copied here; checks 18's single-change saves and check
20's are read from the sibling docx4j checkout, and skip without it.
"""

from __future__ import annotations

import pytest
from revision_helpers import SIBLING, assert_like_word, fixture, revisions

from docx4j_py import load


def find(changes: list, needle: str, kind: str) -> object:
    """The first change listed whose text holds `needle`, of the given type."""
    for change in changes:
        if change.type == kind and needle in change.text:
            return change
    # a move's destination whose text was all deleted again (check 20 case 08)
    # has no accepted text: the tester clicked in its deleted words
    for change in changes:
        if change.type == kind and change.kind == "move" and change.text == "":
            return change
    return None


def assert_one(package: object, golden: object, needle: str, kind: str, accept: bool) -> None:
    change = find(revisions.tracked_changes(package), needle, kind)
    assert change is not None, f"no {kind} change holding {needle!r}"
    (change.accept if accept else change.reject)()
    assert_like_word(package, golden, records=False, label=f"{'accept' if accept else 'reject'} {change!r}")


# ---------------------------------------------------------------------------
# check 22: a deleted paragraph and a deletion across a mark, one change each
# ---------------------------------------------------------------------------

C22 = "core-ts/check22/"


@pytest.mark.parametrize(
    ("golden", "needle", "accept"),
    [
        ("22-accept-whole.docx", "Two.", True),
        ("22-reject-whole.docx", "Two.", False),
        ("22-accept-across.docx", "six.", True),
        ("22-reject-across.docx", "six.", False),
    ],
)
def test_check_22_one_change(golden, needle, accept):
    assert_one(fixture(C22 + "22-deleted.docx"), fixture(C22 + golden), needle, "Deleted", accept)


# ---------------------------------------------------------------------------
# check 18: a move, from either half (Word's saves in the sibling checkout)
# ---------------------------------------------------------------------------

C18 = SIBLING / "core-ts" / "check18"
MOVED = "The second paragraph of the move"

needs_check18 = pytest.mark.skipif(not C18.is_dir(), reason=f"docx4j's check 18 saves are not at {C18}")


@needs_check18
@pytest.mark.parametrize(
    ("golden", "kind", "accept"),
    [
        ("accept-destination.docx", "Added", True),
        ("accept-source.docx", "Deleted", True),
        ("reject-destination.docx", "Added", False),
        ("reject-source.docx", "Deleted", False),
    ],
)
def test_check_18_a_move_from_either_half(golden, kind, accept):
    assert_one(fixture("editor/revisions/revisions-word15.docx"), load(C18 / golden), MOVED, kind, accept)


# ---------------------------------------------------------------------------
# check 20: edited and broken moves, resolved from the destination's "second"
# ---------------------------------------------------------------------------

C20 = SIBLING / "core-ts" / "check20"

needs_check20 = pytest.mark.skipif(not C20.is_dir(), reason=f"docx4j's check 20 files are not at {C20}")

CHECK20 = [
    "01-type-inside-a",
    "03-type-at-start",
    "04-type-at-end",
    "06-delete-part-a",
    "07-delete-part-b",
    "08-delete-all",
    "10-move-again",
    "12-partner-missing",
    "13-moveTo-no-ranges",
]


@needs_check20
@pytest.mark.parametrize("accept", [True, False], ids=["accept", "reject"])
@pytest.mark.parametrize("name", CHECK20)
def test_check_20_edited_and_broken_moves(name, accept):
    golden = C20 / f"{name}{'-accept' if accept else '-reject'}.docx"
    assert_one(load(C20 / f"{name}.docx"), load(golden), "second", "Added", accept)


@needs_check20
@pytest.mark.parametrize(
    ("golden", "kind", "accept"),
    [
        ("14-moveFrom-no-ranges-accept-input.docx", "Deleted", True),
        ("14-moveFrom-no-ranges-reject-input.docx", "Deleted", False),
        ("14-moveFrom-no-ranges-accept-dest-input.docx", "Added", True),
        ("14-moveFrom-no-ranges-reject-dest-input.docx", "Added", False),
    ],
)
def test_check_20_a_move_from_with_no_ranges(golden, kind, accept):
    assert_one(load(C20 / "moveFrom-no-ranges.docx"), load(C20 / golden), "second", kind, accept)
