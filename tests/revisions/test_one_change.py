"""One change accepted or rejected, against Word's Review tab doing the same. CR-010 section 3.6.

The port of docx4j-core-tests' ``TrackedChangesOneChangeTest``: Accept This
Change and Reject Change on the Word-made fixtures --- docx4j-core-ts's checks
18, 20 and 22 --- compared with Word's save after the same click. The change is
the one the check's tester clicked in, found by its text and type; the rest of
the document keeps its revisions, which the comparison includes.

Check 22's files are copied here; checks 18's single-change saves and check
20's are read from the sibling docx4j checkout, and skip without it.

Word's Accept All rules where a mark taken away has no paragraph to join
(CR-010 sections 13 and 13.3a; docx4j ``29dbf5964`` and ``e55a83b8e``) have no
one-change saves: each case's changes of check 54 and of docx4j's text-left
probe are resolved one at a time, and the case's region compared with Word's
Accept All of the whole file, or with the original for the reject side.
"""

from __future__ import annotations

import pytest
from revision_helpers import (
    FIXTURES,
    SIBLING,
    assert_like_word,
    describe,
    fixture,
    revisions,
    round_trip,
)

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


# ---------------------------------------------------------------------------
# check 54 (CR-010 section 13): the three Accept All rules, one change at a time
# ---------------------------------------------------------------------------

#: Office JS's checks' files and Word's saves of them (CR-013 Phase C's fixtures, with ``SHA1SUMS``).
OFFICE_JS = FIXTURES.parent.parent / "office_js"

#: Each case's changes, by a word of their text: Word's save lists 9's and 13's as one change each (the
#: paragraph's text and mark), and 19's as two (the head's tail and mark, and the control's first word).
CHECK54 = {9: ("BODY9A",), 13: ("LIST13A",), 19: ("HEAD19B", "BODY19A")}


def region(description: str, n: int) -> list[str]:
    """Case `n`'s lines of a description, between its ``BEGIN<n>`` and ``END<n>`` paragraphs.

    The lists in it named again by first appearance in the region, since the
    rest of the document is not the same in the two files compared.
    """
    import re

    lines = description.split("\n")
    start = next(i for i, line in enumerate(lines) if f"«BEGIN{n} " in line)
    end = next(i for i, line in enumerate(lines) if f"«END{n}»" in line and i > start)
    names: dict[str, str] = {}

    def renamed(match: re.Match) -> str:
        return names.setdefault(match.group(1), f"num=@{len(names) + 1}/")

    return [re.sub(r"num=(#\d+)/", renamed, line) for line in lines[start + 1 : end]]


@pytest.mark.parametrize("accept", [True, False], ids=["accept", "reject"])
@pytest.mark.parametrize("n", sorted(CHECK54), ids=[f"case-{n}" for n in sorted(CHECK54)])
def test_check_54_a_cases_changes_one_at_a_time(n, accept):
    """Each of the case's changes accepted (rejected) on its own, first to last: the region Word's Accept All left.

    Rejected, the region the file Word read held: the reject side of the
    three rules is unmeasured, and none of them is reached by rejecting a
    deletion. Accepted, rule 1 (case 9), rule 2 (13) and rule 3 (19, its two
    changes in either order) give Word's Accept All of the region.
    """
    package = load(OFFICE_JS / "54-writes-word365.docx")
    golden = load(OFFICE_JS / ("54-writes-word365-accepted.docx" if accept else "54-writes.docx"))
    needles = CHECK54[n]
    for _round in range(10):
        mine = [change for change in revisions.tracked_changes(package) if any(w in change.text for w in needles)]
        if not mine:
            break
        (mine[0].accept if accept else mine[0].reject)()
    else:
        raise AssertionError("no end")
    ours = region(describe(round_trip(package), False), n)
    assert ours == region(describe(golden, False), n)


#: docx4j's text-left probe (``e55a83b8e``, CR-010 section 13.3a): each case's change, by a word of its text. Case 4, a
#: deleted mark on a cell's last paragraph, is not followed (Word collapsed the one-cell table), as docx4j does not.
S17 = "cr043/s17/"
PROBE = {1: ("HEAD1B",), 2: ("LIST2B",), 3: ("HEAD3A",), 5: ("HEAD5B",)}


@pytest.mark.parametrize("accept", [True, False], ids=["accept", "reject"])
@pytest.mark.parametrize("n", sorted(PROBE), ids=[f"case-{n}" for n in sorted(PROBE)])
def test_the_text_left_probes_cases_one_change_at_a_time(n, accept):
    """Each case's change accepted (rejected) on its own: the region Word 365's Accept All (Reject All) of the probe left.

    Text left before a table joins the first cell's first paragraph (case 1);
    a control's last paragraph with text left keeps a bare mark (2); an
    emptied paragraph before a control goes (3); text left before a control
    takes its first paragraph's properties, its number among them (5).
    """
    package = fixture(S17 + "17-text-left.docx")
    golden = fixture(S17 + ("17-text-left-accepted.docx" if accept else "17-text-left-rejected.docx"))
    for _round in range(10):
        mine = [change for change in revisions.tracked_changes(package) if any(w in change.text for w in PROBE[n])]
        if not mine:
            break
        (mine[0].accept if accept else mine[0].reject)()
    else:
        raise AssertionError("no end")
    assert region(describe(round_trip(package), False), n) == region(describe(golden, False), n)


def test_check_54_case_19_in_the_other_order_too():
    """Rule 3 from the control's side first: the control's first word accepted, then the head's tail and mark."""
    package = load(OFFICE_JS / "54-writes-word365.docx")
    for needle in ("BODY19A", "HEAD19B"):
        [change] = [c for c in revisions.tracked_changes(package) if needle in c.text]
        change.accept()
    golden = load(OFFICE_JS / "54-writes-word365-accepted.docx")
    assert region(describe(round_trip(package), False), 19) == region(describe(golden, False), 19)
