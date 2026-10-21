"""The views against the resolved copy (REQ-406). CR-011 sections 3.2, 3.6 and 4.

docx4j's ``ViewAgreementTest`` at ``198d015a3``, over this package's readers
as they are before the text model: the walk is ``Body.get_text(view=)`` on the
fixture as loaded, the reference the text of a ``dry_run`` copy resolved by
``revisions.accept_all`` (for ``accepted``) or ``reject_all`` (for
``original``) and read by ``get_text()``, which on a copy with nothing left to
resolve reads the same in either view. The copy is the reference: where the two
differ, the walk is wrong. A trial copies the parsed tree and saves nothing, so
it keeps the ``Requires`` prefix a save drops (probe case 6), which is why
docx4j loads its copy a second time instead.

Only the default scope: today's ``get_text`` reads the main story alone and
takes no options, where docx4j's test reads every story too.

Over three sets:

* CR-010's eleven agreement fixtures and the probe, which are in this
  repository: one test per fixture and view, ``xfail(strict=True)`` where they
  differ today, with what differs, so that each one CR-011 closes shows;
* docx4j's 147 revisions and text-extraction fixtures at ``198d015a3``,
  where the sibling checkout has them (``slow``): the same, from the outcomes
  measured in ``agreement-docx4j.tsv``. A file docx4j adds later is not in the
  table, and is not read.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from conformance import PHASE0, ROOT, table

from docx4j_py import load
from docx4j_py.model import revisions

#: docx4j-core-tests' resources, in the sibling checkout.
SIBLING = ROOT.parent / "docx4j" / "docx4j-core-tests" / "src" / "test" / "resources"

#: The agreement run's eleven Word-made fixtures (CR-010 section 3.6).
REVISIONS = ROOT / "tests" / "fixtures" / "revisions"
COPIED = [
    *(
        REVISIONS / name
        for name in (
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
        )
    ),
    PHASE0 / "extraction-probe.docx",
]

#: Measured 2026-10-21: where the walk differs from the resolved copy, and how.
#: Each is a deleted (or, in ``original``, an inserted) paragraph mark not
#: joined, or a paragraph the resolution removes left as an empty line.
DIFFERS = {
    ("cr043-tracked", "accepted"): "C06's, C09's and C10b's deleted marks not joined; "
    "a deleted table's and row's cells as 6 empty lines",
    ("cr043-tracked", "original"): "C10's inserted mark not joined; C01's inserted field result and "
    "C16's inserted link kept; C15's move destination and C25's inserted table as 6 empty lines",
    ("revisions-word15", "accepted"): "a deleted mark not joined; a moved paragraph's source and "
    "a deleted row's cells as 4 empty lines",
    ("revisions-word15", "original"): "an inserted mark not joined; a moved paragraph's destination and "
    "an inserted row's cells as 4 empty lines",
    ("review-for-word", "accepted"): "a deleted mark not joined; a deleted row's cells and "
    "a deleted paragraph as 4 empty lines",
    ("review-for-word", "original"): "an inserted mark not joined; an inserted row's cells and "
    "inserted paragraphs as 7 empty lines",
    ("33a-deleted", "accepted"): "a deleted heading left as an empty line",
    ("22-deleted", "accepted"): "deleted marks not joined ('Four.\\nFive \\neight.' for "
    "'Four.Five eight.'); 'Two.' as an empty line",
    ("extraction-probe", "accepted"): "case 11's deleted mark not joined; case 13's deleted row "
    "as an empty line",
    ("extraction-probe", "original"): "case 12's inserted mark not joined; case 13's inserted row "
    "as an empty line",
}

VIEWS = {"accepted": revisions.accept_all, "original": revisions.reject_all}

#: docx4j's fixtures at 198d015a3, under docx4j-core-tests' resources, with
#: today's outcome in each view.
DOCX4J = table(Path(__file__).resolve().parent / "agreement-docx4j.tsv")


def _walk_and_copy(path: Path, view: str) -> tuple[str, str]:
    """The walk's text in a view, and the text of a copy resolved to that view."""
    package = load(path)
    walk = package.body.get_text(view=view)
    with package.dry_run() as trial:
        VIEWS[view](trial)
        copy = trial.body.get_text()
    return walk, copy


def _first_difference(walk: str, copy: str) -> str:
    """docx4j's ``firstDifference``: the first line that differs, both ways."""
    w, c = walk.split("\n"), copy.split("\n")
    i = 0
    while i < min(len(w), len(c)) and w[i] == c[i]:
        i += 1
    return f"line {i + 1}\n  walk: {w[i] if i < len(w) else '<end>'}\n  copy: {c[i] if i < len(c) else '<end>'}"


def _assert_agree(path: Path, view: str) -> None:
    walk, copy = _walk_and_copy(path, view)
    assert walk == copy, f"{path.name} {view}: {_first_difference(walk, copy)}"


def _copied():
    for path in COPIED:
        for view in VIEWS:
            why = DIFFERS.get((path.stem, view))
            marks = [pytest.mark.xfail(strict=True, reason=f"CR-011: {why}")] if why else []
            yield pytest.param(path, view, marks=marks, id=f"{path.stem}-{view}")


@pytest.mark.parametrize(("path", "view"), list(_copied()))
def test_the_view_reads_as_the_resolved_copy(path, view):
    _assert_agree(path, view)


def test_the_pins_are_the_measured_ones():
    """Each pin names a copied fixture, so that a renamed one cannot leave a pin behind."""
    assert {name for name, _view in DIFFERS} <= {path.stem for path in COPIED}
    assert len(DIFFERS) == 10


def _docx4j():
    for row in DOCX4J:
        for view in VIEWS:
            marks = [pytest.mark.slow]
            if row[view] == "differs":
                marks.append(pytest.mark.xfail(strict=True, reason="CR-011: measured 2026-10-21"))
            yield pytest.param(row["fixture"], view, marks=marks, id=f"{row['fixture']}-{view}")


@pytest.mark.parametrize(("fixture", "view"), list(_docx4j()))
def test_docx4js_fixtures_read_as_their_resolved_copies(fixture, view):
    path = SIBLING / fixture
    if not path.is_file():
        pytest.skip(f"docx4j's {fixture} is not checked out at {SIBLING}")
    _assert_agree(path, view)


def test_the_docx4j_table_is_the_measured_one():
    """147 fixtures at 198d015a3: 80 differ in accepted and 81 in original today."""
    assert len(DOCX4J) == 147
    assert [sum(row[view] == "differs" for row in DOCX4J) for view in VIEWS] == [80, 81]
    assert {row[view] for row in DOCX4J for view in VIEWS} == {"agrees", "differs"}
