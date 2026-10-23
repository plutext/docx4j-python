"""The views against the resolved copy (REQ-406). CR-011 sections 3.2, 3.6 and 12.

docx4j's ``ViewAgreementTest`` at ``198d015a3``, over the text model: the walk
is :func:`~docx4j_py.model.text.extract` on the fixture as loaded, in the
``accepted`` or ``original`` view, which asks the views' plan
(:class:`~docx4j_py.model.text.view.ResolvedView`) what Accept All or Reject
All would leave; the reference is the same fixture loaded a second time,
resolved by CR-010's ``revisions.accept_all`` or ``reject_all``, and read by
:func:`~docx4j_py.model.text.extract_resolved`, with no view logic. Both are
written by :func:`~docx4j_py.model.text.plain_text`. The copy is the
reference: where the two differ, the walk is wrong.

docx4j loads its copy a second time, and so does this test: a ``dry_run``
trial keeps a comment the resolution empties, which the walk drops, as Word
does (``review-for-word``'s comment 0, rejected), because the engine's flush
does not take a trial's copy of a part for a part (CR-011 section 12).

Two scopes, as docx4j's test has them: the default, and every story with
hidden text and field codes (``include: headers, footers``, ``comments``,
``hidden``, ``field_codes``).

Over two sets:

* CR-010's eleven agreement fixtures and the probe, which are in this
  repository, and since CR-010 section 13 check 54's three Word saves and
  docx4j's text-left probe with Word's two (docx4j ``29dbf5964`` and
  ``e55a83b8e``, whose ``ViewAgreementTest`` walks them);
* docx4j's 147 revisions and text-extraction fixtures at ``198d015a3``,
  where the sibling checkout has them (``slow``), as listed in
  ``agreement-docx4j.tsv`` with their outcomes. A file docx4j adds later is not
  in the table, and is not read.

Before Phase B the walk was ``get_text`` (CR-011 section 11.4): 10 of the
copied 24 differed, and 80 and 81 of docx4j's 147. None differs now.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from conformance import PHASE0, ROOT, table

from docx4j_py import load
from docx4j_py.model import revisions
from docx4j_py.model.text import extract, extract_resolved, plain_text

#: docx4j-core-tests' resources, in the sibling checkout.
SIBLING = ROOT.parent / "docx4j" / "docx4j-core-tests" / "src" / "test" / "resources"

#: The agreement run's eleven Word-made fixtures (CR-010 section 3.6).
REVISIONS = ROOT / "tests" / "fixtures" / "revisions"
#: Office JS's checks' files and Word's saves of them (CR-013 Phase C).
OFFICE_JS = ROOT / "tests" / "office_js"
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
    # check 54 (CR-010 section 13): Word's save after the writes, and Word's own Accept All of it, as docx4j's
    # ViewAgreementTest walks them since 29dbf5964
    OFFICE_JS / "54-writes-word365.docx",
    OFFICE_JS / "54-writes-word365-accepted.docx",
    # Word's Reject All of it, and docx4j's text-left probe with Word's two saves (e55a83b8e, CR-010 section 13.3a)
    OFFICE_JS / "54-writes-word365-rejected.docx",
    *(REVISIONS / "cr043" / "s17" / name for name in ("17-text-left.docx", "17-text-left-accepted.docx", "17-text-left-rejected.docx")),
]

VIEWS = {"accepted": revisions.accept_all, "original": revisions.reject_all}

#: docx4j's two scopes: the default, and every story with hidden text and field codes.
SCOPES = {
    "default": {},
    "everything": {"include": "headers, footers", "comments": True, "hidden": True, "field_codes": True},
}

#: docx4j's fixtures at 198d015a3, under docx4j-core-tests' resources, with
#: today's outcome in each view.
DOCX4J = table(Path(__file__).resolve().parent / "agreement-docx4j.tsv")


def _walk_and_copy(path: Path, view: str, scope: str = "default") -> tuple[str, str]:
    """The walk's text in a view, and the text of a copy resolved to that view."""
    walk = plain_text(extract(load(path), view=view, **SCOPES[scope]))
    copy = load(path)
    VIEWS[view](copy)
    return walk, plain_text(extract_resolved(copy, view=view, **SCOPES[scope]))


def _first_difference(walk: str, copy: str) -> str:
    """docx4j's ``firstDifference``: the first line that differs, both ways."""
    w, c = walk.split("\n"), copy.split("\n")
    i = 0
    while i < min(len(w), len(c)) and w[i] == c[i]:
        i += 1
    return f"line {i + 1}\n  walk: {w[i] if i < len(w) else '<end>'}\n  copy: {c[i] if i < len(c) else '<end>'}"


def _assert_agree(path: Path, view: str, scope: str) -> None:
    walk, copy = _walk_and_copy(path, view, scope)
    assert walk == copy, f"{path.name} {view} {scope}: {_first_difference(walk, copy)}"


def _copied():
    for path in COPIED:
        for view in VIEWS:
            for scope in SCOPES:
                yield pytest.param(path, view, scope, id=f"{path.stem}-{view}-{scope}")


@pytest.mark.parametrize(("path", "view", "scope"), list(_copied()))
def test_the_view_reads_as_the_resolved_copy(path, view, scope):
    _assert_agree(path, view, scope)


def test_the_walk_does_not_change_the_package():
    """REQ-204: a reading in every view leaves the probe unread, so it saves byte for byte."""
    path = PHASE0 / "extraction-probe.docx"
    package = load(path)
    before = package.save()
    for view in ("accepted", "original", "markup"):
        for scope in SCOPES:
            extract(package, view=view, **SCOPES[scope])
    assert not package.main_document_part.is_unmarshalled
    assert package.save() == before


def _docx4j():
    for row in DOCX4J:
        for view in VIEWS:
            for scope in SCOPES:
                marks = [pytest.mark.slow]
                if row[view] == "differs":
                    marks.append(pytest.mark.xfail(strict=True, reason="CR-011: measured"))
                yield pytest.param(row["fixture"], view, scope, marks=marks, id=f"{row['fixture']}-{view}-{scope}")


@pytest.mark.parametrize(("fixture", "view", "scope"), list(_docx4j()))
def test_docx4js_fixtures_read_as_their_resolved_copies(fixture, view, scope):
    path = SIBLING / fixture
    if not path.is_file():
        pytest.skip(f"docx4j's {fixture} is not checked out at {SIBLING}")
    _assert_agree(path, view, scope)


def test_the_docx4j_table_is_the_measured_one():
    """147 fixtures at 198d015a3, every one agreeing in both views since Phase B (80 and 81 differed before)."""
    assert len(DOCX4J) == 147
    assert [sum(row[view] == "differs" for row in DOCX4J) for view in VIEWS] == [0, 0]
    assert {row[view] for row in DOCX4J for view in VIEWS} == {"agrees"}
