"""A paragraph appended with tracking on, in Word's form (check 29). CR-010 section 12.

Word's Enter at the end of a paragraph ends that paragraph with the **new**
mark --- the inserted one, by the typist --- and leaves the old mark, with the
properties it carried, on the new line. After a Heading 1, whose next style is
Normal, the new line turns Normal and Word records that in a ``w:pPrChange`` on
it holding Heading 1 (docx4j-core-ts check 29, ``29a-heading-end-tracked``);
after a Normal paragraph there is nothing to record and it writes none
(``29b-normal-end-tracked``). The tracker's appended paragraph at the end of a
container is the same shape (CR-003 section 16.10), and since CR-010 Phase C it
writes the record too, so that rejecting an append after a heading gives the
heading back (CR-010 section 10.6).

The fixtures are docx4j-core-ts's (``tests/README.md`` has the provenance).
Word's files are the oracle twice: the markup this package writes is described
as Word's is, and resolving it gives what resolving Word's does.
"""

from __future__ import annotations

import re

import pytest
from description import describe
from revision_helpers import fixture, paragraphs, revisions, round_trip, text_of

from docx4j_py import WordprocessingMLPackage, create_package
from docx4j_py.model.content import Author
from docx4j_py.wml import P

AUTHOR = "Author B"
HEADING = "Case a heading"
NORMAL = "Case b: the end of this Normal paragraph."


def _cut_after(package, text):
    """Everything in the body after the paragraph reading `text` removed, untracked."""
    content = package.main_document_part.contents.body.content
    index = next(i for i, item in enumerate(content) if isinstance(item, P) and text_of(item) == text)
    del content[index + 1 :]
    return package


def _input(last):
    """29-input with `last` as the body's last paragraph, and Author B typing with tracking on."""
    package = _cut_after(fixture("core-ts/check29/29-input.docx"), last)
    package.author = Author(AUTHOR)
    package.change_tracking_mode = "TrackAll"
    return package


def _word(name):
    """Word's save of the case, cut after the paragraph Word typed (``New``)."""
    return _cut_after(fixture(f"core-ts/check29/{name}"), "New")


def _style(p):
    style = p.p_pr.p_style if p.p_pr is not None else None
    return style.val if style is not None else "Normal"


def _styles(package):
    """Each body paragraph's style id and text."""
    return [(_style(p), text_of(p)) for p in paragraphs(package)]


def _resolved(package, accept):
    copy = WordprocessingMLPackage.load(package.save())
    (revisions.accept_all if accept else revisions.reject_all)(copy)
    assert revisions.tracked_changes(copy) == []
    return copy


def _p_pr_xml(package, index=-1):
    """A paragraph's ``w:pPr`` with ids and dates left out: Word's ``w16du:dateUtc`` too (CR-010 10.3 item 1)."""
    p = paragraphs(package)[index]
    from docx4j_py.fragments import to_xml

    xml = to_xml(p.p_pr) if p.p_pr is not None else ""
    return re.sub(r' (w:date|w:id|w16du:dateUtc)="[^"]*"', "", re.sub(r" xmlns:\w+=\"[^\"]*\"", "", xml))


# ---------------------------------------------------------------------------
# the markup, beside Word's
# ---------------------------------------------------------------------------


def test_an_append_after_a_heading_records_the_heading_as_29a_does():
    package = _input(HEADING)
    package.body.insert_paragraph("New")

    last, heading = paragraphs(package)[-1], paragraphs(package)[-2]
    assert heading.p_pr.p_style.val == "Heading1"
    assert heading.p_pr.r_pr.ins.author == AUTHOR, "the heading ends with the new, inserted mark"
    assert last.p_pr.p_style is None and last.p_pr.r_pr is None, "the new line keeps the old mark, Normal"
    change = last.p_pr.p_pr_change
    assert change is not None and change.author == AUTHOR
    assert _p_pr_xml(package) == (
        f'<w:pPr><w:pPrChange w:author="{AUTHOR}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        "</w:pPrChange></w:pPr>"
    )
    word = _word("29a-heading-end-tracked.docx")
    assert _p_pr_xml(word) == _p_pr_xml(package), "Word's 29a writes the same w:pPr"
    assert describe(round_trip(package)) == describe(word)


def test_an_append_after_a_normal_paragraph_records_nothing_as_29b():
    package = _input(NORMAL)
    package.body.insert_paragraph("New")

    last = paragraphs(package)[-1]
    assert last.p_pr is None, "nothing to record: the new line's properties are the old mark's"
    assert paragraphs(package)[-2].p_pr.r_pr.ins.author == AUTHOR
    word = _word("29b-normal-end-tracked.docx")
    assert "pPrChange" not in describe(word)
    assert describe(round_trip(package)) == describe(word)


@pytest.mark.parametrize("accept", [False, True], ids=["reject", "accept"])
@pytest.mark.parametrize(
    "case", [("29a-heading-end-tracked.docx", HEADING), ("29b-normal-end-tracked.docx", NORMAL)], ids=["29a", "29b"]
)
def test_resolving_the_append_resolves_as_words_markup_does(case, accept):
    name, last = case
    package = _input(last)
    package.body.insert_paragraph("New")
    ours, word = _resolved(package, accept), _resolved(_word(name), accept)
    assert describe(ours) == describe(word)


def test_rejecting_an_append_after_a_heading_gives_one_heading_back():
    package = _input(HEADING)
    package.body.insert_paragraph("New")
    rejected = _resolved(package, False)
    assert _styles(rejected)[-1] == ("Heading1", HEADING)
    assert [text for _style, text in _styles(rejected)].count(HEADING) == 1
    assert "New" not in [text for _style, text in _styles(rejected)]
    assert "pPrChange" not in rejected.main_document_part.get_xml()


def test_accepting_an_append_after_a_heading_keeps_both_with_their_own_properties():
    package = _input(HEADING)
    package.body.insert_paragraph("New")
    accepted = _resolved(package, True)
    assert _styles(accepted)[-2:] == [("Heading1", HEADING), ("Normal", "New")]


# ---------------------------------------------------------------------------
# every route that appends
# ---------------------------------------------------------------------------


def _created(last_style):
    """A created document whose last paragraph is in `last_style`, tracking on."""
    package = create_package()
    package.body.insert_paragraph("Intro.")
    package.body.insert_paragraph("Last", style=last_style)
    package.author = Author(AUTHOR)
    package.change_tracking_mode = "TrackAll"
    return package


def _record(package):
    """What the last paragraph's ``w:pPrChange`` records, or None."""
    p_pr = paragraphs(package)[-1].p_pr
    change = p_pr.p_pr_change if p_pr is not None else None
    if change is None:
        return None
    style = change.p_pr.p_style if change.p_pr is not None else None
    return style.val if style is not None else "Normal"


ROUTES = {
    "Body.insert_paragraph": lambda pkg: pkg.body.insert_paragraph("New"),
    "Paragraph.insert_paragraph(style=Normal)": lambda pkg: pkg.body.paragraphs[-1].insert_paragraph(
        "New", location="After", style="Normal"
    ),
    "Body.insert_markdown": lambda pkg: pkg.body.insert_markdown("New"),
    "Body.insert_xml": lambda pkg: pkg.body.insert_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>New</w:t></w:r></w:p>'
    ),
}


@pytest.mark.parametrize("route", list(ROUTES), ids=list(ROUTES))
def test_every_route_that_appends_writes_29as_record(route):
    package = _created("Heading 1")
    ROUTES[route](package)
    assert _record(package) == "Heading1"
    assert _styles(_resolved(package, False))[-1] == ("Heading1", "Last")
    assert _styles(_resolved(package, True))[-2:] == [("Heading1", "Last"), ("Normal", "New")]


def _through_the_facade(last_style):
    """The same document built through the python-docx facade, which appends as python-docx does."""
    from docx4j_py.docx import Document

    document = Document()
    document.add_paragraph("Intro.")
    document.add_paragraph("Last", style=last_style)
    document.package.author = Author(AUTHOR)
    document.package.change_tracking_mode = "TrackAll"
    return document


def test_the_facades_appends_write_29as_record():
    document = _through_the_facade("Heading 1")
    document.add_paragraph("New")
    package = document.package
    assert _record(package) == "Heading1"
    assert _styles(_resolved(package, False))[-1] == ("Heading1", "Last")
    assert _styles(_resolved(package, True))[-2:] == [("Heading1", "Last"), ("Normal", "New")]

    heading = _through_the_facade("Normal")
    heading.add_heading("New", 1)
    assert _record(heading.package) == "Normal", "add_heading sets the style after the append"
    assert _styles(_resolved(heading.package, False))[-1] == ("Normal", "Last")
    assert _styles(_resolved(heading.package, True))[-1] == ("Heading1", "New")


def test_insert_paragraph_after_the_last_paragraph_copies_its_properties_and_needs_no_record():
    """Office JS's ``insertParagraph("After")`` takes this paragraph's properties: nothing differs."""
    package = _created("Heading 1")
    package.body.paragraphs[-1].insert_paragraph("New", location="After")
    assert _record(package) is None
    assert _styles(_resolved(package, False))[-1] == ("Heading1", "Last")
    assert _styles(_resolved(package, True))[-2:] == [("Heading1", "Last"), ("Heading1", "New")]


def test_a_style_set_on_the_append_is_recorded_against_the_old_mark():
    """``insert_paragraph(style=...)`` sets the style after the insert; the record follows it."""
    same = _created("Heading 1")
    same.body.insert_paragraph("New", style="Heading 1")
    assert _record(same) is None, "back to the heading's properties: the record changes nothing and goes"

    other = _created("Normal")
    other.body.insert_paragraph("New", style="Heading 2")
    assert _record(other) == "Normal", "the old mark was Normal"
    assert _styles(_resolved(other, False))[-1] == ("Normal", "Last")
    assert _styles(_resolved(other, True))[-1] == ("Heading2", "New")


def test_a_fragment_appended_records_on_its_last_paragraph_only():
    package = _created("Heading 1")
    package.body.insert_markdown("one\n\ntwo\n\nthree")
    records = [p.p_pr is not None and p.p_pr.p_pr_change is not None for p in paragraphs(package)]
    assert records == [False, False, False, False, True]
    assert _record(package) == "Heading1"
    assert _styles(_resolved(package, False))[-1] == ("Heading1", "Last")
    assert len(paragraphs(_resolved(package, False))) == 2

    same = _created("Heading 1")
    same.body.insert_markdown("one\n\n# three")
    assert _record(same) is None, "a fragment ending in a heading has nothing to record"


def test_the_record_moves_with_the_old_mark_and_comes_back_when_the_append_is_taken_back():
    package = _created("Heading 1")
    package.body.insert_paragraph("A")
    b = package.body.insert_paragraph("B")
    records = [p.p_pr is not None and p.p_pr.p_pr_change is not None for p in paragraphs(package)]
    assert records == [False, False, False, True], "on B, which holds the old mark now; not on A too"
    assert _styles(_resolved(package, False))[-1] == ("Heading1", "Last")

    b.delete()  # this author's own insertion, taken back
    assert [text_of(p) for p in paragraphs(package)] == ["Intro.", "Last", "A"]
    assert _record(package) == "Heading1", "A holds the old mark again, and its record"
    assert _styles(_resolved(package, False))[-1] == ("Heading1", "Last")

    package.body.paragraphs[-1].delete()
    assert [text_of(p) for p in paragraphs(package)] == ["Intro.", "Last"]
    assert "w:ins" not in package.body.get_xml() and "pPrChange" not in package.body.get_xml()


def test_an_insert_in_the_middle_carries_its_own_mark_and_no_record():
    """Check 29's record belongs to the old final mark; in the middle each paragraph has its own."""
    package = _created("Heading 1")
    package.body.insert_paragraph("Middle", location="Start")
    package.body.paragraphs[1].insert_paragraph("After intro", location="After", style="Normal")
    assert "pPrChange" not in package.body.get_xml()


def test_a_cell_is_a_container_too():
    """A paragraph appended at the end of a table cell after a heading: the same record, the same reject."""
    package = create_package()
    table = package.body.insert_table(1, 1, values=[["In the cell"]])
    cell = table.cell(0, 0)
    cell.body.paragraphs[0].style = "Heading 1"
    package.body.insert_paragraph("After the table.")
    package.author = Author(AUTHOR)
    package.change_tracking_mode = "TrackAll"
    cell.insert_paragraph("New")

    in_cell = cell.body.paragraphs
    assert in_cell[-1].element.p_pr.p_pr_change.p_pr.p_style.val == "Heading1"
    assert in_cell[-2].element.p_pr.r_pr.ins.author == AUTHOR
    for accept, expected in ((False, ["Heading1"]), (True, ["Heading1", "Normal"])):
        resolved = _resolved(package, accept)
        cell_paragraphs = resolved.body.tables[0].cell(0, 0).body.paragraphs
        assert [p.style_id for p in cell_paragraphs] == expected
